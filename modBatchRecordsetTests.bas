Attribute VB_Name = "modBatchRecordsetTests"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : modBatchRecordsetTests
' Purpose : Tests for clsBatchRecordset that need no database server: they run on an
'           in-memory recordset passed to Attach. Results go to the Immediate window as
'           PASS / FAIL lines with a summary at the end.
'
' Install : Import into a DEVELOPMENT copy of a database, next to clsBatchRecordset.cls.
'           Do not ship it.
'
' Usage   : In the Immediate window:   RunBatchRecordsetTests
'
' Not covered : Retrieve, Reload and Commit need a SQL Server table, and the form
'           methods need a form. Test Commit by hand with a deliberate conflict:
'           open lines in the class, change one line in SSMS, edit that line and
'           another here, call Commit - expect one failure message and one saved line.
'
' Notes   : - Not run when written. Expect to fix small things the first time.
'---------------------------------------------------------------------------------------

Private m_Pass As Long
Private m_Fail As Long

Public Sub RunBatchRecordsetTests()

    m_Pass = 0
    m_Fail = 0

    Debug.Print String$(72, "-")
    Debug.Print "clsBatchRecordset tests   " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(72, "-")

    TestNavigation
    TestApplyToAll
    TestApplyToFiltered
    TestCancel

    Debug.Print String$(72, "-")
    Debug.Print "Passed: " & m_Pass & "    Failed: " & m_Fail
    Debug.Print String$(72, "-")

End Sub

' Five lines, ids 1-5, all "Open", amounts 10..50. Saved locally so none are pending.
Private Function BuildLines() As ADODB.Recordset

    Dim rst As ADODB.Recordset
    Dim i As Long

    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient

    rst.Fields.Append "Line_ID", adInteger
    rst.Fields.Append "Line_Status", adVarChar, 20, adFldIsNullable
    rst.Fields.Append "Amount", adCurrency

    rst.LockType = adLockBatchOptimistic
    rst.Open

    For i = 1 To 5
        rst.AddNew
        rst.Fields("Line_ID").Value = i
        rst.Fields("Line_Status").Value = "Open"
        rst.Fields("Amount").Value = i * 10
        rst.Update
    Next i

    rst.UpdateBatch
    rst.MoveFirst

    Set BuildLines = rst

End Function

Private Function NewLines() As clsBatchRecordset

    Dim lines As clsBatchRecordset

    Set lines = New clsBatchRecordset
    lines.Attach BuildLines(), "Line_ID"

    Set NewLines = lines

End Function

Private Sub TestNavigation()

    Dim lines As clsBatchRecordset

    Set lines = NewLines()

    Check "Count is 5", lines.Count = 5
    Check "starts at position 1", lines.Position = 1
    Check "IsFirst at start", lines.IsFirst
    Check "MovePrevious at first returns False", Not lines.MovePrevious
    Check "MoveNext moves", lines.MoveNext And lines.Position = 2
    Check "MoveLast moves to 5", lines.MoveLast And lines.Position = 5
    Check "IsLast at end", lines.IsLast
    Check "MoveNext at last returns False", Not lines.MoveNext
    Check "position stays at 5", lines.Position = 5
    Check "MoveFirst returns to 1", lines.MoveFirst And lines.Position = 1

End Sub

Private Sub TestApplyToAll()

    Dim lines As clsBatchRecordset
    Dim lngChanged As Long

    Set lines = NewLines()
    lines.MoveNext
    lines.MoveNext                       ' on line 3

    Check "nothing pending at start", lines.PendingCount = 0

    lngChanged = lines.ApplyToAll("Line_Status", "Denied")

    Check "ApplyToAll changed 5 rows", lngChanged = 5
    Check "5 rows pending", lines.PendingCount = 5
    Check "position unchanged (3)", lines.Position = 3
    Check "value is set", lines.Recordset.Fields("Line_Status").Value = "Denied"
    Check "second ApplyToAll changes 0 rows", lines.ApplyToAll("Line_Status", "Denied") = 0

End Sub

Private Sub TestApplyToFiltered()

    Dim lines As clsBatchRecordset
    Dim lngChanged As Long

    Set lines = NewLines()

    ' Make lines 2 and 4 "Approved", so only 1, 3 and 5 are still "Open".
    lines.MoveNext
    lines.Recordset.Fields("Line_Status").Value = "Approved"
    lines.Recordset.Update
    lines.MoveNext
    lines.MoveNext
    lines.Recordset.Fields("Line_Status").Value = "Approved"
    lines.Recordset.Update
    lines.MoveFirst

    ' The criteria field is the field being changed: the case that can skip rows.
    lngChanged = lines.ApplyToFiltered("Line_Status = 'Open'", "Line_Status", "Denied")

    Check "filtered apply changed 3 rows", lngChanged = 3
    Check "5 rows pending (2 approved + 3 denied)", lines.PendingCount = 5
    Check "no rows left Open", CountWhere(lines, "Line_Status = 'Open'") = 0
    Check "3 rows Denied", CountWhere(lines, "Line_Status = 'Denied'") = 3
    Check "2 rows still Approved", CountWhere(lines, "Line_Status = 'Approved'") = 2
    Check "position back at first row", lines.Position = 1
    Check "criteria matching nothing changes 0", _
          lines.ApplyToFiltered("Line_Status = 'Nope'", "Amount", 99) = 0

End Sub

Private Sub TestCancel()

    Dim lines As clsBatchRecordset

    Set lines = NewLines()

    lines.ApplyToAll "Line_Status", "Denied"
    lines.Cancel

    Check "Cancel leaves nothing pending", lines.PendingCount = 0
    Check "Cancel restores the value", _
          CountWhere(lines, "Line_Status = 'Open'") = 5

End Sub

Private Function CountWhere(ByVal lines As clsBatchRecordset, ByVal Criteria As String) As Long

    lines.Recordset.Filter = Criteria
    CountWhere = lines.Recordset.RecordCount
    lines.Recordset.Filter = adFilterNone

End Function

Private Sub Check(ByVal TestName As String, ByVal Passed As Boolean)

    If Passed Then
        m_Pass = m_Pass + 1
        Debug.Print "PASS  " & TestName
    Else
        m_Fail = m_Fail + 1
        Debug.Print "FAIL  " & TestName
    End If

End Sub
