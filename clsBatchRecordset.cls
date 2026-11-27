VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatchRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Class   : clsBatchRecordset
' Purpose : Edit a set of rows (for example the lines of one invoice) in memory, and write
'           them back to SQL Server in one go. The rows are read once into a disconnected
'           ADO recordset with a batch lock. Edits - from an unbound form one record at a
'           time, or in code across many records - only change that local copy until
'           Commit is called. Nothing reaches the server before then.
'
' Install : Import clsBatchRecordset.cls. Needs a reference to Microsoft ActiveX Data
'           Objects (2.8 or later). The tests are in modBatchRecordsetTests.bas
'           (optional, development database only).
'
' Usage   : Private m_Lines As clsBatchRecordset
'
'           ' Open the invoice. KeyField identifies a row in the failure message;
'           ' LabelField (optional) is what the user recognises, e.g. the line number.
'           Set m_Lines = New clsBatchRecordset
'           m_Lines.Retrieve strConnect, _
'               "SELECT * FROM Invoice_Line WHERE Invoice_ID = " & lngInvoice & _
'               " ORDER BY Line_Number", "Invoice_Line_ID", "Line_Number"
'           m_Lines.PopulateForm Me                  ' show line 1
'
'           ' Next button: keep the edits, move on
'           m_Lines.UpdateFromForm Me                ' form -> current row (local only)
'           If m_Lines.MoveNext Then m_Lines.PopulateForm Me
'
'           ' "Deny all" (no form needed)
'           m_Lines.ApplyToAll "Line_Status", "Denied"
'           m_Lines.ApplyToFiltered "Line_Status = 'Open'", "Line_Status", "Denied"
'
'           ' Done: confirm, then write
'           If m_Lines.PendingCount > 0 Then
'               If MsgBox("Save " & m_Lines.PendingCount & " changed line(s)?", _
'                         vbYesNo) = vbYes Then
'                   strMsg = m_Lines.Commit
'                   If Len(strMsg) > 0 Then MsgBox strMsg      ' some lines failed
'               Else
'                   m_Lines.Cancel                 ' throw all edits away
'               End If
'           End If
'
' Rules   : - Control-to-field mapping: the control's Tag holds the field name, the same
'             rule as PopulateUnboundControls. Text boxes, combo boxes and check boxes
'             with a Tag are used; a Tag naming no field is skipped (Debug.Print only).
'           - Commit is NOT one transaction. The rows that can be saved are saved, and the
'             ones that cannot are reported in the returned message. A row fails when
'             someone else changed it since it was read (optimistic locking compares the
'             original values) or when the server rejects it (constraint, permission).
'             The failed rows stay pending in the buffer so the user can see them; call
'             Retrieve again to discard them and re-read the server's values.
'           - Commit returns "" when everything was saved. An error that is not tied to a
'             row (for example the connection dropped) is raised, and every change stays
'             pending so Commit can simply be called again.
'
' Notes   : - The recordset must come from ONE table, with a primary key. UpdateBatch
'             cannot update a join, and needs the key to find each row.
'           - UpdateBatch sends one UPDATE per changed row. That is fine for the tens or
'             hundreds of lines on an invoice. For thousands of rows at once, a single
'             set-based UPDATE on the server is much faster; this class does not do that.
'           - ApplyToFiltered takes an ADO Filter string, e.g. "Line_Status = 'Open'"
'             (not SQL; strings in single quotes, AND / OR allowed, no functions).
'           - The form's own dirty flag is not part of this class. Use clsDirtyTracker to
'             know whether the user edited the CURRENT record before moving on.
'           - Attach is for a recordset that is already built (tests, or one made with
'             Fields.Append). Such a class cannot Commit - there is no connection string.
'           - Written but not run against a server. The first run should include a
'             deliberate conflict: change a line in SSMS while it is open here and
'             check that Commit reports it and saves the others.
'
' History : 11/25/26 Created.
'           11/27/26 Release renamed Dispose (Release is reserved). Attach now rejects
'           a recordset it cannot work with.
'---------------------------------------------------------------------------------------

Private Const ERR_NOT_OPEN As Long = vbObjectError + 517
Private Const ERR_NO_ROW As Long = vbObjectError + 518
Private Const ERR_NO_FIELD As Long = vbObjectError + 519
Private Const ERR_NO_CONNECTION As Long = vbObjectError + 520
Private Const ERR_BAD_RECORDSET As Long = vbObjectError + 521

Private m_Rst As ADODB.Recordset
Private m_ConnString As String
Private m_Sql As String
Private m_KeyField As String
Private m_LabelField As String
Private m_FailedCount As Long

Private Sub Class_Terminate()

    Dispose

End Sub

'---------------------------------------------------------------------------------------
' Getting the rows
'---------------------------------------------------------------------------------------

' Reads the rows into a disconnected recordset and closes the connection. Any earlier
' contents, including pending edits, are discarded. Raises an error if the query or
' connection fails; the class is then empty.
Public Sub Retrieve(ByVal ConnectionString As String, ByVal SourceSql As String, _
                    ByVal KeyField As String, Optional ByVal LabelField As String = "")

    Dim cn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim lngErr As Long
    Dim strSrc As String
    Dim strDesc As String

    Dispose

    On Error GoTo Fail

    Set cn = New ADODB.Connection
    cn.Open ConnectionString

    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open SourceSql, cn, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rst.ActiveConnection = Nothing

    cn.Close

    Set m_Rst = rst
    m_ConnString = ConnectionString
    m_Sql = SourceSql
    m_KeyField = KeyField
    m_LabelField = LabelField

    Exit Sub

Fail:
    lngErr = Err.Number
    strSrc = Err.Source
    strDesc = Err.Description

    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strDesc

End Sub

' Throws away every edit and reads the same rows again from the server.
Public Sub Reload()

    RequireOpen
    If Len(m_ConnString) = 0 Then
        Err.Raise ERR_NO_CONNECTION, "clsBatchRecordset.Reload", _
                  "This recordset was attached, not retrieved, so it cannot be reloaded."
    End If

    Retrieve m_ConnString, m_Sql, m_KeyField, m_LabelField

End Sub

' Uses a recordset that is already built. It must be client-side with a batch lock
' (CursorLocation = adUseClient, LockType = adLockBatchOptimistic) and open. Anything
' else raises an error and leaves the class as it was. (A forward-only server cursor
' reports RecordCount = -1, so its rows would otherwise look like an empty set.)
Public Sub Attach(ByVal Rst As ADODB.Recordset, ByVal KeyField As String, _
                  Optional ByVal LabelField As String = "")

    Dim strProblem As String

    ' ElseIf, because VBA evaluates both sides of Or: each test needs the one above.
    If Rst Is Nothing Then
        strProblem = "is Nothing"
    ElseIf Rst.State <> adStateOpen Then
        strProblem = "is not open"
    ElseIf Rst.CursorLocation <> adUseClient Then
        strProblem = "is not client-side (CursorLocation = adUseClient)"
    ElseIf Rst.LockType <> adLockBatchOptimistic Then
        strProblem = "does not have a batch lock (LockType = adLockBatchOptimistic)"
    End If

    If Len(strProblem) > 0 Then
        Err.Raise ERR_BAD_RECORDSET, "clsBatchRecordset.Attach", _
                  "The recordset " & strProblem & "."
    End If

    Dispose

    Set m_Rst = Rst
    m_ConnString = vbNullString
    m_Sql = vbNullString
    m_KeyField = KeyField
    m_LabelField = LabelField

End Sub

' Closes the recordset and forgets everything. Called by Class_Terminate.
Public Sub Dispose()

    On Error Resume Next
    If Not m_Rst Is Nothing Then
        If m_Rst.State <> adStateClosed Then m_Rst.Close
    End If
    Set m_Rst = Nothing
    On Error GoTo 0

End Sub

' The underlying recordset, for anything this class does not cover.
Public Property Get Recordset() As ADODB.Recordset

    Set Recordset = m_Rst

End Property

'---------------------------------------------------------------------------------------
' Moving between rows (these never leave the first or last row)
'---------------------------------------------------------------------------------------

Public Property Get HasRows() As Boolean

    If m_Rst Is Nothing Then Exit Property
    If m_Rst.State <> adStateOpen Then Exit Property

    HasRows = (m_Rst.RecordCount > 0)

End Property

Public Property Get Count() As Long

    If HasRows Then Count = m_Rst.RecordCount

End Property

' 1-based position of the current row; 0 if there are no rows.
Public Property Get Position() As Long

    If HasRows Then Position = m_Rst.AbsolutePosition

End Property

Public Property Get IsFirst() As Boolean

    IsFirst = (Position = 1)

End Property

Public Property Get IsLast() As Boolean

    IsLast = (HasRows And Position = Count)

End Property

' Each Move returns True if the current row changed, False if there was nowhere to go.
Public Function MoveFirst() As Boolean

    If Not HasRows Then Exit Function
    If IsFirst Then Exit Function

    m_Rst.MoveFirst
    MoveFirst = True

End Function

Public Function MoveLast() As Boolean

    If Not HasRows Then Exit Function
    If IsLast Then Exit Function

    m_Rst.MoveLast
    MoveLast = True

End Function

Public Function MoveNext() As Boolean

    If Not HasRows Then Exit Function
    If IsLast Then Exit Function

    m_Rst.MoveNext
    MoveNext = True

End Function

Public Function MovePrevious() As Boolean

    If Not HasRows Then Exit Function
    If IsFirst Then Exit Function

    m_Rst.MovePrevious
    MovePrevious = True

End Function

'---------------------------------------------------------------------------------------
' Form <-> current row
'---------------------------------------------------------------------------------------

' Copies the current row into the form's controls (control Tag = field name).
Public Sub PopulateForm(ByVal frm As Access.Form)

    Dim ctl As Access.Control
    Dim fld As ADODB.Field

    RequireCurrentRow "PopulateForm"

    For Each ctl In frm.Controls

        If IsDataControl(ctl) Then

            Set fld = FindField(ctl.Tag)

            If fld Is Nothing Then
                Debug.Print "clsBatchRecordset: field not found: " & ctl.Tag & _
                            " (" & ctl.Name & ")"
            Else
                ctl.Value = fld.Value
            End If

        End If

    Next ctl

End Sub

' Copies the form's controls into the current row, only where the value differs, and
' updates the local row. Returns how many fields changed. Nothing goes to the server.
' A value the field cannot hold (text in a number field) raises the ADO error.
Public Function UpdateFromForm(ByVal frm As Access.Form) As Long

    Dim ctl As Access.Control
    Dim fld As ADODB.Field
    Dim lngChanged As Long

    RequireCurrentRow "UpdateFromForm"

    For Each ctl In frm.Controls

        If IsDataControl(ctl) Then

            Set fld = FindField(ctl.Tag)

            If Not fld Is Nothing Then
                If ValuesDiffer(fld.Value, ctl.Value) Then
                    fld.Value = ctl.Value
                    lngChanged = lngChanged + 1
                End If
            End If

        End If

    Next ctl

    If lngChanged > 0 Then m_Rst.Update

    UpdateFromForm = lngChanged

End Function

'---------------------------------------------------------------------------------------
' Changing many rows without a form
'---------------------------------------------------------------------------------------

' Sets one field on every row. Returns how many rows actually changed (rows that
' already held the value are left alone). The current row does not move.
Public Function ApplyToAll(ByVal FieldName As String, ByVal NewValue As Variant) As Long

    ApplyToAll = ApplyToFiltered(vbNullString, FieldName, NewValue)

End Function

' Sets one field on every row that matches an ADO Filter string, for example
' "Line_Status = 'Open'". Returns how many rows changed. The current row does not move.
' The matching rows are noted first and changed afterwards, so setting the field the
' criteria test (Open -> Denied) cannot make the loop skip rows.
Public Function ApplyToFiltered(ByVal Criteria As String, ByVal FieldName As String, _
                                ByVal NewValue As Variant) As Long

    Dim bmHome As Variant
    Dim marks As Collection
    Dim bm As Variant
    Dim fld As ADODB.Field
    Dim lngChanged As Long
    Dim lngErr As Long
    Dim strSrc As String
    Dim strDesc As String

    RequireOpen

    If FindField(FieldName) Is Nothing Then
        Err.Raise ERR_NO_FIELD, "clsBatchRecordset.ApplyToFiltered", _
                  "No field named " & FieldName
    End If

    If Not HasRows Then Exit Function

    On Error GoTo Fail

    bmHome = SavePosition()
    Set marks = New Collection

    If Len(Criteria) > 0 Then
        m_Rst.Filter = Criteria
    Else
        m_Rst.Filter = adFilterNone
    End If

    Do While Not m_Rst.EOF
        marks.Add m_Rst.Bookmark
        m_Rst.MoveNext
    Loop

    m_Rst.Filter = adFilterNone

    For Each bm In marks
        m_Rst.Bookmark = bm
        Set fld = m_Rst.Fields(FieldName)
        If ValuesDiffer(fld.Value, NewValue) Then
            fld.Value = NewValue
            m_Rst.Update
            lngChanged = lngChanged + 1
        End If
    Next bm

    RestorePosition bmHome

    ApplyToFiltered = lngChanged

    Exit Function

Fail:
    lngErr = Err.Number
    strSrc = Err.Source
    strDesc = Err.Description

    On Error Resume Next
    m_Rst.CancelUpdate
    m_Rst.Filter = adFilterNone
    RestorePosition bmHome
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strDesc

End Function

'---------------------------------------------------------------------------------------
' Saving
'---------------------------------------------------------------------------------------

' How many rows have edits waiting to be written. The current row does not move.
Public Function PendingCount() As Long

    Dim bmHome As Variant

    If Not HasRows Then Exit Function

    bmHome = SavePosition()

    m_Rst.Filter = adFilterPendingRecords
    PendingCount = m_Rst.RecordCount
    m_Rst.Filter = adFilterNone

    RestorePosition bmHome

End Function

' How many rows the last Commit could not save.
Public Property Get FailedCount() As Long

    FailedCount = m_FailedCount

End Property

' Throws away every edit made since the rows were read (or since the last Commit).
Public Sub Cancel()

    RequireOpen

    m_Rst.CancelBatch adAffectAll

End Sub

' Writes the edits to the server. Rows that save are saved even if others fail.
' Returns "" if everything saved; otherwise a message listing each row that did not,
' and why. The failed rows stay pending. Raises an error for a problem that is not
' about a particular row (no connection, for instance); all edits stay pending then.
Public Function Commit() As String

    Dim cn As ADODB.Connection
    Dim bmHome As Variant
    Dim lngErr As Long
    Dim strErr As String
    Dim strFailed As String
    Dim strLabel As String

    RequireOpen

    If Len(m_ConnString) = 0 Then
        Err.Raise ERR_NO_CONNECTION, "clsBatchRecordset.Commit", _
                  "This recordset was attached, not retrieved, so there is no connection to save to."
    End If

    m_FailedCount = 0
    bmHome = SavePosition()

    On Error GoTo Fail

    Set cn = New ADODB.Connection
    cn.Open m_ConnString
    Set m_Rst.ActiveConnection = cn

    ' UpdateBatch tries every pending row, then raises one error if any failed.
    ' That error is read, not acted on, until the failed rows have been collected.
    On Error Resume Next
    m_Rst.UpdateBatch adAffectAll
    lngErr = Err.Number
    strErr = Err.Description
    Err.Clear
    On Error GoTo Fail

    m_Rst.Filter = adFilterConflictingRecords

    Do While Not m_Rst.EOF

        m_FailedCount = m_FailedCount + 1

        If Len(m_LabelField) > 0 Then
            strLabel = m_LabelField & " " & m_Rst.Fields(m_LabelField).Value
        Else
            strLabel = m_KeyField & " " & m_Rst.Fields(m_KeyField).Value
        End If

        strFailed = strFailed & vbCrLf & "  " & strLabel & " - " & _
                    StatusReason(m_Rst.Status)

        m_Rst.MoveNext

    Loop

    m_Rst.Filter = adFilterNone

    Disconnect cn
    RestorePosition bmHome

    If lngErr <> 0 And m_FailedCount = 0 Then
        Err.Raise lngErr, "clsBatchRecordset.Commit", strErr
    End If

    If m_FailedCount > 0 Then
        Commit = m_FailedCount & " line(s) were not saved:" & strFailed
    End If

    Exit Function

Fail:
    lngErr = Err.Number
    strErr = Err.Description

    On Error Resume Next
    m_Rst.Filter = adFilterNone
    Disconnect cn
    RestorePosition bmHome
    On Error GoTo 0

    Err.Raise lngErr, "clsBatchRecordset.Commit", strErr

End Function

'---------------------------------------------------------------------------------------
' Helpers
'---------------------------------------------------------------------------------------

Private Sub Disconnect(ByRef cn As ADODB.Connection)

    On Error Resume Next

    If Not m_Rst Is Nothing Then Set m_Rst.ActiveConnection = Nothing

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If

    Set cn = Nothing

End Sub

Private Sub RequireOpen()

    If Not HasRows Then
        If m_Rst Is Nothing Then
            Err.Raise ERR_NOT_OPEN, "clsBatchRecordset", _
                      "No recordset. Call Retrieve or Attach first."
        End If
        If m_Rst.State <> adStateOpen Then
            Err.Raise ERR_NOT_OPEN, "clsBatchRecordset", "The recordset is closed."
        End If
    End If

End Sub

Private Sub RequireCurrentRow(ByVal Caller As String)

    RequireOpen

    If Not HasRows Then
        Err.Raise ERR_NO_ROW, "clsBatchRecordset." & Caller, "There are no rows."
    End If

End Sub

' A bookmark for the current row, or Empty if there is none.
Private Function SavePosition() As Variant

    SavePosition = Empty

    If Not HasRows Then Exit Function
    If m_Rst.BOF Or m_Rst.EOF Then Exit Function

    SavePosition = m_Rst.Bookmark

End Function

Private Sub RestorePosition(ByVal Bookmark As Variant)

    If IsEmpty(Bookmark) Then Exit Sub

    m_Rst.Bookmark = Bookmark

End Sub

Private Function FindField(ByVal FieldName As String) As ADODB.Field

    Dim fld As ADODB.Field

    If Len(FieldName) = 0 Then Exit Function

    For Each fld In m_Rst.Fields
        If StrComp(fld.Name, FieldName, vbTextCompare) = 0 Then
            Set FindField = fld
            Exit Function
        End If
    Next fld

End Function

' A control that carries a field name in its Tag. Calculated controls are skipped.
Private Function IsDataControl(ByVal ctl As Access.Control) As Boolean

    Select Case ctl.ControlType
        Case acTextBox, acComboBox, acCheckBox
            If Len(ctl.Tag) > 0 Then
                IsDataControl = (Left$(Nz(ctl.ControlSource, vbNullString), 1) <> "=")
            End If
    End Select

End Function

' True if NewValue is not what the field already holds. Compared as the field's own
' type, so a text box showing "5.00" does not look like a change to a currency 5.
Private Function ValuesDiffer(ByVal CurrentValue As Variant, _
                              ByVal NewValue As Variant) As Boolean

    If IsNull(CurrentValue) And IsNull(NewValue) Then Exit Function

    If IsNull(CurrentValue) Or IsNull(NewValue) Then
        ValuesDiffer = True
        Exit Function
    End If

    Select Case VarType(CurrentValue)

        Case vbDate
            If IsDate(NewValue) Then
                ValuesDiffer = (CDate(CurrentValue) <> CDate(NewValue))
            Else
                ValuesDiffer = True
            End If

        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            If IsNumeric(NewValue) Then
                ValuesDiffer = (CDec(CurrentValue) <> CDec(NewValue))
            Else
                ValuesDiffer = True
            End If

        Case vbBoolean
            ValuesDiffer = (CBool(CurrentValue) <> CBool(NewValue))

        Case Else
            ValuesDiffer = (StrComp(CStr(CurrentValue), CStr(NewValue), vbBinaryCompare) <> 0)

    End Select

End Function

' Plain-language reason for a row that UpdateBatch did not save.
Private Function StatusReason(ByVal Status As Long) As String

    If (Status And adRecConcurrencyViolation) <> 0 Then
        StatusReason = "changed by someone else since it was loaded"
    ElseIf (Status And adRecIntegrityViolation) <> 0 Then
        StatusReason = "rejected by the database (constraint violation)"
    ElseIf (Status And adRecPermissionDenied) <> 0 Then
        StatusReason = "permission denied"
    ElseIf (Status And adRecSchemaViolation) <> 0 Then
        StatusReason = "does not match the table definition"
    ElseIf (Status And adRecDBDeleted) <> 0 Then
        StatusReason = "no longer exists in the database"
    Else
        StatusReason = "rejected (ADO status " & Status & ")"
    End If

End Function
