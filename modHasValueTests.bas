Attribute VB_Name = "modHasValueTests"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : modHasValueTests
' Purpose : Tests for modHasValue (HasValue and HasRows). Results go to the Immediate
'           window as PASS / FAIL lines with a summary at the end.
'
' Install : Import into a DEVELOPMENT copy of a database, next to modHasValue.bas.
'           Do not ship it. Needs a database that allows design changes (not an ACCDE).
'
' Usage   : In the Immediate window:   RunHasValueTests
'
' Notes   : - The run creates, then deletes, a temporary form (zzHasValueTest) and a
'             temporary table (zzHasValueTestTbl) in the current database. If a run is
'             interrupted, delete those two objects by hand or just run it again.
'           - ADO is tested with an in-memory recordset (late bound), so nothing is
'             read from disk. DAO is tested against the temporary table.
'           - A FAIL line that says "raised error N" shows the error number, which is how
'             the closed-DAO-recordset error (expected 3420) gets verified.
'           - INFO lines report behavior that is not asserted, such as the starting value
'             of an unbound check box.
'           - Not run when written. Expect to fix small things the first time.
'---------------------------------------------------------------------------------------

Private Const TEST_FORM As String = "zzHasValueTest"
Private Const TEST_TABLE As String = "zzHasValueTestTbl"
Private Const ADO_INTEGER As Long = 3
Private Const DB_FAIL_ON_ERROR As Long = 128

Private m_Pass As Long
Private m_Fail As Long

Public Sub RunHasValueTests()
    m_Pass = 0
    m_Fail = 0

    Debug.Print String$(72, "-")
    Debug.Print "modHasValue tests   " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(72, "-")

    TestRecordsets
    TestControls
    TestRaises

    Debug.Print String$(72, "-")
    Debug.Print "Passed: " & m_Pass & "    Failed: " & m_Fail
    Debug.Print String$(72, "-")
End Sub

'---------------------------------------------------------------------------------------
' Recordsets
'---------------------------------------------------------------------------------------

Private Sub TestRecordsets()
    Dim db As Object
    Dim rs As Object
    Dim rsClone As Object

    On Error GoTo HandleError

    Debug.Print "HasRows"

    '--- ADO, in memory ---
    Set rs = CreateObject("ADODB.Recordset")
    CheckRows "ADO never opened (closed)", False, rs

    rs.Fields.Append "x", ADO_INTEGER
    rs.Open
    CheckRows "ADO open, no rows", False, rs

    rs.AddNew
    rs.Fields("x").Value = 1
    rs.Update
    CheckRows "ADO open, one row", True, rs

    rs.Close
    CheckRows "ADO closed after use", False, rs
    Set rs = Nothing

    '--- DAO, against a temporary table ---
    Set db = CurrentDb
    DropTestTable db
    db.Execute "CREATE TABLE " & TEST_TABLE & " (x LONG)", DB_FAIL_ON_ERROR
    db.Execute "INSERT INTO " & TEST_TABLE & " (x) VALUES (1)", DB_FAIL_ON_ERROR

    Set rs = db.OpenRecordset("SELECT * FROM " & TEST_TABLE)
    CheckRows "DAO one row", True, rs
    rs.Close
    Set rs = Nothing

    Set rs = db.OpenRecordset("SELECT * FROM " & TEST_TABLE & " WHERE x = 2")
    CheckRows "DAO no rows", False, rs
    rs.Close
    CheckRows "DAO closed (expects False, no error; an error number is shown if not)", False, rs
    Set rs = Nothing

    '--- DAO, every way of getting one. HasRows treats "Recordset2" as DAO and plain
    '--- "Recordset" as ADO, so each path must report "Recordset2" (the TypeName is shown).
    db.TableDefs.Refresh
    Set rs = db.TableDefs(TEST_TABLE).OpenRecordset()
    CheckRowsTyped "DAO from TableDef.OpenRecordset", True, rs
    rs.Close

    Set rs = db.CreateQueryDef("", "SELECT * FROM " & TEST_TABLE).OpenRecordset()
    CheckRowsTyped "DAO from an unnamed QueryDef", True, rs
    rs.Close

    Set rs = db.OpenRecordset("SELECT * FROM " & TEST_TABLE)
    Set rsClone = rs.Clone
    CheckRowsTyped "DAO from Recordset.Clone", True, rsClone
    rsClone.Close
    rs.Close
    Set rs = Nothing
    Set rsClone = Nothing

    TestFormRecordsets

CleanUp:
    On Error Resume Next
    Set rs = Nothing
    Set rsClone = Nothing
    DropTestTable db
    Exit Sub

HandleError:
    Fail "TestRecordsets stopped: error " & Err.Number & " " & Err.Description
    Resume CleanUp
End Sub

' Form.Recordset and Form.RecordsetClone, on the temporary form bound to the temporary
' table. Called from TestRecordsets while the table still exists.
Private Sub TestFormRecordsets()
    Dim frm As Form

    On Error GoTo HandleError

    RemoveTestForm
    BuildTestForm
    DoCmd.OpenForm TEST_FORM, acNormal, , , , acHidden
    Set frm = Forms(TEST_FORM)
    frm.RecordSource = "SELECT * FROM " & TEST_TABLE

    CheckRowsTyped "DAO from Form.Recordset", True, frm.Recordset
    CheckRowsTyped "DAO from Form.RecordsetClone", True, frm.RecordsetClone

CleanUp:
    On Error Resume Next
    Set frm = Nothing
    RemoveTestForm
    Exit Sub

HandleError:
    Fail "TestFormRecordsets stopped: error " & Err.Number & " " & Err.Description
    Resume CleanUp
End Sub

Private Sub DropTestTable(ByVal db As Object)
    On Error Resume Next
    db.Execute "DROP TABLE " & TEST_TABLE
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Controls: built on a temporary unbound form, opened hidden
'---------------------------------------------------------------------------------------

Private Sub TestControls()
    Dim frm As Form
    Dim blnSet As Boolean

    On Error GoTo HandleError

    Debug.Print "HasValue"

    RemoveTestForm
    BuildTestForm

    DoCmd.OpenForm TEST_FORM, acNormal, , , , acHidden
    Set frm = Forms(TEST_FORM)

    '--- TextBox ---
    CheckValue "TextBox fresh (Null)", False, frm.Controls("txtTest")
    SetVal frm, "txtTest", ""
    CheckValue "TextBox empty string", False, frm.Controls("txtTest")
    SetVal frm, "txtTest", "   "
    CheckValue "TextBox spaces only", False, frm.Controls("txtTest")
    SetVal frm, "txtTest", vbTab & vbCr & vbLf & ChrW$(160) & ChrW$(8203)
    CheckValue "TextBox tab, line breaks, non-breaking and zero-width spaces", False, frm.Controls("txtTest")
    SetVal frm, "txtTest", "x"
    CheckValue "TextBox text", True, frm.Controls("txtTest")
    SetVal frm, "txtTest", "  x  "
    CheckValue "TextBox text with padding", True, frm.Controls("txtTest")
    SetVal frm, "txtTest", 0
    CheckValue "TextBox zero is a value", True, frm.Controls("txtTest")

    '--- ComboBox, Limit To List ---
    CheckValue "ComboBox limited, fresh", False, frm.Controls("cboLimited")
    SetVal frm, "cboLimited", "b"
    CheckValue "ComboBox limited, value set in code to a listed item", True, frm.Controls("cboLimited")
    blnSet = TrySetVal(frm, "cboLimited", "zzz")
    If blnSet Then
        CheckValue "ComboBox limited, value not in the list", False, frm.Controls("cboLimited")
    Else
        Info "ComboBox limited: Access refused to set a value that is not in the list (skipped)"
    End If

    '--- ComboBox, free text ---
    CheckValue "ComboBox free, fresh", False, frm.Controls("cboFree")
    SetVal frm, "cboFree", "abc"
    CheckValue "ComboBox free, text", True, frm.Controls("cboFree")
    SetVal frm, "cboFree", "  "
    CheckValue "ComboBox free, spaces only", False, frm.Controls("cboFree")

    '--- ListBox ---
    CheckValue "ListBox single, fresh", False, frm.Controls("lstSingle")
    SetVal frm, "lstSingle", "b"
    CheckValue "ListBox single, item selected", True, frm.Controls("lstSingle")

    CheckValue "ListBox multi, fresh", False, frm.Controls("lstMulti")
    frm.Controls("lstMulti").Selected(1) = True
    CheckValue "ListBox multi, one item selected", True, frm.Controls("lstMulti")
    frm.Controls("lstMulti").Selected(1) = False
    CheckValue "ListBox multi, item deselected", False, frm.Controls("lstMulti")

    '--- CheckBox ---
    Info "CheckBox plain, fresh: Value = " & Describe(frm.Controls("chkPlain").Value) & _
         ", HasValue = " & HasValueText(frm.Controls("chkPlain"))
    Info "CheckBox tri-state, fresh: Value = " & Describe(frm.Controls("chkTri").Value) & _
         ", HasValue = " & HasValueText(frm.Controls("chkTri"))
    SetVal frm, "chkTri", Null
    CheckValue "CheckBox tri-state, Null", False, frm.Controls("chkTri")
    SetVal frm, "chkTri", False
    CheckValue "CheckBox False is a value", True, frm.Controls("chkTri")
    SetVal frm, "chkTri", True
    CheckValue "CheckBox True is a value", True, frm.Controls("chkTri")

    '--- OptionGroup ---
    CheckValue "OptionGroup fresh (no option chosen)", False, frm.Controls("fraTest")
    SetVal frm, "fraTest", 1
    CheckValue "OptionGroup option chosen", True, frm.Controls("fraTest")

    '--- ToggleButton ---
    SetVal frm, "tglTest", False
    CheckValue "ToggleButton False is a value", True, frm.Controls("tglTest")
    SetVal frm, "tglTest", True
    CheckValue "ToggleButton True is a value", True, frm.Controls("tglTest")

    '--- Label ---
    frm.Controls("lblTest").Caption = ""
    CheckValue "Label empty caption", False, frm.Controls("lblTest")
    frm.Controls("lblTest").Caption = "Hello"
    CheckValue "Label with caption", True, frm.Controls("lblTest")

CleanUp:
    On Error Resume Next
    Set frm = Nothing
    RemoveTestForm
    Exit Sub

HandleError:
    Fail "TestControls stopped: error " & Err.Number & " " & Err.Description
    Resume CleanUp
End Sub

'---------------------------------------------------------------------------------------
' Errors that must be raised
'---------------------------------------------------------------------------------------

Private Sub TestRaises()
    Dim frm As Form
    Dim objNothing As Object
    Dim rs As Object

    On Error GoTo HandleError

    Debug.Print "Errors"

    CheckValueRaises "HasValue(Nothing) raises 91", 91, objNothing
    CheckRowsRaises "HasRows(Nothing) raises 91", 91, objNothing

    Set rs = CreateObject("ADODB.Recordset")
    CheckValueRaises "HasValue(recordset) raises 5", 5, rs

    RemoveTestForm
    BuildTestForm
    DoCmd.OpenForm TEST_FORM, acNormal, , , , acHidden
    Set frm = Forms(TEST_FORM)

    CheckRowsRaises "HasRows(control) raises 5", 5, frm.Controls("txtTest")
    CheckValueRaises "HasValue(option button) raises 5", 5, frm.Controls("optA")

CleanUp:
    On Error Resume Next
    Set frm = Nothing
    Set rs = Nothing
    RemoveTestForm
    Exit Sub

HandleError:
    Fail "TestRaises stopped: error " & Err.Number & " " & Err.Description
    Resume CleanUp
End Sub

'---------------------------------------------------------------------------------------
' Temporary form
'---------------------------------------------------------------------------------------

Private Sub BuildTestForm()
    Dim frm As Form
    Dim strAuto As String
    Dim ctl As Control
    Dim lngTop As Long

    Set frm = Application.CreateForm
    strAuto = frm.Name
    lngTop = 200

    Set ctl = CreateControl(strAuto, acTextBox, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "txtTest"

    Set ctl = CreateControl(strAuto, acComboBox, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "cboLimited"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "a;b;c"
    ctl.LimitToList = True

    Set ctl = CreateControl(strAuto, acComboBox, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "cboFree"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "a;b;c"
    ctl.LimitToList = False

    Set ctl = CreateControl(strAuto, acListBox, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "lstSingle"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "a;b;c"
    ctl.MultiSelect = 0

    Set ctl = CreateControl(strAuto, acListBox, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "lstMulti"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "a;b;c"
    ctl.MultiSelect = 1

    Set ctl = CreateControl(strAuto, acCheckBox, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "chkPlain"

    Set ctl = CreateControl(strAuto, acCheckBox, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "chkTri"
    ctl.TripleState = True

    Set ctl = CreateControl(strAuto, acOptionGroup, acDetail, , , 2000, NextTop(lngTop), 1800, 700)
    ctl.Name = "fraTest"
    Set ctl = CreateControl(strAuto, acOptionButton, acDetail, "fraTest", , 2200, lngTop - 500)
    ctl.Name = "optA"
    ctl.OptionValue = 1

    Set ctl = CreateControl(strAuto, acToggleButton, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "tglTest"

    Set ctl = CreateControl(strAuto, acLabel, acDetail, , , 2000, NextTop(lngTop))
    ctl.Name = "lblTest"

    Set ctl = Nothing
    Set frm = Nothing

    DoCmd.Close acForm, strAuto, acSaveYes
    DoCmd.Rename TEST_FORM, acForm, strAuto
End Sub

' Advances the running Top position and returns it.
Private Function NextTop(ByRef lngTop As Long) As Long
    lngTop = lngTop + 900
    NextTop = lngTop
End Function

Private Sub RemoveTestForm()
    On Error Resume Next
    DoCmd.Close acForm, TEST_FORM, acSaveNo
    DoCmd.DeleteObject acForm, TEST_FORM
    On Error GoTo 0
End Sub

Private Sub SetVal(ByVal frm As Form, ByVal strControl As String, ByVal varValue As Variant)
    frm.Controls(strControl).Value = varValue
End Sub

' Like SetVal, but returns False if Access refuses the value.
Private Function TrySetVal(ByVal frm As Form, ByVal strControl As String, ByVal varValue As Variant) As Boolean
    On Error Resume Next
    frm.Controls(strControl).Value = varValue
    TrySetVal = (Err.Number = 0)
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Assertions. Objects come in ByVal, then go to the function under test as a plain Object
' variable, so these tests never depend on the ByRef parameter rules.
'---------------------------------------------------------------------------------------

Private Sub CheckValue(ByVal strName As String, ByVal blnExpected As Boolean, ByVal obj As Object)
    Dim blnActual As Boolean
    Dim lngErr As Long

    blnActual = TryHasValue(obj, lngErr)
    ReportResult strName, blnExpected, blnActual, lngErr
End Sub

Private Sub CheckRows(ByVal strName As String, ByVal blnExpected As Boolean, ByVal obj As Object)
    Dim blnActual As Boolean
    Dim lngErr As Long

    blnActual = TryHasRows(obj, lngErr)
    ReportResult strName, blnExpected, blnActual, lngErr
End Sub

' Same as CheckRows, with the object's TypeName in the message.
Private Sub CheckRowsTyped(ByVal strName As String, ByVal blnExpected As Boolean, ByVal obj As Object)
    CheckRows strName & " [" & TypeName(obj) & "]", blnExpected, obj
End Sub

Private Sub CheckValueRaises(ByVal strName As String, ByVal lngExpected As Long, ByVal obj As Object)
    Dim lngErr As Long

    TryHasValue obj, lngErr
    ReportRaise strName, lngExpected, lngErr
End Sub

Private Sub CheckRowsRaises(ByVal strName As String, ByVal lngExpected As Long, ByVal obj As Object)
    Dim lngErr As Long

    TryHasRows obj, lngErr
    ReportRaise strName, lngExpected, lngErr
End Sub

Private Function TryHasValue(ByRef obj As Object, ByRef lngErr As Long) As Boolean
    On Error Resume Next
    TryHasValue = HasValue(obj)
    lngErr = Err.Number
    On Error GoTo 0
End Function

Private Function TryHasRows(ByRef obj As Object, ByRef lngErr As Long) As Boolean
    On Error Resume Next
    TryHasRows = HasRows(obj)
    lngErr = Err.Number
    On Error GoTo 0
End Function

Private Sub ReportResult(ByVal strName As String, ByVal blnExpected As Boolean, _
                         ByVal blnActual As Boolean, ByVal lngErr As Long)
    If lngErr <> 0 Then
        Fail strName & "  -> raised error " & lngErr
    ElseIf blnActual = blnExpected Then
        Pass strName
    Else
        Fail strName & "  -> expected " & blnExpected & ", got " & blnActual
    End If
End Sub

Private Sub ReportRaise(ByVal strName As String, ByVal lngExpected As Long, ByVal lngErr As Long)
    If lngErr = lngExpected Then
        Pass strName
    ElseIf lngErr = 0 Then
        Fail strName & "  -> no error raised"
    Else
        Fail strName & "  -> expected error " & lngExpected & ", got " & lngErr
    End If
End Sub

Private Sub Pass(ByVal strMessage As String)
    m_Pass = m_Pass + 1
    Debug.Print "  PASS  " & strMessage
End Sub

Private Sub Fail(ByVal strMessage As String)
    m_Fail = m_Fail + 1
    Debug.Print "  FAIL  " & strMessage
End Sub

Private Sub Info(ByVal strMessage As String)
    Debug.Print "  INFO  " & strMessage
End Sub

Private Function HasValueText(ByVal obj As Object) As String
    Dim blnActual As Boolean
    Dim lngErr As Long

    blnActual = TryHasValue(obj, lngErr)
    If lngErr <> 0 Then
        HasValueText = "error " & lngErr
    Else
        HasValueText = CStr(blnActual)
    End If
End Function

Private Function Describe(ByVal varValue As Variant) As String
    If IsNull(varValue) Then
        Describe = "Null"
    Else
        Describe = CStr(varValue)
    End If
End Function
