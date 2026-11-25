Attribute VB_Name = "modHasValue"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : modHasValue
' Purpose : Two small tests, so callers never have to remember the right check:
'             HasValue(ctl)  Did the user enter or select something in this control?
'             HasRows(rst)   Did this recordset return any rows?
'           Both work on bound and unbound forms.
'
' Install : Import the .bas. Needs no references: controls are recognised by Access's own
'           types, and DAO and ADO recordsets are recognised by name. Tests are in
'           modHasValueTests.bas (optional, development database only).
'
' Usage   : If HasValue(Me.txtName) Then ...
'           If Not HasValue(frm.Controls("cboStatus")) Then ...
'           If HasRows(rst) Then ...           ' ADO or DAO recordset
'           If HasRows(Me.Recordset) Then ...  ' a bound form's recordset is DAO
'
' Rules   : HasValue
'           TextBox      Null, "" and whitespace-only are empty. Whitespace means spaces,
'                        tabs, line breaks, non-breaking spaces and zero-width spaces.
'           ComboBox     Limit To List: empty unless an item is selected (ListIndex),
'                        so an invalid stored value does not count as a value.
'                        Not Limit To List: empty if Null, "" or whitespace-only.
'           ListBox      Empty unless an item is selected, single or multi select. Same
'                        rule as a limit-to-list combo: selection, not the stored value.
'           CheckBox, OptionGroup, ToggleButton
'                        Only Null is empty. False (0) is a real value.
'           Label        Empty if Caption is "".
'
'           HasRows
'           ADO          False if closed, or BOF and EOF.
'           DAO          False if BOF and EOF, or if the recordset is closed or otherwise
'                        invalid (error 3420), like a closed ADO recordset.
'
' Notes   : - Fails loudly. Nothing, or a type the function has no check for (a recordset
'             passed to HasValue, a control passed to HasRows, an option button, a
'             command button), raises an error instead of returning an answer that
'             looks valid. Errors are not logged here; the caller's own handler decides
'             what to do with them.
'           - Recordsets are told apart by TypeName: "Recordset2" is DAO (Access 2007 and
'             later), "Recordset" is ADO. An old DAO 3.6 recordset also reports
'             "Recordset", and is not supported.
'           - A bound control raises 2424 when its form has no recordset (for example a
'             check run before the recordset is assigned). Deliberately not handled: it
'             means the form is misconfigured or the check ran too early. Check
'             frm.Recordset first, or run the check later.
'           - BOF/EOF describes cursor position, not row count. Immediately after the last
'             remaining record is deleted, both can stay False until the recordset is
'             repositioned, so HasRows can say True for a recordset with no rows left.
'             Not corrected here because that would require moving the caller's cursor.
'           - Pass a control through its form, frm.Controls("x"), not Me.Parent.Controls
'             ("x"); the latter failed in 2017 (see history).
'           - ByRef on the object parameters is kept on purpose (see history).
'
' History : 8/16/17  NoValue: use frm.Controls("x") instead of Me.Parent.Controls("x").
'           8/25/17  NoValue: ComboBox judged by ListIndex, because a Null check lets an
'                    invalid value pass.
'           9/19/17  NoValue: parameter changed to ByRef, error handler added.
'           11/25/26 NoValue replaced by HasValue (positive form, no double negatives),
'                    and split: recordsets now have their own HasRows. Controls tested
'                    with TypeOf instead of TypeName strings. Added CheckBox, OptionGroup,
'                    ToggleButton, Label, multi select ListBox and Limit To List
'                    handling. Nothing and unsupported types now raise (91 and 5).
'                    Reviewed against the Code Review answers to question 175679 and a
'                    Copilot review: closed-DAO handling in its own helper so it cannot
'                    mask other errors; editable ComboBox now trims like TextBox; error
'                    2424 is no longer handled; no DAO or ADO reference needed; ListBox
'                    is judged by selection like a limit-to-list ComboBox; tabs, line
'                    breaks and non-breaking and zero-width spaces count as blank text;
'                    error 3420 documented as "invalid", not only "closed"; NoValue
'                    removed.
'---------------------------------------------------------------------------------------

Private Const ADO_STATE_CLOSED As Long = 0        ' avoids an ADO reference
Private Const ERR_INVALID_ARG As Long = 5
Private Const ERR_NULL_OBJECT As Long = 91
Private Const ERR_DAO_INVALID As Long = 3420       ' "Object is invalid or no longer set". UNVERIFIED that BOF raises it on a closed recordset, test.

Public Function HasValue(ByRef ctl As Object) As Boolean
    Dim result As Boolean

    If ctl Is Nothing Then
        Err.Raise ERR_NULL_OBJECT, "modHasValue.HasValue", _
                  "Object to test cannot be Nothing."
    End If

    On Error GoTo HandleError

    Select Case True
    Case TypeOf ctl Is Access.TextBox
        result = Not IsBlank(ctl.Value)

    Case TypeOf ctl Is Access.ComboBox
        ' A Null check alone lets an invalid stored value pass (8/25/17), so a
        ' limit-to-list combo is judged by whether an item is selected.
        If ctl.LimitToList Then
            result = (ctl.ListIndex <> -1)
        Else
            result = Not IsBlank(ctl.Value)
        End If

    Case TypeOf ctl Is Access.ListBox
        ' Selection, not Value: same rule as a limit-to-list combo, and it also covers
        ' multi select, where Value is always Null.
        result = (ctl.ItemsSelected.Count > 0)

    Case TypeOf ctl Is Access.CheckBox, TypeOf ctl Is Access.OptionGroup, _
         TypeOf ctl Is Access.ToggleButton
        result = Not IsNull(ctl.Value)              ' False/0 is a value; only Null is empty

    Case TypeOf ctl Is Access.Label
        result = Len(ctl.Caption) > 0

    Case Else
        Err.Raise ERR_INVALID_ARG, "modHasValue.HasValue", _
                  "No check implemented for type " & TypeName(ctl) & "."
    End Select

    HasValue = result
    Exit Function

HandleError:
    Err.Raise Err.Number, "modHasValue.HasValue", Err.Description
End Function

Public Function HasRows(ByRef rst As Object) As Boolean
    Dim result As Boolean

    If rst Is Nothing Then
        Err.Raise ERR_NULL_OBJECT, "modHasValue.HasRows", _
                  "Recordset to test cannot be Nothing."
    End If

    On Error GoTo HandleError

    Select Case TypeName(rst)
    Case "Recordset2"                               ' DAO
        result = DaoRecordsetHasRows(rst)

    Case "Recordset"                                ' ADO
        If rst.State = ADO_STATE_CLOSED Then
            result = False                          ' do not touch BOF on a closed recordset
        Else
            result = Not (rst.BOF And rst.EOF)
        End If

    Case Else
        Err.Raise ERR_INVALID_ARG, "modHasValue.HasRows", _
                  "No check implemented for type " & TypeName(rst) & "."
    End Select

    HasRows = result
    Exit Function

HandleError:
    Err.Raise Err.Number, "modHasValue.HasRows", Err.Description
End Function

' Row test for a DAO recordset. An invalid recordset (3420: closed, orphaned or otherwise
' unusable) counts as no rows, the same as a closed ADO recordset. This is the only place
' 3420 is handled, so it cannot hide an error from anywhere else.
Private Function DaoRecordsetHasRows(ByRef rst As Object) As Boolean
    On Error GoTo HandleError

    DaoRecordsetHasRows = Not (rst.BOF And rst.EOF)
    Exit Function

HandleError:
    If Err.Number = ERR_DAO_INVALID Then
        DaoRecordsetHasRows = False
    Else
        Err.Raise Err.Number, "modHasValue.DaoRecordsetHasRows", Err.Description
    End If
End Function

' True if the value is Null, "", or only whitespace. Trim removes ordinary spaces only, so
' tabs, line breaks, non-breaking spaces (common in pasted text) and zero-width spaces are
' turned into spaces first. ChrW keeps this source file plain ASCII.
Private Function IsBlank(ByVal varValue As Variant) As Boolean
    Dim strText As String

    strText = varValue & vbNullString               ' Null becomes ""
    strText = Replace(strText, vbTab, " ")
    strText = Replace(strText, vbCr, " ")
    strText = Replace(strText, vbLf, " ")
    strText = Replace(strText, ChrW$(160), " ")     ' non-breaking space
    strText = Replace(strText, ChrW$(8203), " ")    ' zero-width space

    IsBlank = (Len(Trim$(strText)) = 0)
End Function
