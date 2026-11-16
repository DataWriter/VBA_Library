VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDirtyTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : clsDirtyTracker
' Purpose : Dirty flag for an UNBOUND form. Replaces the usual module-level "mbDirty" and
'           the one-line AfterUpdate handler on every control. The flag lives in the
'           form's own module, so each instance of a form has its own - it works the same
'           for instanced forms (clsFormInstances) and ordinary ones.
'
' Install : Import clsDirtyTracker.cls and clsDirtyControl.cls. No references needed.
'
' Usage   : Private m_Dirty As clsDirtyTracker
'
'           Private Sub Form_Load()
'               Set m_Dirty = New clsDirtyTracker
'               m_Dirty.Watch Me            ' call AFTER the form is set up
'           End Sub
'
'           Private Sub Form_Close()
'               m_Dirty.Unwatch             ' Required - see Notes
'           End Sub
'
'           ' Was: mbDirty = True / If mbDirty Then / mbDirty = False
'           m_Dirty.MarkDirty               ' only for changes made in code
'           If m_Dirty.IsDirty Then ...     ' prompt to save
'           m_Dirty.Clear                   ' after populating or saving a record
'
' Notes   : - Watches text boxes, combo boxes and check boxes. A control whose Tag
'             contains IgnoreDirty is skipped, as are calculated controls.
'           - A change is noticed when the user's edit completes (AfterUpdate). Values
'             set in code do not count - call MarkDirty for those.
'           - Existing AfterUpdate procedures on the form keep working; the tracker
'             hears the event as well. A control that uses a macro or =Function() for
'             AfterUpdate raises an error in Watch, since the tracker would replace it.
'           - Call Clear after loading a record, or the load itself must not fire
'             AfterUpdate (assigning .Value in code does not).
'           - Unwatch is required for the same reason as clsFormInstances.CloseAll:
'             each watched control keeps its watcher alive, so it never terminates by
'             itself.
'           - Subform controls are not watched. Use the subform's own Dirty if bound.
'---------------------------------------------------------------------------------------

Private m_Items As Collection
Private m_Dirty As Boolean

Private Sub Class_Initialize()

    Set m_Items = New Collection

End Sub

Private Sub Class_Terminate()

    Unwatch

End Sub

' Starts watching the form's controls and clears the flag. Safe to call again.
Public Sub Watch(frm As Access.Form)

    Dim ctl As Access.Control
    Dim oItem As clsDirtyControl

    Unwatch

    For Each ctl In frm.Controls
        If IsWatchable(ctl) Then
            Set oItem = New clsDirtyControl
            oItem.Init Me, ctl
            m_Items.Add oItem
        End If
    Next ctl

    m_Dirty = False

End Sub

' Stops watching and releases every control.
Public Sub Unwatch()

    Dim oItem As clsDirtyControl

    If m_Items Is Nothing Then Exit Sub

    For Each oItem In m_Items
        oItem.Detach
    Next oItem

    Set m_Items = New Collection

End Sub

Public Property Get IsDirty() As Boolean

    IsDirty = m_Dirty

End Property

Public Sub MarkDirty()

    m_Dirty = True

End Sub

Public Sub Clear()

    m_Dirty = False

End Sub

Private Function IsWatchable(ctl As Access.Control) As Boolean

    Select Case ctl.ControlType
        Case acTextBox, acComboBox, acCheckBox
            If InStr(1, ctl.Tag, "IgnoreDirty", vbTextCompare) > 0 Then Exit Function
            If Left$(Nz(ctl.ControlSource, vbNullString), 1) = "=" Then Exit Function
            IsWatchable = True
    End Select

End Function
