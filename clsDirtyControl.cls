VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDirtyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : clsDirtyControl
' Purpose : Listens to AfterUpdate on one control and tells clsDirtyTracker. Not used
'           directly - see clsDirtyTracker for usage.
'---------------------------------------------------------------------------------------

Private Const EVENT_PROC As String = "[Event Procedure]"

Private WithEvents m_Txt As Access.TextBox
Private WithEvents m_Cbo As Access.ComboBox
Private WithEvents m_Chk As Access.CheckBox
Private m_Parent As clsDirtyTracker

Public Sub Init(Parent As clsDirtyTracker, ctl As Access.Control)

    ' A macro or =Function() in AfterUpdate would be replaced by the sink below.
    If Len(ctl.AfterUpdate) > 0 And ctl.AfterUpdate <> EVENT_PROC Then
        Err.Raise vbObjectError + 514, "clsDirtyControl.Init", _
            "Control '" & ctl.Name & "' uses a macro or expression for AfterUpdate. " & _
            "Move it to an AfterUpdate event procedure, or set its Tag to IgnoreDirty."
    End If

    Set m_Parent = Parent

    Select Case ctl.ControlType
        Case acTextBox:  Set m_Txt = ctl
        Case acComboBox: Set m_Cbo = ctl
        Case acCheckBox: Set m_Chk = ctl
    End Select

    ' Sinks the event, so the form needs no code for the control.
    ctl.AfterUpdate = EVENT_PROC

End Sub

' Disconnects from the control and the tracker.
Public Sub Detach()

    Set m_Txt = Nothing
    Set m_Cbo = Nothing
    Set m_Chk = Nothing
    Set m_Parent = Nothing

End Sub

Private Sub m_Txt_AfterUpdate()

    If Not m_Parent Is Nothing Then m_Parent.MarkDirty

End Sub

Private Sub m_Cbo_AfterUpdate()

    If Not m_Parent Is Nothing Then m_Parent.MarkDirty

End Sub

Private Sub m_Chk_AfterUpdate()

    If Not m_Parent Is Nothing Then m_Parent.MarkDirty

End Sub
