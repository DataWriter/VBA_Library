VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : clsFormInstance
' Purpose : Holds one open form instance for clsFormInstances and tells it when the
'           form closes. Not used directly - see clsFormInstances for usage.
'---------------------------------------------------------------------------------------

Private Const EVENT_PROC As String = "[Event Procedure]"

Private WithEvents m_Form As Access.Form
Private m_Parent As clsFormInstances
Private m_Key As String

Public Sub Init(Parent As clsFormInstances, frm As Access.Form, Key As String)

    ' A macro or =Function() in OnClose would be replaced by the sink below.
    If Len(frm.OnClose) > 0 And frm.OnClose <> EVENT_PROC Then
        Err.Raise vbObjectError + 513, "clsFormInstance.Init", _
            "Form '" & frm.Name & "' uses a macro or expression for OnClose. " & _
            "Move it to a Form_Close event procedure."
    End If

    Set m_Parent = Parent
    Set m_Form = frm
    m_Key = Key

    ' Sinks the Close event, so the form itself needs no code.
    m_Form.OnClose = EVENT_PROC

End Sub

Public Property Get Form() As Access.Form

    Set Form = m_Form

End Property

Public Property Get Key() As String

    Key = m_Key

End Property

' Disconnects from the form and the manager. Clearing m_Form explicitly matters: the
' event sink keeps this object alive, so dropping it from the collection is not enough.
Public Sub Detach()

    Set m_Parent = Nothing
    Set m_Form = Nothing

End Sub

Private Sub m_Form_Close()

    Dim oParent As clsFormInstances

    ' Local reference: Removed releases this object, and possibly the manager too.
    Set oParent = m_Parent
    If Not oParent Is Nothing Then oParent.Removed m_Key

End Sub
