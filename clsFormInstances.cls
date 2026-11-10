VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormInstances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : clsFormInstances
' Purpose : Keeps multiple instances of Access forms open and removes each one
'           automatically when the user closes it. Works with any form that has a
'           code module (HasModule = Yes). The form needs no code of its own.
'
' Install : Import clsFormInstances.cls and clsFormInstance.cls. No references needed.
'
' Usage   : ' In the owner (a main form, a controller class, or a module-level variable).
'           ' Must outlive the instances.
'           Private m_Instances As clsFormInstances
'
'           Private Sub Form_Load()
'               Set m_Instances = New clsFormInstances
'           End Sub
'
'           Private Sub Form_Close()
'               m_Instances.CloseAll    ' Required - see Notes
'           End Sub
'
'           ' Open one. The caller creates it because VBA cannot "New" a form by name.
'           Dim frm As Form_Customer
'           Set frm = New Form_Customer
'           frm.Filter = "CustomerID=" & lngID
'           frm.FilterOn = True
'           m_Instances.Add frm, "Customer:" & lngID
'
'           ' One window per record: bring the existing one forward instead.
'           If Not m_Instances.Activate("Customer:" & lngID) Then ...open as above...
'
'           ' Application level: import modFormInstances.bas and call FormInstances.Add
'           ' from anywhere instead of declaring m_Instances. No owner form is needed.
'
' Notes   : - Key is optional. Without it, the form's hWnd is used. Add returns the key.
'           - Add makes the form visible, so set it up first.
'           - Open the form as a local variable, as above. A module-level variable in the
'             caller keeps the instance alive and stops Remove from closing it.
'           - CloseAll is required because each item refers back to the manager, so the
'             manager never terminates by itself while instances are open. Without
'             CloseAll they stay open after the owner closes.
'           - Remove and CloseAll update Count at once, but Access only closes the form
'             at its next idle or DoEvents. Forms.Count lags until then.
'           - Instances share one name, so Forms("Customer") and DoCmd.Close acForm
'             reach only the first one. Use Item() and Remove() instead.
'           - Instances do not support acDialog. They open non-modal, and code continues
'             straight after Add.
'           - New instances open on top of each other unless the form sets its position.
'---------------------------------------------------------------------------------------

Private m_Items As Collection

Private Sub Class_Initialize()

    Set m_Items = New Collection

End Sub

Private Sub Class_Terminate()

    Set m_Items = Nothing

End Sub

' Adds an open instance, shows it, and returns its key.
' Raises error 457 if the key is already in use.
Public Function Add(frm As Access.Form, Optional Key As String) As String

    Dim oItem As clsFormInstance
    Dim sKey As String

    sKey = Key
    If Len(sKey) = 0 Then sKey = CStr(frm.hWnd)
    If Exists(sKey) Then Err.Raise 457, "clsFormInstances.Add", "Key already in use: " & sKey

    Set oItem = New clsFormInstance
    oItem.Init Me, frm, sKey
    m_Items.Add oItem, sKey

    frm.Visible = True
    Add = sKey

End Function

Public Property Get Count() As Long

    Count = m_Items.Count

End Property

Public Function Exists(Key As String) As Boolean

    Dim oItem As clsFormInstance

    On Error Resume Next
    Set oItem = m_Items(Key)
    Exists = Not oItem Is Nothing

End Function

' Returns the form, or Nothing if the key is not open.
Public Function Item(Key As String) As Access.Form

    If Exists(Key) Then Set Item = m_Items(Key).Form

End Function

' Brings the instance forward. Returns False if the key is not open.
Public Function Activate(Key As String) As Boolean

    If Exists(Key) Then
        m_Items(Key).Form.SetFocus
        Activate = True
    End If

End Function

' Closes one instance. Does nothing if the key is not open.
Public Sub Remove(Key As String)

    If Exists(Key) Then Removed Key

End Sub

' Closes every instance.
Public Sub CloseAll()

    Do While m_Items.Count > 0
        Removed m_Items(1).Key
    Loop

End Sub

' Drops the item, which closes the form if nothing else refers to it.
' Called by Remove, CloseAll and clsFormInstance when the user closes the form.
Public Sub Removed(Key As String)

    Dim oItem As clsFormInstance

    Set oItem = m_Items(Key)
    m_Items.Remove Key
    oItem.Detach

End Sub
