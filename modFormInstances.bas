Attribute VB_Name = "modFormInstances"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module  : modFormInstances
' Purpose : One application-wide clsFormInstances manager, so any form or procedure can
'           open and track form instances without owning the manager itself.
'
' Install : Import with clsFormInstances.cls and clsFormInstance.cls. No references needed.
'
' Usage   : ' Open one. Configure the form first, because Add makes it visible.
'           Dim frm As Form_Customer
'           Set frm = New Form_Customer
'           frm.Filter = "CustomerID=" & lngID
'           frm.FilterOn = True
'           FormInstances.Add frm, "Customer:" & lngID
'
'           ' One window per record: bring the existing one forward instead.
'           If Not FormInstances.Activate("Customer:" & lngID) Then ...open as above...
'
'           ' Optional sweep, such as on logout or user switch.
'           CloseAllFormInstances
'
' Notes   : - Always go through FormInstances. The variable is private on purpose, so the
'             manager is recreated if a VBA project reset (End, or Reset in the VBE)
'             clears it. Forms that were already open then become untracked, but nothing
'             crashes with "Object variable not set".
'           - Use one manager for the whole application. Keep keys unique with a prefix
'             such as "Customer:123" or "Invoice:456".
'           - Open each instance as a local variable in the calling procedure. A
'             module-level variable in the caller keeps the instance alive and stops
'             Remove from closing it.
'           - CloseAllFormInstances is not needed for cleanup. Access closes the forms
'             when the application quits. Call it only for a deliberate sweep.
'           - See clsFormInstances for the remaining notes on keys, timing and dialogs.
'---------------------------------------------------------------------------------------

Private m_Instances As clsFormInstances

' The application-wide manager. Created on first use.
Public Function FormInstances() As clsFormInstances

    If m_Instances Is Nothing Then Set m_Instances = New clsFormInstances
    Set FormInstances = m_Instances

End Function

' Closes every tracked instance. Does nothing if no manager has been created.
Public Sub CloseAllFormInstances()

    If Not m_Instances Is Nothing Then m_Instances.CloseAll

End Sub
