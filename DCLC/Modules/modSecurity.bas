Attribute VB_Name = "modSecurity"
Option Explicit

'Control DML Access
Public Sub ValidateAccessLevel(ByRef thisForm As Form, ByRef strAccessRequest As String)
    On Error Resume Next
    'thisForm.cmdSave.Enabled = False
    thisForm.cmdSave.Caption = strAccessRequest
    thisForm.cmdSave.Enabled = ((LCase(User.UserType) = "registrar") And (LCase(strAccessRequest) <> "update")) Or LCase(User.UserType) <> "registrar"
    thisForm.cmdDelete.Enabled = (LCase(User.UserType) = "administrator")
End Sub
