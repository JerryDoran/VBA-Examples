VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OwlMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private olApp As New Outlook.Application      '<--Create auto-instancing object variable
Public WithEvents Email As Outlook.MailItem
Attribute Email.VB_VarHelpID = -1
Private mailSent As Boolean

Private Sub Class_Initialize()
    
    Set Email = olApp.CreateItem(olMailItem)
    mailSent = False
    
End Sub

Private Sub Email_Close(Cancel As Boolean)
    
    Dim CloseAndDiscard As VbMsgBoxResult
    
    If mailSent Then Exit Sub  '<--If mail has been sent then the rest of the close event is ignored
    
    Email.GetInspector.WindowState = olMinimized
    AppActivate Application.Caption
    
    CloseAndDiscard = MsgBox("Close and discard?", vbYesNo + vbExclamation)
    
    If CloseAndDiscard = vbYes Then
        Email.Close olDiscard
'        Debug.Print "Owl mail was closed", Now
        RecordStatus
        Set olEmail = Nothing
    Else
        Cancel = True
        Email.Display
        Exit Sub
    End If
    
End Sub

Private Sub Email_Send(Cancel As Boolean)

'    Debug.Print "Owl mail was sent", Now
    mailSent = True
    RecordStatus
    Set olEmail = Nothing
    
End Sub

Private Sub RecordStatus()

    Sheet1.Range("A1048576").End(xlUp).Offset(1, 0).Value = mailSent
    Sheet1.Range("B1048576").End(xlUp).Offset(1, 0).Value = Now
    Sheet1.Range("C1048576").End(xlUp).Offset(1, 0).Value = Environ("UserName")
    
End Sub
































