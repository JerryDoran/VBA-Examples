Attribute VB_Name = "UsingOutlook"
Option Explicit

Public olEmail As OwlMail

Sub SendBasicEmail()

    'Check to see if there already is an olEmail object created
    If Not olEmail Is Nothing Then
        MsgBox "You have an email to deal with already", vbCritical
        olEmail.Email.Display
        Exit Sub
    End If
    
    Set olEmail = New OwlMail       '<--Creates new instance of the OwlMail class
    
    With olEmail.Email
        .BodyFormat = olFormatHTML  '<--Set this property first to be able to display a signature in the email
        .Display
        .HTMLBody = "<H1>Test email from Excel</H1><br>" & "<br>" & .HTMLBody
        '.Attachments.Add Environ("UserProfile") & "\Desktop\Wise Owl\Test.txt"
        .Display
        .To = "someone@somewhere.com;someoneelse@somewhereelse.com"   '<--Can also put in the name of a distribution list
        .Subject = "Movie Report"
        '.send - will automatically send out emails.
    End With
    
End Sub

Sub SendBasicEmailLateBinding()
'
''Using late binding means we can uncheck the reference to use Outlook from Excel's reference library
'
'    Dim olApp As Object
'    Dim olEmail As Object
'
'    Set olApp = CreateObject("Outlook.Application")
'    Set olEmail = olApp.CreateItem(0)      '<--The 0 represents the value of the constant olMailItem
'
'    With olEmail
'        .BodyFormat = 2        '<--The 2 represents the value of the constant olFormatHTML
'        .Display
'        .HTMLBody = "<H1>Test email from Excel</H1><br>" & "<br>" & .HTMLBody
'        .Attachments.Add Environ("UserProfile") & "\Desktop\Wise Owl\Test.txt"
'        .Display
'        .To = "debbie.johnson@pccairfoils.com"
'        .Subject = "Movie Report"
'        '.send - will automatically send out emails.
'    End With
'
End Sub






















