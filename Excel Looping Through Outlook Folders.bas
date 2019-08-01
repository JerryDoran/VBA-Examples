Attribute VB_Name = "Module1"
Option Explicit

Sub AccessNonDefaultFolder()

    Dim ol As Outlook.Application
    Dim ns As Outlook.Namespace
    Dim fol As Outlook.Folder
    Dim rootfol As Outlook.Folder
    Dim i As Object
    Dim mi As Outlook.MailItem
    Dim n As Long
        
    Set ol = New Outlook.Application
    Set ns = ol.GetNamespace("MAPI")
    Set rootfol = ns.Folders(1)
    
    Set fol = rootfol.Folders("Inbox").Folders("Plants").Folders("Cannon").Folders("Chris Serene")
    
    Worksheets.Add
    
    For Each i In fol.Items
        If i.Class = olMail Then
            Set mi = i
            n = n + 1
            Cells(n, 1).Value = mi.Sender
            Cells(n, 2).Value = mi.ReceivedTime
            Cells(n, 3).Value = mi.Body
        End If
    Next i
    
            
End Sub
