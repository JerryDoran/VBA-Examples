Attribute VB_Name = "SimpleEmailWordEditor"
Option Explicit

Sub CreateSimpleEmail()

    Dim ol As Outlook.Application
    Dim mi As Outlook.MailItem
    Dim doc As Word.Document
    Dim messageText As String
    
    Set ol = New Outlook.Application
    Set mi = ol.CreateItem(olMailItem)
    
    mi.Display
    mi.To = "someone@somewhere.com"
    mi.Subject = "Movies"
    
    'Get reference to the Word Editor
    Set doc = mi.GetInspector.WordEditor
    
    messageText = vbNewLine & vbNewLine & "Please reply with questions."
    doc.Range(0, 0).InsertAfter messageText
    
    
    Sheet1.ChartObjects(1).Chart.ChartArea.Copy     '---> Copy a chart from excel spreadsheet
    doc.Range(0, 0).Paste
    
    messageText = vbNewLine & vbNewLine & "Please see chart below." & vbNewLine & vbNewLine
    doc.Range(0, 0).InsertAfter messageText
    
    Sheet1.Range("A1").CurrentRegion.Copy
    doc.Range(1, 1).Paste
    
'    Will replace the entire content of the email
'    doc.Range.Text = "Dear Someone"

    messageText = "Dear Someone," & vbNewLine & vbNewLine & "Please see table below:" & vbNewLine
    
'   Inserting Text
    doc.Range.InsertBefore Text:=messageText
    
'    doc.Range(Len(messageText), Len(messageText)).InsertAfter Text:="Test"





























    
    
    
End Sub
