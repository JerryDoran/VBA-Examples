Attribute VB_Name = "QueryWebPagesTwo"
Option Explicit

Sub ImportExchangeRates()

    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim URL As String
    
    URL = "http://www.x-rates.com/table/?from=USD&amount=1"
    
    Set ws = Worksheets.Add
    
    Set qt = ws.QueryTables.Add(Connection:="URL;" & URL, Destination:=Range("A5"))
    
    With qt
        .RefreshOnFileOpen = True       'refreshes the query table each time the file is opened
        '.RefreshPeriod = 1             'automatically refreshes the table after 1 minute
        .Name = "XRates"
        .WebFormatting = xlWebFormattingRTF
        '.WebSelectionType = xlAllTables
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .Refresh
    End With
    
End Sub
Sub UpdateExchangeRates()

    Dim qt As QueryTable
    Dim URL As String
    
    If wsRates.Range("B1").Value = "" Then
        MsgBox "You must choose a currency!", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(wsRates.Range("B2").Value) Then
        MsgBox "The amount must be a number!", vbExclamation
        Exit Sub
    End If
    
    If wsRates.Range("B2").Value < 0.1 Or wsRates.Range("B2").Value > 100 Then
        MsgBox "That amount is to low or high!", vbExclamation
        Exit Sub
    End If
    
    URL = "http://www.x-rates.com/table/?from=" & wsRates.Range("B1").Value & "&amount=" & wsRates.Range("B2").Value
        
    Set qt = wsRates.QueryTables("XRates")
    
    With qt
        .Connection = "URL;" & URL
        .Refresh
    End With
    
End Sub




























