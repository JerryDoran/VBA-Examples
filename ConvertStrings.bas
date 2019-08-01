Attribute VB_Name = "ConvertStrings"
Option Explicit
'Option Compare Text     'case insensitive

Sub ComparingStringCase()

    Dim s As String
    Dim r As Range
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
    Set ws = Worksheets.Add
    
    wsFilmData.Range("A1").EntireRow.Copy ws.Range("A1")
    Range("A2").Select
    
    For Each r In wsFilmData.Range("A2", wsFilmData.Range("A1").End(xlDown))
    
        's = LCase(r.Offset(0, 5).Value)
        s = r.Offset(0, 5).Value
        
        'If s = "action" Then Debug.Print r.Value
        If StrComp(s, "action", vbTextCompare) = 0 Then     '<--the two strings are equal if the StrComp function equals zero.
            r.EntireRow.Copy ActiveCell
            ActiveCell.Offset(1, 0).Select
        End If
    
    Next r
    
    Range("A1").Select
    ActiveCell.CurrentRegion.EntireColumn.AutoFit
    Application.ScreenUpdating = True
    
End Sub
