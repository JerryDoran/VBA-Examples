Attribute VB_Name = "BasicConcatenation"
Option Explicit

Sub BasicConcatenation()

    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    
    s1 = "a"
    s2 = "b"
    s3 = s1 & s2
    
    Debug.Print s3
    
End Sub
Sub ConcatenateCellValues()

    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    
    wsFilmData.Select
    Range("A9").Select
    
    s1 = ActiveCell.Value
    s2 = ActiveCell.Offset(0, 3).Value
    s3 = s1 & ", " & s2
    
    Debug.Print s3
    
End Sub

Sub ConcatenationOperators()

    'Debug.Print "a" & "b"
    'Debug.Print "2" & 2
    
    Debug.Print ActiveCell.Value & ", " & ActiveCell.Offset(0, 3).Value
    
End Sub

Sub ConcatenateTabs()

    Dim s As String
    Dim r As Range
    
    For Each r In Range(ActiveCell, ActiveCell.End(xlToRight))
        s = s & r.Value & vbTab
    Next r
    
    Debug.Print s
    MsgBox s
    
    
End Sub

Sub ConcatenateNewLines()

    Dim s As String
    Dim r As Range
    
    For Each r In Range(ActiveCell, ActiveCell.End(xlToRight))
        s = s & r.Value & IIf(r.Offset(0, 1).Value = "", "", vbNewLine)
    Next r
    
    Worksheets.Add
    ActiveCell.Value = s
    
'    Debug.Print s
'    MsgBox s
    
End Sub

Sub AddingNewLines()

    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    
    s1 = "a"
    s2 = "b"
    s3 = s1 & vbNewLine & s2
    
    Debug.Print s3
    Range("A4").Value = s3
    MsgBox s3
    
End Sub




















































