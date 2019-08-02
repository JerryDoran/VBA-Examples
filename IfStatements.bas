Attribute VB_Name = "IfStatements"
Option Explicit

Sub TestingFilmLength()

    Dim FilmName As String
    Dim FilmLength As Integer
    Dim FilmDescription As String
    
    'Arbitrarily select a cell in the data region
    Range("B4").Select
    
    FilmName = ActiveCell.Value
    FilmLength = ActiveCell.Offset(0, 2).Value
    
    If FilmLength < 100 Then
        FilmDescription = "Short"
    ElseIf FilmLength < 120 Then
        FilmDescription = "Medium"
    ElseIf FilmLength < 150 Then
        FilmDescription = "Long"
    Else
        FilmDescription = "Epic"
    End If
    
    MsgBox FilmName & " is " & FilmDescription
    
    
End Sub

Sub NestedIf()

    Dim FilmName As String
    Dim FilmLength As Integer
    Dim FilmDescription As String
    
    'Arbitrarily select a cell in the data region
    Range("B4").Select
    
    FilmName = ActiveCell.Value
    FilmLength = ActiveCell.Offset(0, 2).Value
    
    If FilmLength < 100 Then
        FilmDescription = "Short"
    Else
        If FilmLength < 120 Then
            FilmDescription = "Medium"
        Else
            If FilmLength < 150 Then
                FilmLength = "Long"
            Else
                FilmDescription = "Epic"
            End If
        End If
    End If
    
    MsgBox FilmName & " is " & FilmDescription

End Sub

Sub MultipleConditionsInIfStatements()

    Dim FilmName As String
    Dim FilmLength As Integer
    Dim FilmDescription As String
    
    'Arbitrarily select a cell in the data region
    Range("B4").Select
    
    FilmName = ActiveCell.Value
    FilmLength = ActiveCell.Offset(0, 2).Value
    
    If FilmLength > 120 And Len(FilmName) > 15 Then
        FilmDescription = "Epic"
    Else
        FilmDescription = "Normal"
    End If
    
    MsgBox FilmName & " is " & FilmDescription
    
End Sub



















