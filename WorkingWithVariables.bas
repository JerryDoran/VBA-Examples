Attribute VB_Name = "WorkingWithVariables"
Option Explicit

Dim FilmId As Integer           'If you replace the word Dim with Public then all modules in your project will have access to the varialbles.
Dim NewFilmName As String
Dim FilmDate As Date
Dim FilmLength As Integer
Sub GetUserInput()
  
    NewFilmName = InputBox("Type in a new film name")
    FilmDate = InputBox("Type in the release date")
    FilmLength = InputBox("Type in the length in minutes")
    
    Call AddFilmToList
    
End Sub

Sub AddFilmToList()

    Worksheets("Sheet1").Activate
    Range("B2").End(xlDown).Offset(1, 0).Select
    
    FilmId = ActiveCell.Offset(-1, -1).Value
    FilmId = FilmId + 1
    
    ActiveCell.Offset(0, -1).Value = FilmId
    ActiveCell.Value = NewFilmName
    ActiveCell.Offset(0, 1).Value = FilmDate
    ActiveCell.Offset(0, 2).Value = FilmLength
    
    MsgBox NewFilmName & " was added to the list"
    
End Sub
