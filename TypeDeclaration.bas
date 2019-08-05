Attribute VB_Name = "TypeDeclaration"
Option Explicit

Type Film
    ID As Integer
    Name As String
    Date As Date
    Length As Integer
    Genre As Genres
End Type

Enum Genres
    Action
    Adventure
    Animation
    SciFi
End Enum
Sub TestFilmType()

    Dim NewFilm As Film
    
    NewFilm.ID = 99
    NewFilm.Name = InputBox("Type in a film name")
    NewFilm.Date = Range("C10").Value
    NewFilm.Genre = Adventure
    
    MsgBox NewFilm.ID & " " & NewFilm.Name & " " & GenreText(NewFilm.Genre)
    
End Sub

Function GenreText(Value As Genres) As String

    Select Case Value
        Case Action
            GenreText = "Action"
        Case Adventure
            GenreText = "Adventure"
    End Select
    
End Function























































































