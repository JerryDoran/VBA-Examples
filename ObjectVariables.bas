Attribute VB_Name = "ObjectVariables"
Option Explicit

Sub StoreRangeOfCells()

    Dim FilmNameCells As Range
    
    'Have to use the keyword 'Set' to assign a value to object variables
    Set FilmNameCells = Range("B3", Range("B3").End(xlDown))
    
    Sheet2.Activate
    
    FilmNameCells.Font.Color = rgbRed
    FilmNameCells.Font.Italic = False
    
End Sub

Sub ReferencingAWorksheetInAVariable()

    Dim MyNewSheet As Worksheet
    
    Set MyNewSheet = Worksheets.Add
    
    wsMovies.Activate
    Range("A1").CurrentRegion.Copy
    
    MyNewSheet.Activate
    MyNewSheet.PasteSpecial
    
End Sub

Sub OtherExample()

    Dim MyNewBook As Workbook
    
    Set MyNewBook = Workbooks.Add
    
    Dim MyNewChart As Chart
    
    Set MyNewChart = Charts.Add
    
End Sub

Sub FindingARange()

    Dim FilmToFind As String
    Dim FilmCell As Range
    
    FilmToFind = InputBox("Enter film name")
    
    Set FilmCell = Range("B3", Range("B3").End(xlDown)).Find(FilmToFind)
    
    If FilmCell Is Nothing Then
        MsgBox FilmToFind & " was not found"
    Else
        MsgBox FilmCell.Value & " was found in " & FilmCell.Address
    End If
    
End Sub











































