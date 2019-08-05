Attribute VB_Name = "SwitchFunction2"
Option Explicit

Sub ListFilmsByLength()

    Dim r As Range
    Dim Length As String
    Dim SheetNames()
    Dim v As Variant
    Dim Cols As Integer
    
    'Finds number of columns of data there are in the data sheet
    Cols = wsFilmData.Range("A1").CurrentRegion.Columns.Count
    
    SheetNames = Array("Short", "Medium", "Long", "Epic", "Way Too Long")
    
    For Each v In SheetNames
        CreateLengthSheet (v)
    Next v
    
    For Each r In wsFilmData.Range("A2", Range("A1").End(xlDown))

        Length = FilmLength(r.Offset(0, 3).Value)

        r.Resize(1, Cols).Copy Worksheets(Length).Range("A1048576").End(xlUp).Offset(1, 0)

    Next r
    
    For Each v In SheetNames
        Worksheets(v).Range("A1").CurrentRegion.EntireColumn.AutoFit
    Next v
    
End Sub
Sub CreateLengthSheet(ByVal SheetName As String)

    On Error GoTo CreateSheet
    Worksheets(SheetName).Cells.Clear
    On Error GoTo 0
    
    'Copy column headings from the data sheet
    wsFilmData.Range("A1", wsFilmData.Range("A1").End(xlToRight)).Copy Worksheets(SheetName).Range("A1")
    
    Exit Sub
    
CreateSheet:
    Worksheets.Add.Name = SheetName
    Resume Next
    
End Sub
