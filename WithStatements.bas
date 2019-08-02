Attribute VB_Name = "WithStatements"
Option Explicit

Sub FormatFilmReleaseDates()

    
    With Worksheets("Anything").Range("C3", Worksheets("Anything").Range("C2").End(xlDown))
        .Interior.Color = rgbAquamarine
        .Font.Color = rgbRed
        .Font.Size = 12
        .NumberFormat = "dddd dd mmm yyyy"
    End With
    
End Sub
