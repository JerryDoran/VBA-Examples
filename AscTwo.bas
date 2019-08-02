Attribute VB_Name = "AscTwo"
Option Explicit

Enum CharType
    Unicode = 0
    UpperCase = 1
    LowerCase = 2
    Number = 3
    Other = 4
End Enum

Sub ListingCharacterTypes()

    Dim r As Range
    Dim n As Long
    Dim nRows As Long
    Dim s As String
    Dim i As Integer
    Dim ct As CharType
    Dim arr()
    
    wsFilmData.Select
    
    Set r = Range("A2", Range("A1").End(xlDown))
    nRows = r.Rows.Count
    
    ReDim arr(1 To nRows, 0 To 4)
    
    For n = 1 To nRows
        s = r.Cells(n, 1).Value
        For i = 1 To Len(s)
            ct = CharacterType(Mid(s, i, 1))    '<--pass in the string (s), start in position i, and return one character of that string (evaluates one character at a time)
'            Debug.Print ct
            arr(n, ct) = arr(n, ct) + 1
        Next i
    Next n
    
    r.Offset(0, 1).Resize(nRows, 5) = arr
    
    
End Sub

Function CharacterType(Character As String) As CharType

    Select Case AscW(Character)
        Case Is < 0
            CharacterType = Unicode
        Case Is > 255
            CharacterType = Unicode
        Case 48 To 57
            CharacterType = Number
        Case 65 To 90
            CharacterType = UpperCase
        Case 97 To 122
            CharacterType = LowerCase
        Case Else
            CharacterType = Other
    End Select
End Function
































