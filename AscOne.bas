Attribute VB_Name = "AscOne"
Option Explicit

Sub TheBasics()

'    Debug.Print Asc("5")
    
    Debug.Print Chr(120)
    
End Sub

Sub ListCharacters()

    Dim i As Integer
    
    Worksheets.Add
    
    For i = 0 To 255
        Cells(i + 1, 1).Value = i
        Cells(i + 1, 2).Value = Chr(i)
    Next i
    
End Sub

Sub ListAsciCodes()

    Dim i As Integer
    Dim s As String
    
    s = ActiveCell.Value
    
    For i = 1 To Len(s)
        Debug.Print Asc(Mid(s, i, 1))
    Next i
    
End Sub

Sub ControlCharacters()

    Const s1 As String = "Wise"
    Const s2 As String = "Owl"
    
    Debug.Print s1 & Chr(13) & Chr(10) & s2
    Debug.Print s1 & vbCrLf & s2
    Debug.Print s1 & vbNewLine & s2
    Debug.Print s1 & Chr(9) & s2
    Debug.Print s1 & vbTab & s2
    
End Sub

Sub ListUnicodeCharacters()

    Dim n As Long
    
    Application.ScreenUpdating = False
    
    Worksheets.Add
    
    For n = -32768 To 65535                 '<--Range of numbers for unicode characters
        Cells(n + 32769, 1).Value = ChrW(n)
        Cells(n + 32769, 2).Value = n
        Cells(n + 32769, 3).Value = Hex(n)
    Next n
    
    Application.ScreenUpdating = False
    
End Sub
















































