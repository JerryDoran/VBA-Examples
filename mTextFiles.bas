Attribute VB_Name = "mTextFiles"
Option Explicit

Sub SplitTabDelimitedData()

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim arr() As String
    Dim i As Integer
    Dim j As Integer
    
    Set ts = fso.OpenTextFile(Environ("UserProfile") & "\Desktop\HighGross.txt")
    
    Worksheets.Add
    
    Do Until ts.AtEndOfStream
        arr = Split(ts.ReadLine, vbTab)
        
        For i = LBound(arr) To UBound(arr)
            Cells(i + 1, j + 1).Value = arr(i)
        Next i
        
        j = j + 1
        
        Erase arr
        
    Loop
    
    ts.Close
    ActiveCell.CurrentRegion.EntireColumn.AutoFit
    
End Sub
