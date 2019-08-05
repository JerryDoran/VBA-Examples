Attribute VB_Name = "WorkingWithTextFiles"
Option Explicit

'There is a procedure to create log files when a worksheet is changed.  This is under Sheet1 Worksheet_Change event.

Sub CreatingANewTextFile()

    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject
    Set ts = fso.CreateTextFile(Environ("UserProfile") & "\Desktop\Wise Owl\Test.txt")
    
    'When the text file is created it is available to have data written to it
    ts.Write "Created on " & Now & vbNewLine
    
    'Do not have to use vbNewLine character when using the 'WriteLine' method
    ts.WriteLine "Created by " & Environ("UserName")
    ts.WriteBlankLines 2
    ts.WriteLine "Data starts here"
    
    ts.Close
    Set fso = Nothing
    
End Sub

Sub AddDataToTextFile()

    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim ColCount As Integer
    Dim i As Integer
    
    Set fso = New Scripting.FileSystemObject
    
    'Opens text file to allow more data to be written to it
    Set ts = fso.OpenTextFile(Filename:=Environ("UserProfile") & "\Desktop\Wise Owl\Test.txt", IOMode:=ForAppending)
    
    Sheet1.Activate
    
    'Gets the number of columns in the data
    ColCount = Range("A2", Range("A2").End(xlToRight)).Cells.Count
    
    'This loop will process the column of cells
    For Each r In Range("A2", Range("A1").End(xlDown))
        'This loop will process the row of cells
        For i = 1 To ColCount
            ts.Write r.Offset(0, i - 1).Value
            
            If i < ColCount Then ts.Write vbTab
            
        Next i
        
        ts.WriteLine
        
    Next r
    
    ts.Close
    
    Set fso = Nothing
    
End Sub

Sub AddDataToCSVFile()

    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim ColCount As Integer
    Dim i As Integer
    
    Set fso = New Scripting.FileSystemObject
    
    Set ts = fso.OpenTextFile(Filename:=Environ("UserProfile") & "\Desktop\Wise Owl\Test.csv", IOMode:=ForAppending, Create:=True)
    
    Sheet1.Activate
    
    ColCount = Range("A2", Range("A2").End(xlToRight)).Cells.Count
    
    'This loop will process the column of cells
    For Each r In Range("A2", Range("A1").End(xlDown))
        'This loop will process the row of cells
        For i = 1 To ColCount
            ts.Write r.Offset(0, i - 1).Value
            
            If i < ColCount Then ts.Write ","
            
        Next i
        
        ts.WriteLine
        
    Next r
    
    ts.Close
    
    Set fso = Nothing
    
End Sub

Sub ReadFromTextFile()

    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim TextLine As String
    Dim TabPosition As Integer
        
    Set fso = New Scripting.FileSystemObject
    
    Set ts = fso.OpenTextFile(Filename:=Environ("UserProfile") & "\Desktop\Wise Owl\Test.txt", IOMode:=ForReading)
    
    Workbooks.Add
    
    Do Until ts.ReadLine = "Data starts here"
    Loop
    
    'Reading lines from a text file
    Do Until ts.AtEndOfStream
        TextLine = ts.ReadLine
        
        'This line will tell us the position of the first tab character in the string 'TextLine'
        TabPosition = InStr(TextLine, vbTab)
        
        Do Until TabPosition = 0
            ActiveCell.Value = Left(TextLine, TabPosition - 1)
            ActiveCell.Offset(0, 1).Select
            TextLine = Right(TextLine, Len(TextLine) - TabPosition)
            TabPosition = InStr(TextLine, vbTab)
        Loop
        
        ActiveCell.Value = TextLine
        
        ActiveCell.Offset(1, 0).End(xlToLeft).Select
    Loop
    
    ts.Close
    Set fso = Nothing
    
End Sub

Sub ReadFromTextFileEasyWay()

    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject
    
    Set ts = fso.OpenTextFile(Environ("UserProfile") & "\Desktop\Wise Owl\Test.txt", ForReading)
    
    Workbooks.Add
    
    'Reading lines from a text file
    Do Until ts.AtEndOfStream
               
        ActiveCell.Value = ts.ReadLine
        
        ActiveCell.Offset(1, 0).Select
    Loop
    
    Range("A:A").TextToColumns Tab:=True
    
    ts.Close
    Set fso = Nothing
    
End Sub

Sub ReadFromTextFileEasiestMethod()

    Workbooks.OpenText Filename:=Environ("UserProfile") & "\Desktop\Wise Owl\Test.txt", Tab:=True
    
End Sub


















































































