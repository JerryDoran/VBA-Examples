Attribute VB_Name = "StringConcatenation"
Option Explicit

Sub BasicConcatenation()

    Dim s As String
    
    Range("A2").Select
    
    s = ActiveCell.Value & "," & ActiveCell.Offset(0, 2).Value
    
    Debug.Print s
    
End Sub

Sub AccumulatingStrings()

    Dim s As String
    Dim r As Range
    Dim Cols As Integer
    
    Cols = Range("A1").CurrentRegion.Columns.Count
    
    For Each r In ActiveCell.Resize(1, Cols)
        's = s & r.Value & ","
        's = s & r.Value & vbTab
        s = s & r.Value & vbNewLine
    Next r
    
    s = Left(s, Len(s) - 1)     ' strips off the last character at the end of a string
    
    Debug.Print s
    
End Sub

Sub JoinFunction()

    Dim s As String
    
    s = Join(Array("a", "b", "c"), vbTab)
    
    Debug.Print s
    
End Sub

Sub JoiningRangeValues()

    Dim s As String
    Dim Cols As Integer
    Dim arr() As Variant
    
    Cols = Range("A1").CurrentRegion.Columns.Count
    
    'Need to use the Transpose function twice to create a one dimensional array for the Join function
    arr = Application.Transpose(Application.Transpose(ActiveCell.Resize(1, Cols).Value))
    
    s = Join(arr, vbTab)    '<--Join function can only accept a one-dimensional array.
        
    Debug.Print s
     
End Sub

Sub ListActionFilms()

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim s As String
    Dim Cols As Integer
    Dim arr() As Variant
    
    Cols = Range("A1").CurrentRegion.Columns.Count
    
    'Using the OpenTextFile method will open an existing text file if it is there. If not this method will also create the text file by setting create parameter to true.
    Set ts = fso.OpenTextFile(Environ("UserProfile") & "\Desktop\Action.txt", ForAppending, True)
    
    For Each r In Range("A2", Range("A1").End(xlDown))
    
        If LCase(r.Offset(0, 5).Value) = "action" Then
            arr = Application.Transpose(Application.Transpose(r.Resize(1, Cols).Value))
            s = Join(arr, vbTab)
            ts.WriteLine s
        End If
        
    Next r
    
    ts.Close
     
End Sub

Sub CreateGenreFiles()

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim s As String
    Dim Cols As Integer
    Dim arr() As Variant
    Dim FolPath As String
    Dim Genre As String
    
    Cols = Range("A1").CurrentRegion.Columns.Count
    
    FolPath = Environ("UserProfile") & "\Desktop\Genres"
    
    If Not fso.FolderExists(FolPath) Then
        fso.CreateFolder (FolPath)
    End If
    
    For Each r In Range("A2", Range("A1").End(xlDown))
        
        Genre = r.Offset(0, 5).Value
        Set ts = fso.OpenTextFile(FolPath & "\" & Genre & ".txt", ForAppending, True)
        
        arr = Application.Transpose(Application.Transpose(r.Resize(1, Cols).Value))
        s = Join(arr, vbTab)
        ts.WriteLine s
        ts.Close
        
    Next r
     
End Sub













































