Attribute VB_Name = "CreateFilms"
Option Explicit

Sub CreateFilmFiles()

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim FolderPath As String
    Dim FileName As String
    Dim Cols As Integer
    Dim IllegalChars()
    Dim v As Variant
    
    IllegalChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    
    FolderPath = Environ("UserProfile") & "\Desktop\Films"
    
    If fso.FolderExists(FolderPath) Then
        fso.DeleteFolder FolderPath
    End If
    
    fso.CreateFolder FolderPath
    
    Cols = Range("A1").CurrentRegion.Columns.Count
    
    For Each r In Range("A2", Range("A1").End(xlDown))
        
        FileName = r.Value
        
        For Each v In IllegalChars
            FileName = Replace(FileName, v, "")
        Next v
        
        FileName = FileName & ".txt"
        
        Set ts = fso.OpenTextFile(FolderPath & "\" & FileName, ForAppending, True)
        
        'write to text
        ts.WriteLine Join(Application.Transpose(Application.Transpose(r.Resize(1, Cols).Value)), vbTab)
        
        ts.Close
        
        Set ts = Nothing
        
    Next r
End Sub
