Attribute VB_Name = "PivotTableBasics"
Option Explicit
Sub DeleteAllButMoviesSheet()

    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is wsMovies Then ws.Delete
    Next ws
    
End Sub
Sub DeleteAllConnections()

    Dim cn As WorkbookConnection
    
    For Each cn In ThisWorkbook.Connections
        cn.Delete
    Next cn
End Sub

Sub CreatePivotTableAndPivotCache()

    'A pivot cash is a copy of the source data that pivot tables are based on
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    'Data source is a single excel table - Use xlDatabase
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsMovies.Name & "!" & wsMovies.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)
        
     Worksheets.Add
     Range("A3").Select
     
     Set pt = pc.CreatePivotTable( _
        Tabledestination:=ActiveCell, _
        TableName:="MoviePivot")
        
    Debug.Print ThisWorkbook.PivotCaches.Count
    Debug.Print pc.MemoryUsed, pc.RecordCount, pc.Version
        
End Sub
Sub AddPivotTableAndPivotCache()

    'A pivot cash is a copy of the source data that pivot tables are based on
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ws As Worksheet
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsMovies.Name & "!" & wsMovies.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)
        
     Set ws = Worksheets.Add
     Range("A3").Select
     
     Set pt = ws.PivotTables.Add( _
        PivotCache:=pc, _
        Tabledestination:=ActiveCell, _
        TableName:="MoviePivot2")
                
    Debug.Print ThisWorkbook.PivotCaches.Count
    Debug.Print pc.MemoryUsed, pc.RecordCount, pc.Version
        
End Sub

Sub CreatePivotTableUsingExistingPivotCache()

    'A pivot cash is a copy of the source data that pivot tables are based on
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    If ThisWorkbook.PivotCaches.Count = 0 Then
        Set pc = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=wsMovies.Name & "!" & wsMovies.Range("A1").CurrentRegion.Address, _
            Version:=xlPivotTableVersion15)
    Else
        Set pc = ThisWorkbook.PivotCaches(1)        'Use an existing pivot cache -- indexed starting from 1
    End If
    
    Worksheets.Add
    Range("A3").Select
     
    Set pt = pc.CreatePivotTable( _
       Tabledestination:=ActiveCell, _
       TableName:="MoviePivot")
        
    Debug.Print ThisWorkbook.PivotCaches.Count
    Debug.Print pc.MemoryUsed, pc.RecordCount, pc.Version
        
End Sub





















