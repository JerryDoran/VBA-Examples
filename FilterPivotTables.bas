Attribute VB_Name = "FilterPivotTables"
Option Explicit

Sub EditingPivotTable()

    'A pivot cash is a copy of the source data that pivot tables are based on
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    
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
        
'    Set pf = pt.PivotFields("Genre")
'    pf.Orientation = xlRowField
'    Set pf = pt.PivotFields("Country")
'    pf.Orientation = xlRowField
'    pf.Position = 1                 'position the "Country" field to be the first row field
'
'    Set pf = pt.PivotFields("Certificate")
'    pf.Orientation = xlColumnField
'
'    Set pf = pt.PivotFields("Oscar Wins")
'    pf.Orientation = xlDataField
'
'    Set pf = pt.PivotFields("Distributor")
'    pf.Orientation = xlPageField
'    Set pf = pt.PivotFields("Studio")
'    pf.Orientation = xlPageField
'    pf.Position = 1                 'Page fields positioning starts from the bottom up.  First position starts from bottom of page fields.

'-------------------------------OR-------------------------------------
    
    pt.AddFields _
        RowFields:=Array("Genre", "Country"), _
        ColumnFields:="Certificate", _
        PageFields:=Array("Studio", "Language")
        
    pt.AddDataField pt.PivotFields("Oscar Wins"), , xlAverage

    Set pf = pt.DataFields(1)
    pf.Function = xlMax
    
End Sub

Sub FilteringPivotTable()

    'A pivot cash is a copy of the source data that pivot tables are based on
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    
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
        
    Set pf = pt.PivotFields("Genre")
    pf.Orientation = xlRowField
   
    Set pf = pt.PivotFields("Certificate")
    pf.Orientation = xlColumnField

    Set pf = pt.PivotFields("Oscar Wins")
    pf.Orientation = xlDataField

    Set pf = pt.PivotFields("Country")
    pf.Orientation = xlPageField
    
End Sub

Sub FilterExistingPivotTable()

    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem
    
    Set pt = ActiveSheet.PivotTables("MoviePivot")
    Set pf = pt.PivotFields("Country")
    
    
'    For Each pi In pf.PivotItems
'        If pi.RecordCount >= 10 Then
'            pi.Visible = True
'        Else
'            pi.Visible = False
'        End If
'    Next pi

    pf.ClearAllFilters

'    pf.PivotFilters.Add2 xlValueIsGreaterThan, pt.DataFields(1), 10

'    pf.CurrentPage = "United States"

    pf.EnableMultiplePageItems = True
'
    pf.PivotItems("United States").Visible = False
    pf.PivotItems("United Kingdom").Visible = False
    
    
End Sub























