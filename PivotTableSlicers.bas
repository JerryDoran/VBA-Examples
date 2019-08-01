Attribute VB_Name = "PivotTableSlicers"
Option Explicit

Sub CreatePivotTable()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsMovies.Name & "!" & wsMovies.Range("A1").CurrentRegion.Address)
    
    Set ws = Worksheets.Add
    ws.Name = "MoviePivot"
    ws.Range("A3").Select
    
    Set pt = pc.CreatePivotTable(TableDestination:=ActiveCell, TableName:="MoviePivot")
    
    pt.AddFields _
        RowFields:="Genre", _
        ColumnFields:="Certificate"
        
    pt.AddDataField _
        Field:=pt.PivotFields("Run Time"), _
        Function:=XlConsolidationFunction.xlAverage
        
    pt.DataFields(1).NumberFormat = "0.00"
    
End Sub

Sub AddSlicer()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim r As Range
    
    Set ws = Worksheets("MoviePivot")
    Set pt = ws.PivotTables("MoviePivot")
    
    On Error Resume Next
    ThisWorkbook.SlicerCaches("CountrySlicerCache").Delete
    
    'This statement resets the error handlers back to their normal state
    On Error GoTo 0
    
    Set sc = ThisWorkbook.SlicerCaches.Add2(pt, "Country", "CountrySlicerCache", XlSlicerCacheType.xlSlicer)
    
    Set sl = sc.Slicers.Add(ws, , "CountrySlicer", "Choose Countries")
    
    Set r = pt.TableRange1
    
    sl.Top = r.Top
    sl.Left = r.Left + r.Width + 20
    sl.Height = r.Height
    sl.Width = 150
    sl.Style = "SlicerStyleLight1"
    
    'sl.NumberOfColumns = 2
    'sl.ColumnWidth = 150
    
End Sub

Sub DeleteAllSlicersFromPivot()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sl As Slicer
    
    Set ws = Worksheets("MoviePivot")
    Set pt = ws.PivotTables("MoviePivot")
    
    'Loop over the slicers collection of a pivot table
    For Each sl In pt.Slicers
       sl.Delete
    Next sl
    
End Sub

Sub FilterPivotWithSlicer()

    Dim sc As SlicerCache
    Dim si As SlicerItem
    
    Set sc = ThisWorkbook.SlicerCaches("CountrySlicerCache")
        
'    sc.SlicerItems("United States").Selected = False
    
'    For Each si In sc.SlicerItems
'        If InStr(si.Name, " ") > 0 Then
'            si.Selected = False
'        Else
'            si.Selected = True
'        End If
'    Next si
    
    sc.ClearAllFilters
    
End Sub

Sub CreateAdditionalPivotTableUsingPivotCacheFromFirstPivotTableCreatedOnWorksheet()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Set ws = Worksheets("MoviePivot")
   
    ws.Range("A31").Select
    
    Set pc = ws.PivotTables("MoviePivot").PivotCache
    
    Set pt = pc.CreatePivotTable(TableDestination:=ActiveCell, TableName:="MoviePivot1")
    
    pt.AddFields _
        RowFields:="Distributor", _
        ColumnFields:="Certificate"
        
    pt.AddDataField _
        Field:=pt.PivotFields("Run Time"), _
        Function:=XlConsolidationFunction.xlAverage
        
    pt.DataFields(1).NumberFormat = "0.00"
    
End Sub

Sub ConnectSlicerToMultiplePivots()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache
    
    Set ws = Worksheets("MoviePivot")
    Set pt = ws.PivotTables("MoviePivot1")
    Set sc = ThisWorkbook.SlicerCaches("CountrySlicerCache")
    
    sc.PivotTables.AddPivotTable pt
    
    sc.SlicerItems("United States").Selected = False
    
End Sub

Sub DisconnectSlicerToMultiplePivots()
    
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache
    
    Set ws = Worksheets("MoviePivot")
    Set pt = ws.PivotTables("MoviePivot1")
    Set sc = ThisWorkbook.SlicerCaches("CountrySlicerCache")
    
    sc.PivotTables.RemovePivotTable pt
    
    sc.SlicerItems("United States").Selected = False
    
End Sub

    
    















































