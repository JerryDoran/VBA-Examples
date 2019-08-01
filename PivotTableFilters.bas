Attribute VB_Name = "PivotTableFilters"
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
        RowFields:="Release Date", _
        ColumnFields:="Certificate"
        
    pt.AddDataField _
        Field:=pt.PivotFields("Run Time"), _
        Function:=XlConsolidationFunction.xlAverage
        
    pt.DataFields(1).NumberFormat = "0.00"
    
    'pt.RowAxisLayout xlTabularRow
    'pt.RowAxisLayout xlOutlineRow
    pt.RowAxisLayout xlCompactRow
    
    pt.CompactLayoutRowHeader = "Dates"
    pt.CompactLayoutColumnHeader = "Certificates"
    
End Sub

Sub GroupingDates()

    'Array Elements for Periods parameter
    '1 = Seconds
    '2 = Minutes
    '3 = Hours
    '4 = Days
    '5 = Months
    '6 = Quarters
    '7 = Years

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim r As Range
    Dim pf As PivotField
    
    Set ws = Worksheets("MoviePivot")
    
    Set pt = ws.PivotTables("MoviePivot")
    
    Set r = pt.RowRange.Cells(2, 1)
    
    r.Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)
    
    Set pf = pt.PivotFields("Years")
    pf.AutoSort xlDescending, "Years"
    
    Set pf = pt.PivotFields("Release Date")
    pf.AutoSort xlAscending, "Release Date"
    
    'r.Ungroup
    
    
End Sub

Sub FilterDates()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    
    Set ws = Worksheets("MoviePivot")
    
    Set pt = ws.PivotTables("MoviePivot")
    
    Set pf = pt.PivotFields("Release Date")
    
    pf.ClearAllFilters
'    pf.PivotFilters.Add2 Type:=xlAfter, Value1:="12/31/1999"
'    pf.PivotFilters.Add2 Type:=xlDateBetween, Value1:="01/01/1999", Value2:="12/31/1999"
'    pf.PivotFilters.Add2 Type:=xlDateBetween, Value1:=DateAdd("yyyy", -2, Date), Value2:=Date
    
End Sub

Sub CreateTimeline()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache
    Dim sl As Slicer
    
    Set ws = Worksheets("MoviePivot")
    
    Set pt = ws.PivotTables("MoviePivot")
    
    On Error Resume Next
    ThisWorkbook.SlicerCaches("DateSlicerCache").Delete
    On Error GoTo 0
    
    Set sc = ThisWorkbook.SlicerCaches.Add2(pt, "Release Date", "DateSlicerCache", XlSlicerCacheType.xlTimeline)
    
    Set sl = sc.Slicers.Add(ws, , "DateSlicer", "Select Date Range")
    
End Sub

Sub DeleteSlicersFromPivot()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sl As Slicer
    
    Set ws = Worksheets("MoviePivot")
    
    Set pt = ws.PivotTables("MoviePivot")
    
    For Each sl In pt.Slicers
        sl.Delete
    Next sl
End Sub

Sub FormatTimeline()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sl As Slicer
    
    Set ws = Worksheets("MoviePivot")
    Set pt = ws.PivotTables("MoviePivot")
    Set sl = pt.Slicers("DateSlicer")
    
    sl.Top = pt.TableRange1.Top
    sl.Left = pt.TableRange1.Left + pt.TableRange1.Width + 20
    sl.Width = 500
    
    sl.Style = "TimeSlicerStyleLight5"
    sl.TimelineViewState.Level = xlTimelineLevelYears
'    sl.TimelineViewState.Level = xlTimelineLevelMonths
    
End Sub

Sub FilterDatesWithTimeline()

    Dim sc As SlicerCache
    Dim r As Range
    
    Set sc = ThisWorkbook.SlicerCaches("DateSlicerCache")
    
    'sc.TimelineState.SetFilterDateRange StartDate:="1/1/2000", EndDate:="12/31/2000"
    'sc.TimelineState.SetFilterDateRange StartDate:=Range("F1"), EndDate:=Range("F2")
    sc.TimelineState.SetFilterDateRange StartDate:=DateAdd("yyyy", -3, Date), EndDate:=Date
    'sc.ClearAllFilters
    
End Sub

Sub CreateNewPivotTable()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Set ws = Worksheets("MoviePivot")
    ws.Range("J12").Select
    
    Set pc = ws.PivotTables("MoviePivot").PivotCache
    
    Set pt = pc.CreatePivotTable(TableDestination:=ActiveCell, TableName:="MoviePivot1")
    
    pt.AddFields _
        RowFields:="Genre", _
        ColumnFields:="Certificate"
        
    pt.AddDataField _
        Field:=pt.PivotFields("Run Time"), _
        Function:=XlConsolidationFunction.xlAverage
        
    pt.DataFields(1).NumberFormat = "0.00"
    
    'pt.RowAxisLayout xlTabularRow
    'pt.RowAxisLayout xlOutlineRow
    pt.RowAxisLayout xlCompactRow
    
    pt.CompactLayoutRowHeader = "Genre"
    pt.CompactLayoutColumnHeader = "Certificates"
    
End Sub
Sub ConnectTimelineToPivots()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache
    
    Set ws = Worksheets("MoviePivot")
    Set pt = ws.PivotTables("MoviePivot1")
    Set sc = ThisWorkbook.SlicerCaches("DateSlicerCache")
    
    sc.PivotTables.AddPivotTable pt
    
    sc.TimelineState.SetFilterDateRange StartDate:=DateAdd("yyyy", -3, Date), EndDate:=Date
    
End Sub



















































































