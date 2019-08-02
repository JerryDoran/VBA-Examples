Attribute VB_Name = "PivotCharts"
Option Explicit

Sub CreatePivotTable()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsMovies.Name & "!" & wsMovies.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)
        
    Set ws = Worksheets.Add
    ws.Name = "MovieTable"
    
    Range("A3").Select
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ActiveCell, _
        TableName:="MoviePivot")
        
    pt.AddFields _
        RowFields:="Country", _
        ColumnFields:="Certificate", _
        PageFields:="Studio"
        
    pt.AddDataField _
        Field:=pt.PivotFields("Run Time"), _
        Function:=XlConsolidationFunction.xlAverage
        
    pt.DataFields(1).NumberFormat = "0.00"
    
End Sub

Sub CreatePivotChartEmbedded()

    Dim sh As Shape
    Dim ws As Worksheet
    Dim ch As Chart
    Dim pt As PivotTable
    
    DeleteAllChartObjects
    
    Set ws = Worksheets("MovieTable")
    Set sh = ws.Shapes.AddChart2( _
        XlChartType:=XlChartType.xlColumnStacked, _
        Width:=500, Height:=400)
    
    Set ch = sh.Chart
    Set pt = ws.PivotTables("MoviePivot")
    
    'TableRange2 includes the page field.  TableRange1 excludes the page field
    ch.SetSourceData pt.TableRange1
    
    sh.Top = pt.TableRange1.Top
    sh.Left = pt.TableRange1.Left + pt.TableRange1.Width + 10
    
End Sub

Sub DeleteAllChartObjects()

    Dim co As ChartObject
    
    For Each co In Worksheets("MovieTable").ChartObjects
        co.Delete
    Next co
    
End Sub

Sub CreatePivotChartSheet()

    Dim ch As Chart
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Set ch = Charts.Add
    
    ch.Name = "MovieChart"
    
    Set ws = Worksheets("MovieTable")
    Set pt = ws.PivotTables("MoviePivot")
    
    ch.SetSourceData pt.TableRange2
    
End Sub

Sub EditPivotChart()

    Dim pt As PivotTable
    Dim ch As Chart
    Dim pf As PivotField
    
    'Get reference to the pivot table
    'Set pt = Worksheets("MovieTable").PivotTables("MoviePivot")
    
    'Get reference to the pivot chart
    Set ch = Charts("MovieChart")
    Set pt = ch.PivotLayout.PivotTable
    
    For Each pf In pt.VisibleFields
        pf.Orientation = xlHidden
    Next pf
    
    pt.AddFields _
        RowFields:="Studio", _
        ColumnFields:="Genre"

    pt.AddDataField _
        Field:=pt.PivotFields("Budget ($)"), _
        Function:=XlConsolidationFunction.xlAverage
    
End Sub

Sub FilterSpecificItems()

    Dim pt As PivotTable
    Dim ch As Chart
    Dim pf As PivotField
    
    'Get reference to the pivot chart
    Set ch = Charts("MovieChart")
    Set pt = ch.PivotLayout.PivotTable
    Set pf = pt.PivotFields("Genre")
    
    pf.PivotItems("Action").Visible = False
    pf.PivotItems("Romance").Visible = False
    
    pf.ClearAllFilters
    
End Sub

Sub FilterBasedOnCriteria()

    Dim pt As PivotTable
    Dim ch As Chart
    Dim pf As PivotField
    Dim pi As PivotItem
    
    'Get reference to the pivot chart
    Set ch = Charts("MovieChart")
    Set pt = ch.PivotLayout.PivotTable
    Set pf = pt.PivotFields("Studio")
    
    For Each pi In pf.PivotItems
        If Len(pi.Name) > 12 Then
            pi.Visible = False
        Else
            pi.Visible = True
        End If
    Next pi
    
    pf.ClearAllFilters
    
End Sub

Sub FormattingChart()

    Dim ch As Chart
    Dim pt As PivotTable
    Dim TitleText As String
    
    Set ch = Charts("MovieChart")
    
    ch.ChartType = xlColumnStacked
    ch.ChartStyle = 48
    ch.ApplyLayout 9
    ch.ChartColor = 13
    
    Set pt = ch.PivotLayout.PivotTable
    
    ch.HasTitle = True
    TitleText = pt.DataFields(1).Caption & " by " & pt.RowFields(1).Name & " and " & pt.ColumnFields(1).Name
    
    ch.ChartTitle.Text = TitleText
        
End Sub







































































