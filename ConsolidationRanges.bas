Attribute VB_Name = "ConsolidationRanges"
Option Explicit

Sub CreateBasicConsolidatedPivotFixedDataRange()

    Dim ws As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlConsolidation, _
        SourceData:=Array( _
            Array("'2016 Data'!R1C1:R5C6", "2016"), _
            Array("'2015 Data'!R1C1:R6C3", "2015"), _
            Array("'2014 Data'!R1C1:R5C3", "2014")))
    
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A3"), _
        TableName:="ConsolidatedPivot")
    
End Sub

Sub CreateBasicConsolidatedPivotVariableDataRange()

    Dim ws As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlConsolidation, _
        SourceData:=Array( _
            Array("'" & ws2016.Name & "'!" & ws2016.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1), ws2016.Name), _
            Array("'" & ws2015.Name & "'!" & ws2015.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1), ws2015.Name), _
            Array("'" & ws2014.Name & "'!" & ws2014.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1), ws2014.Name)))
    
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A3"), _
        TableName:="ConsolidatedPivot")
    
End Sub

Sub CreateBasicConsolidatedPivotWithArray()

    Dim ws As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ar(1 To 3, 1 To 2) As String
    
    ar(1, 1) = "'" & ws2016.Name & "'!" & ws2016.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)
    ar(1, 2) = ws2016.Name
    
    ar(2, 1) = "'" & ws2015.Name & "'!" & ws2015.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)
    ar(2, 2) = ws2015.Name
    
    ar(3, 1) = "'" & ws2014.Name & "'!" & ws2014.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)
    ar(3, 2) = ws2014.Name
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlConsolidation, _
        SourceData:=ar)
    
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A3"), _
        TableName:="ConsolidatedPivot")
    
End Sub

Sub DeleteAllButDataSheets()

    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Name Like "*Data" Then ws.Delete
    Next ws
    
End Sub

Sub CreateBasicConsolidatedPivotWithDynamicArray()

    Dim ws As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim ar() As String
    Dim i As Integer
    
    DeleteAllButDataSheets
    
    ReDim ar(1 To Worksheets.Count, 1 To 2)
    
    For i = 1 To Worksheets.Count
        ar(i, 1) = "'" & Worksheets(i).Name & "'!" & Worksheets(i).Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)
        ar(i, 2) = Worksheets(i).Name
    Next i
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlConsolidation, _
        SourceData:=ar)
    
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A3"), _
        TableName:="ConsolidatedPivot")
    
    Set pf = pt.DataFields(1)
    
    pf.Function = xlAverage
    pf.Caption = "Average run time"
    pf.NumberFormat = "0.00"
    
End Sub




















