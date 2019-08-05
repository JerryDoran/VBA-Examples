Attribute VB_Name = "QueryWebPages"
Option Explicit

Sub ImportWiseOwlCourses()

    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim URL As String
    
    URL = "http://www.wiseowl.co.uk/courses/"
    
    Set ws = Worksheets.Add
    
    Set qt = ws.QueryTables.Add(Connection:="URL;" & URL, Destination:=Range("A1"))
    
    With qt
        .RefreshOnFileOpen = True       'refreshes the query table each time the file is opened
        '.RefreshPeriod = 1             'automatically refreshes the table after 1 minute
        .Name = "WOLCourses"
        .WebFormatting = xlWebFormattingRTF
        '.WebSelectionType = xlAllTables
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "3"                '<--Can set a comma separated list of the number of tables you want returned.
        
        .Refresh
    End With
    
End Sub

Sub ImportWiseOwlCoursesWithLoop()

    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim URL As String
    Dim i As Integer
    
    URL = "http://www.wiseowl.co.uk/courses/"
    
    For i = 1 To 2
        Set ws = Worksheets.Add
        
        Set qt = ws.QueryTables.Add(Connection:="URL;" & URL, Destination:=Range("A1"))
        
        With qt
            .RefreshOnFileOpen = True       'refreshes the query table each time the file is opened
            '.RefreshPeriod = 1             'automatically refreshes the table after 1 minute
            .Name = "WOLCourses"
            .WebFormatting = xlWebFormattingRTF
            '.WebSelectionType = xlAllTables
            .WebSelectionType = xlSpecifiedTables
            .WebTables = i
            .Refresh
        End With
    Next i
End Sub





























