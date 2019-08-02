Attribute VB_Name = "DatesThree"
Option Explicit

Sub BasicDateCalculations()

    Dim StartDate As Date
    Dim EndDate As Date
    
    StartDate = Date
    EndDate = Sheet1.Range("A1").Value
    
    Range("A2").Value = EndDate - StartDate
    Range("A3").Value = DateDiff("d", StartDate, EndDate)
    Range("A4").Value = DateDiff("ww", StartDate, EndDate)
    Range("A5").Value = DateDiff("m", StartDate, EndDate)
    Range("A6").Value = WorksheetFunction.NetworkDays(StartDate, EndDate)
    
End Sub
Sub CalculateYearsDifference()

    Dim StartDate As Date
    Dim EndDate As Date
    
    StartDate = #10/5/1990#
    EndDate = Date
    
    Debug.Print DateDiff("yyyy", StartDate, EndDate)
    Debug.Print Evaluate("datedif(A1,Today(),""Y"")")
    Debug.Print AgeInYears(StartDate, #10/5/2002#)
    
End Sub
Function AgeInYears(ByVal StartDate As Date, Optional ByVal EndDate As Date) As Integer

    Dim YearsDiff As Integer
    Dim Anniversary As Date
    
    If EndDate = 0 Then
        EndDate = Date
    End If
    
    YearsDiff = DateDiff("yyyy", StartDate, EndDate)
    
    Anniversary = DateAdd("yyyy", YearsDiff, StartDate)
    
    If Anniversary > EndDate Then YearsDiff = YearsDiff - 1
    
    AgeInYears = YearsDiff
    
End Function























