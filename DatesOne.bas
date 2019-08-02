Attribute VB_Name = "DatesOne"
Option Explicit

Sub DateBasics()

    Dim dt As Date
    
    dt = #9/1/2017#
    
    Debug.Print dt
    MsgBox dt
    Sheet1.Range("A1").Value = dt
    
End Sub

Sub DateRanges()

    Dim dMin As Date
    Dim dMax As Date
    Dim dMinExcel As Date
    
    dMin = #1/1/100#
    dMax = #12/31/9999#
    dMinExcel = #1/1/1900#
    
    MsgBox dMin & vbNewLine & dMax & vbNewLine & dMinExcel
    
    Worksheets.Add
    Range("A1").Value = dMin
    Range("A2").Value = dMax
    Range("A3").Value = dMinExcel
    
End Sub

Sub LeapYearBug()

    Dim dt As Date
    
    Worksheets.Add
    
    dt = #3/1/1900#
    Range("A1").Value = dt
    
    dt = #2/28/1900#
    Range("A2").Value = dt
    
End Sub

Sub CurrentDate()

    Dim dt As Date
    
    dt = Date
    
    Debug.Print dt
    Range("A5").Value = dt
    
    dt = Now
    Debug.Print dt
    Range("A8").Value = dt
    
End Sub

Sub EnterDateFunctions()

    Range("A10").Value = "=Today()"
    Range("A11").Value = "=Now()"
    Range("A11").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
End Sub

































