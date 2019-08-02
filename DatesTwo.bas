Attribute VB_Name = "DatesTwo"
Option Explicit

Sub FormattingDates()

    Dim dt As Date
    
    dt = Date
    
    Worksheets.Add
    
    Range("A1").Value = FormatDateTime(dt, vbGeneralDate)
    Range("A2").Value = FormatDateTime(dt, vbLongDate)
    Range("A3").Value = FormatDateTime(dt, vbShortDate)
    
    Range("A4").Value = Format(dt, "dddd d mmmm yyyy")      'Be careful - stores value in a cell as string.  Save for reports...not calculations
    Range("A5").Value = Format(dt, "dd/mm/yyyy")
    
End Sub
Sub DateParts()

    Dim dt As Date
    
    dt = Date
    
    Worksheets.Add
    
    Range("A1").Value = Year(dt)
    Range("A2").Value = Month(dt)
    Range("A3").Value = Day(dt)
    
    Range("A4").Value = DateSerial(Range("A1").Value, Range("A2").Value, Range("A3").Value)
    
    
End Sub
