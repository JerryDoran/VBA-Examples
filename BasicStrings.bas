Attribute VB_Name = "BasicStrings"
Option Explicit

Sub BasicStrings()

    Dim s As String
    'Dim s As String * 5     ' declaring a fixed length string - very rare to declare string variables like this.
        
    s = "anything you like up to about 2.147 billion characters"

    s = ThisWorkbook.Path
    s = Application.Version
    s = ActiveSheet.Name
    s = ActiveCell.Address
    
    'Using implicit type conversion
'    s = 123
'    s = #2/19/2017#
'
'    s = ActiveCell.Row
'    s = ThisWorkbook.Sheets.Count
'    s = Date
'
'    'explicit type conversion
'    s = CStr(activcell.Row)
'    s = CStr(ThisWorkbook.Sheets.Count)
'    s = CStr(Date)

    s = IIf(IsNull(Null), "", "reference to a field")
    
End Sub
