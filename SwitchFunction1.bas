Attribute VB_Name = "SwitchFunction1"
Option Explicit

Sub BasicSelectCase()

    Dim RunTime As Integer
    Dim Length As String
    
    Range("A4").Select
    RunTime = ActiveCell.Offset(0, 3).Value
    
    Select Case RunTime
        Case Is <= 90
            Length = "Short"
        Case Is <= 120
            Length = "Medium"
        Case Is <= 150
            Length = "Long"
        Case Is <= 180
            Length = "Epic"
        Case Else
            Length = "Too Long"
    End Select
    
    Debug.Print RunTime, Length
End Sub
Sub BasicSwitch()

    Dim RunTime As Integer
    Dim Length As String
    
    Range("A2").Select
    RunTime = ActiveCell.Offset(0, 3).Value
    
    Length = FilmLength(RunTime)
    
    Debug.Print RunTime, Length
End Sub

Function FilmLength(RunTime As Integer) As String

    FilmLength = Switch( _
        RunTime <= 90, "Short", _
        RunTime <= 120, "Medium", _
        RunTime <= 150, "Long", _
        RunTime <= 180, "Epic", _
        True, "Way Too Long")
        
    
    
End Function




























































