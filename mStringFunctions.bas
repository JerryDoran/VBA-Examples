Attribute VB_Name = "mStringFunctions"
Option Explicit

Sub BasicStringSplittingFunctions()

    Dim s As String
    
    s = Range("A15").Value
    
    Debug.Print s
    Debug.Print Left(s, 1)
    'Debug.Print Left$(s, 1)
    'Debug.Print LeftB(s, 6)     'returns number of bytes of data from each string
    
    Debug.Print Right(s, 6)
    Debug.Print Mid(s, 3, 2)
    
       
End Sub
Sub FindingCharacterPositions()

    Dim s As String
    Dim FirstSpace As Long
    
    s = Range("C5").Value
    
    Debug.Print s
'    Debug.Print InStr(1, s, " ")
    FirstSpace = InStr(1, s, " ")       '<--Returns the position of the first space in the string
    Debug.Print Left(s, FirstSpace - 1)
    'Debug.Print Right(s, Len(s) - FirstSpace)
    Debug.Print Mid(s, FirstSpace + 1)
'    Debug.Print Left(s, 3)
'    Debug.Print Right(s, 5)
       
End Sub

Sub FindingCharacterPositionsLastSpace()

    Dim s As String
    Dim FirstSpace As Long
    Dim LastSpace As Long
    
    s = Range("C893").Value
    
    Debug.Print s
    
    If s Like "* *" Then        '<--Checks to see if there is a space in the string
        FirstSpace = InStr(1, s, " ")
        LastSpace = InStrRev(s, " ")        '<--Looks at the last space in a string by starting at the end of the string and move forward
        Debug.Print Left(s, LastSpace - 1)
        'Debug.Print Right(s, Len(s) - FirstSpace)
        Debug.Print Mid(s, LastSpace + 1)
    Else
        Debug.Print s
    End If
          
End Sub

Sub LoopToSplitText()

    Dim s As String
    Dim ThisSpace As Long
    Dim LastSpace As Long
    
    s = Range("A11").Value
    
    ThisSpace = InStr(1, s, " ")
    
    If ThisSpace = 0 Then
        Debug.Print s
    Else
        Do While ThisSpace > 0
            Debug.Print Mid(s, LastSpace + 1, ThisSpace - LastSpace)
            LastSpace = ThisSpace
            ThisSpace = InStr(ThisSpace + 1, s, " ")
        Loop
        Debug.Print Mid(s, LastSpace + 1)
    End If
    
End Sub

Sub EasyWayToSplitText()

    Dim s As String
    Dim arr() As String
    Dim v As Variant
    
    s = Range("A11").Value
    
    arr = Split(s, " ")
    
    For Each v In arr
        Debug.Print v
    Next v
    
End Sub







































