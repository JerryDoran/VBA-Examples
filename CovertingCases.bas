Attribute VB_Name = "CovertingCases"
Option Explicit

Sub ConvertingCase()

    Dim s As String
    
    s = Range("A11").Value
    
    Debug.Print s
    
'    Debug.Print LCase(s)
'    Debug.Print StrConv(s, vbLowerCase)
'
'    Debug.Print UCase(s)
'    Debug.Print StrConv(s, vbUpperCase)
    
'    Debug.Print WorksheetFunction.Proper(s)
'    Debug.Print StrConv(s, vbProperCase)

    s = "Harry Potter and the Order of the Phoenix. Harry Potter and the Goblet of Fire. Harry Potter and the Prisoner of Azkaban."
    
'    Debug.Print UCase(Left(s, 1)) & LCase(Mid(s, 2))
    Debug.Print SentenceCase(s)
    
End Sub

Function SentenceCase(s As String) As String

    Dim Sentences() As String
    Dim i As Integer
    
    Sentences = Split(s, ".")
    
    For i = LBound(Sentences) To UBound(Sentences)
        If Sentences(i) <> "" Then
            s = Trim(Sentences(i))
            s = UCase(Left(s, 1)) & LCase(Mid(s, 2))
            Sentences(i) = s
        End If
    Next i
    
    s = Trim(Join(Sentences, ". "))
    
    SentenceCase = s
    
End Function

Function ToggleCase(s1 As String) As String

    Dim s2 As String
    Dim c As String * 1 'Fixed length string - one character
    Dim n As Long
    
    For n = 1 To Len(s1)
        c = Mid(s1, n, 1)   '<--Will store one character at a time of the string into the c variable
        
        If StrComp(c, UCase(c)) = 0 Then        '<--Will return 0 if the comparison is equal (that the character is upper case)
            c = LCase(c)
        Else
            c = UCase(c)
        End If
        
        s2 = s2 & c
        
    Next n
   
   ToggleCase = s2
    
End Function

























































