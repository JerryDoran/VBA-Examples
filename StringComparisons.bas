Attribute VB_Name = "StringComparisons"
Option Explicit
'Option Compare Text     'makes all string comparisons in this module case insensitive
'Option Compare Binary   'makes all string comparisons in ths module case sensitive

Sub StringComparisons()

    Dim s1 As String
    Dim s2 As String
    
    s1 = "a"
    s2 = "A"
    
    'If s1 = s2 Then Debug.Print "the same" Else Debug.Print "different"
    'If UCase(s1) = UCase(s2) Then Debug.Print "the same" Else Debug.Print "different"
    If LCase(s1) = LCase(s2) Then Debug.Print "the same" Else Debug.Print "different"

End Sub
Sub WildCardComparisons()

    Dim r As Range
    
    For Each r In Range("A2", Range("A1").End(xlDown))
       'If LCase(r.Value) = "king kong" Then
       'If LCase(Left(r.Value, 1)) = "k" Then
       'If LCase(r.Value) Like "k*" Then
       'If LCase(r.Value) Like "*k" Then
       'If LCase(r.Value) Like "*king*" Then
       'If Not LCase(r.Value) Like "*twilight*" Then
       'If Not LCase(r.Value) Like "*twilight*" And LCase(r.Value) Like "*king*" Then
       'If LCase(r.Value) Like "* ?" Then    'gets all films with only a single character at the end!
       'If LCase(r.Value) Like "* ???" Then   'gets all films with three characters at the end!
       'If LCase(r.Value) Like "* #" Then      'Only films with a single digit at the end.
       'If LCase(r.Value) Like "*[?]" Then      'Only films who's last character at the end is a "?".
       'If LCase(r.Value) Like "a*" Or LCase(r.Value) Like "m*" Or LCase(r.Value) Like "g*" Then
'       If LCase(r.Value) Like "[amg]*" Then     'gets all films that start with a,m or g
       'If LCase(r.Value) Like "[j-m]*" Then     'gets all films that start with j,k,l,m as long as they are continuous
       'If LCase(r.Value) Like "[!j-m]*" Then     'gets all films that do not start with j,k,l,m as long as they are continuous
       If LCase(r.Value) Like "? [!h]*'s ????" Then    'gets all films that begin with a single character and the first character of the second word is not a h and find 's and the last word of the film title must contain four characters.
            Debug.Print r.Value & ", " & r.Offset(0, 1).Value
       End If
    Next r
    
End Sub





















