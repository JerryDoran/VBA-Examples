Attribute VB_Name = "MoreTypes"
Option Explicit

Type Address
    NameNumber As String
    Street As String
    Town As String
    County As String
    PostCode As String
End Type


Type Contact
    
    Title As String
    FirstName As String
    LastName As String
    DateOfBirth As Date
    
    HomeAddress As Address
    WorkAddress As Address
    
End Type

Sub TestContact()

    Dim c As Contact
    
    c.FirstName = "Jerry"
    c.HomeAddress.NameNumber = "3"
    
End Sub
