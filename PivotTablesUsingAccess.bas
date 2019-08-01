Attribute VB_Name = "PivotTablesUsingAccess"
Option Explicit

Sub CreateBasicPivot()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsMovies.Name & "!" & wsMovies.Range("A1").CurrentRegion.Address)
        
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable(Tabledestination:=ws.Range("A3"), TableName:="MoviePivot")
    
    pt.AddFields _
        RowFields:="Studio", _
        ColumnFields:="Certificate", _
        PageFields:="Country"
        
    pt.AddDataField _
        Field:=pt.PivotFields("Run Time"), _
        Function:=XlConsolidationFunction.xlAverage
    
    pt.DataFields(1).NumberFormat = "0.00"
     
End Sub
Sub CreateBasicPivotFromAccess()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cn As WorkbookConnection
    
    Set cn = ThisWorkbook.Connections("Movies")
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlExternal, _
        SourceData:=cn)
        
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable(Tabledestination:=ws.Range("A3"), TableName:="MoviePivot")
    
    pt.AddFields _
        RowFields:="FilmStudioID", _
        ColumnFields:="FilmCertificateID", _
        PageFields:="FilmCountryID"
        
    pt.AddDataField _
        Field:=pt.PivotFields("FilmRunTimeMinutes"), _
        Function:=XlConsolidationFunction.xlAverage
    
    pt.DataFields(1).NumberFormat = "0.00"
        
End Sub

Sub CreateConnectionAndPivotFromAccess()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cn As WorkbookConnection
    Dim cnString As String
    
    cnString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\!Excel VBA\Data\Movies.accdb;Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False"
    
    Set cn = ThisWorkbook.Connections.Add2( _
        Name:="AccessFilmTable", _
        Description:="", _
        ConnectionString:=cnString, _
        CommandText:="tblFilm", _
        LCmdType:=XlCmdType.xlCmdTable)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlExternal, _
        SourceData:=cn)
        
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable(Tabledestination:=ws.Range("A3"), TableName:="MoviePivot")
    
    pt.AddFields _
        RowFields:="FilmStudioID", _
        ColumnFields:="FilmCertificateID", _
        PageFields:="FilmCountryID"
        
    pt.AddDataField _
        Field:=pt.PivotFields("FilmRunTimeMinutes"), _
        Function:=XlConsolidationFunction.xlAverage
    
    pt.DataFields(1).NumberFormat = "0.00"
        
End Sub

Sub CreateImplicitConnectionAndPivotFromAccess()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    'Dim cn As WorkbookConnection
    Dim cnString As String
    
    cnString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\!Excel VBA\Data\Movies.accdb;Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False"
    
    'Create the pivot cache
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlExternal)
    
    'Set properties of the pivot cache object
    pc.Connection = cnString
    pc.CommandType = xlCmdTable
    pc.CommandText = "tblFilm"
        
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable(Tabledestination:=ws.Range("A3"), TableName:="MoviePivot")
    
    pt.AddFields _
        RowFields:="FilmStudioID", _
        ColumnFields:="FilmCertificateID", _
        PageFields:="FilmCountryID"
        
    pt.AddDataField _
        Field:=pt.PivotFields("FilmRunTimeMinutes"), _
        Function:=XlConsolidationFunction.xlAverage
    
    pt.DataFields(1).NumberFormat = "0.00"
        
End Sub

Sub GetReferenceToConnection()

    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim cn As WorkbookConnection
   
    Set pt = ActiveSheet.Range("A3").PivotTable
    Set pc = pt.PivotCache
    Set cn = pc.WorkbookConnection
    
    With cn.OLEDBConnection
        
        .CommandText = "tblActor"
        .CommandType = xlCmdTable
        
    End With
    
    cn.Refresh
    
    pt.AddFields "ActorGenderID"
    pt.AddDataField pt.PivotFields("ActorID"), Function:=XlConsolidationFunction.xlCount
    
End Sub

Sub CreateImplicitConnectionAndPivotFromAccessQuery()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cnString As String
    
    cnString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\!Excel VBA\Data\Movies.accdb;Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False"
    
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlExternal)
    
    pc.Connection = cnString
    pc.CommandType = xlCmdTable
    pc.CommandText = "qryMovies"
        
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable(Tabledestination:=ws.Range("A3"), TableName:="MoviePivot")
    
    pt.AddFields _
        RowFields:="StudioName", _
        ColumnFields:="CertificateName", _
        PageFields:="CountryName"
        
    pt.AddDataField _
        Field:=pt.PivotFields("FilmRunTimeMinutes"), _
        Function:=XlConsolidationFunction.xlAverage
    
    pt.DataFields(1).NumberFormat = "0.00"
        
End Sub








































