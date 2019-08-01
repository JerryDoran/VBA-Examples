Attribute VB_Name = "AccessSQL"
Option Explicit

Sub CreateImplicitConnectionAndPivotFromAccessSQL()

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cnString As String
    
    cnString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\!Excel VBA\Data\Movies.accdb;Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False"
    
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlExternal)
    
    pc.Connection = cnString
    pc.CommandType = xlCmdSql
    pc.CommandText = "SELECT tblFilm.FilmName, tblFilm.FilmReleaseDate, tblFilm.FilmRunTimeMinutes, tblFilm.FilmBudgetDollars, tblFilm.FilmBoxOfficeDollars, tblFilm.FilmOscarNominations, tblFilm.FilmOscarWins, tblCertificate.CertificateName, tblCountry.CountryName, tblDirector.DirectorName, tblStudio.StudioName, tblLanguage.LanguageName FROM tblStudio INNER JOIN (tblLanguage INNER JOIN (tblDirector INNER JOIN (tblCountry INNER JOIN (tblCertificate INNER JOIN tblFilm ON tblCertificate.CertificateID = tblFilm.FilmCertificateID) ON tblCountry.CountryId = tblFilm.FilmCountryID) ON tblDirector.DirectorId = tblFilm.FilmDirectorID) ON tblLanguage.LanguageId = tblFilm.FilmLanguageID) ON tblStudio.StudioId = tblFilm.FilmStudioID;"
        
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
Sub CreateAccessSQL(SQLString As String, Optional DataField As String, Optional RowField As String, Optional ColField As String, Optional PageField As String, Optional PivotFunction As XlConsolidationFunction = xlSum)

    Dim pc As PivotCache
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim cnString As String
    
    cnString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=C:\!Excel VBA\Data\Movies.accdb;Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False"
    
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlExternal)
    
    pc.Connection = cnString
    pc.CommandType = xlCmdSql
    pc.CommandText = SQLString
        
    Set ws = Worksheets.Add
    
    Set pt = pc.CreatePivotTable(Tabledestination:=ws.Range("A3"), TableName:="MoviePivot")
    
    On Error GoTo WrongName
    
    If RowField <> "" Then
        Set pf = pt.PivotFields(RowField)
        pf.Orientation = xlRowField
    End If
    
    If ColField <> "" Then
        Set pf = pt.PivotFields(ColField)
        pf.Orientation = xlColumnField
    End If
    
    If PageField <> "" Then
        Set pf = pt.PivotFields(PageField)
        pf.Orientation = xlPageField
    End If
    
    If DataField <> "" Then
        Set pf = pt.PivotFields(DataField)
        
        'If PivotFunction = 0 Then PivotFunction = xlSum
            
        pt.AddDataField _
        Field:=pf, _
        Function:=PivotFunction
    End If
    
    Exit Sub
    
WrongName:
    MsgBox "Wrong column name", vbExclamation
            
End Sub

Sub TestSQLQueries()

    Dim str As String
    
    DeleteAllButMoviesSheet
    DeleteAllConnections
    
'    str = "SELECT FilmCountryID, FilmCertificateID, FilmStudioID, FilmRunTimeMinutes " & _
'          "FROM tblFilm " & _
'          "WHERE FilmOscarWins >=1 AND FilmReleaseDate BETWEEN #01 Jan 2000# AND #31 DEC 2000#"

'    str = "SELECT FilmCountryID, FilmCertificateID, FilmStudioID, FilmRunTimeMinutes " & _
'           "FROM tblFilm " & _
'           "WHERE FilmName LIKE '%king%'"

    str = "SELECT CountryName, CertificateName, StudioName, FilmRunTimeMinutes " & _
          "FROM tblStudio " & _
          "INNER JOIN (tblCertificate INNER JOIN (tblFilm INNER JOIN tblCountry ON tblFilm.FilmCountryID = tblCountry.CountryID) " & _
          "ON tblFilm.FilmCertificateID = tblCertificate.CertificateID) ON tblFilm.FilmStudioID = tblStudio.StudioID"
               
    CreateAccessSQL str, "FilmRunTimeMinutes", "StudioName", "CertificateName", "CountryName", xlCount
    
End Sub














































