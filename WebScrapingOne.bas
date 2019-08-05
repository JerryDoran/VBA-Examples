Attribute VB_Name = "WebScrapingOne"
Option Explicit

Sub ProcessHTMLPage(HTMLPage As MSHTML.HTMLDocument)

    Dim HTMLTable As MSHTML.IHTMLElement
    Dim HTMLTables As MSHTML.IHTMLElementCollection
    Dim HTMLRow As MSHTML.IHTMLElement
    Dim HTMLCell As MSHTML.IHTMLElement
    Dim RowNum As Long
    Dim ColNum As Integer
    
    Set HTMLTables = HTMLPage.getElementsByTagName("table")
    
    For Each HTMLTable In HTMLTables
    
        Worksheets.Add
        
        Range("A1").Value = HTMLTable.className
        Range("B1").Value = Now
        
        RowNum = 2
        
        For Each HTMLRow In HTMLTable.getElementsByTagName("tr")
            'Debug.Print vbTab & HTMLRow.innerText
            
            ColNum = 1
            For Each HTMLCell In HTMLRow.Children
                Cells(RowNum, ColNum) = HTMLCell.innerText
                ColNum = ColNum + 1
            Next HTMLCell
            RowNum = RowNum + 1
        Next HTMLRow
        
    Next HTMLTable
    
End Sub
