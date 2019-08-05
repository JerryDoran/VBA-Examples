Attribute VB_Name = "ExchangeRatesWithQueryString"
Option Explicit

Sub BrowseToExchangeRatesWithQueryString()

    Dim IE As SHDocVw.InternetExplorerMedium
    Dim HTMLDoc As MSHTML.HTMLDocument
   
    Set IE = New InternetExplorerMedium
        
    IE.Navigate "x-rates.com/table/?from=GBP&amount=3"
    IE.Visible = True
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE

    Loop

    Set HTMLDoc = IE.Document
    
    ProcessHTMLPage HTMLDoc
    
End Sub

Sub BrowseToExchangeRatesWithQueryStringAndXML()

    'Create two auto-instancing variables
    Dim XMLPage As New MSXML2.XMLHTTP60
    Dim HTMLDoc As New MSHTML.HTMLDocument
        
    XMLPage.Open "GET", "http://x-rates.com/table/?from=GBP&amount=3", False
    XMLPage.send
    
    HTMLDoc.body.innerHTML = XMLPage.responseText       '<--Creates new HTML document and sets its body equal to the response text of the XMLPage
    
    ProcessHTMLPage HTMLDoc
    
End Sub








































