Attribute VB_Name = "WebScrapingTwo"
Option Explicit

Sub GetExchangeRates(FromCurrency As String, Amount As Double)

    'Create two autoinstancing variables
    Dim XMLPage As New MSXML2.XMLHTTP60
    Dim HTMLDoc As New MSHTML.HTMLDocument
    Dim URL As String
    
    URL = "http://x-rates.com/table/?from=" & FromCurrency & "&amount=" & Amount
        
    XMLPage.Open "GET", URL, False
    XMLPage.send
    
    HTMLDoc.body.innerHTML = XMLPage.responseText
    
    ProcessHTMLPage HTMLDoc
    
End Sub

Sub OpenRatesForm()

    RatesForm.Show

End Sub
