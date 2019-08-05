Attribute VB_Name = "BrowseToExchangeRates"
Option Explicit

Sub BrowseToExchangeRates()

    Dim IE As SHDocVw.InternetExplorer
    Dim HTMLDoc As MSHTML.HTMLDocument
    Dim HTMLInput As MSHTML.IHTMLElement        'Holds a reference to any individual item on a web page
    Dim HTMLTags As MSHTML.IHTMLElementCollection
    Dim HTMLTag As MSHTML.IHTMLElement
    
    'Set IE = CreateObject("InternetExplorer.Application")
    Set IE = New InternetExplorerMedium
        
    IE.Visible = True
    IE.Navigate "x-rates.com"
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE

    Loop

    Set HTMLDoc = IE.Document
    
    Set HTMLInput = HTMLDoc.getElementById("amount")
    HTMLInput.Value = 5
    
    Set HTMLInput = HTMLDoc.getElementById("from")
    HTMLInput.Value = "GBP"
    
    Set HTMLInput = HTMLDoc.getElementById("to")
    HTMLInput.Value = "USD"
    
    Set HTMLTags = HTMLDoc.getElementsByTagName("a")
    
    For Each HTMLTag In HTMLTags
        'Debug.Print HTMLTag.getAttribute("classname"), HTMLTag.getAttribute("href"), HTMLTag.getAttribute("rel")
        If HTMLTag.getAttribute("href") = "http://x-rates.com/table/" And HTMLTag.getAttribute("rel") = "ratestable" Then
            HTMLTag.Click
            Exit For
        End If
    Next HTMLTag
    
    'Debug.Print HTMLTags.Length        tells how many web page elements that were looped through on the web page (a - hyperlinks in this case)
    
End Sub
