Attribute VB_Name = "GetHTML"
Option Explicit

Sub GetHTMLDocument()

    Dim IE As SHDocVw.InternetExplorerMedium
    'Dim IE As Object
    Dim HTMLDoc As MSHTML.HTMLDocument
    Dim HTMLInput As MSHTML.IHTMLElement        'Holds a reference to any individual item on a web page
    Dim HTMLButtons As MSHTML.IHTMLElementCollection
    Dim HTMLButton As MSHTML.IHTMLElement
    
    'Set IE = CreateObject("InternetExplorer.Application")
    Set IE = New InternetExplorerMedium
        
    IE.Navigate "wiseowl.co.uk"
    IE.Visible = True
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE

    Loop

    Set HTMLDoc = IE.Document
    Set HTMLInput = HTMLDoc.getElementById("what")
    HTMLInput.Value = "Excel VBA"
    
    Set HTMLButtons = HTMLDoc.getElementsByTagName("button")
    
'    For Each HTMLButton In HTMLButtons
'        Debug.Print HTMLButton.className, HTMLButton.tagName, HTMLButton.ID, HTMLButton.innerText
'    Next HTMLButton
    
    HTMLButtons(0).Click
    
    
End Sub
