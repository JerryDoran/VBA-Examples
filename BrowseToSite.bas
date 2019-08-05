Attribute VB_Name = "BrowseToSite"
Option Explicit

Sub BrowseToSite()

    Dim IE As SHDocVw.InternetExplorer
    
    Set IE = New SHDocVw.InternetExplorerMedium
    
    IE.Visible = True
    IE.Navigate "https://en.wikipedia.org/wiki/Main_Page"
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE
        
    Loop
    
    Debug.Print IE.LocationName, IE.LocationURL
    
    IE.Document.forms("searchform").elements("search").Value = "Document Object Model"
    IE.Document.forms("searchform").elements("go").Click
    
End Sub
