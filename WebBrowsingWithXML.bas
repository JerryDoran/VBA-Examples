Attribute VB_Name = "WebBrowsingWithXML"
Option Explicit
Const WOVidURL As String = "http://wiseowl.co.uk/videos/"

Sub GetVideoPage()

    'Create two autoinstancing variables
    Dim XMLReq As New MSXML2.XMLHTTP60
    Dim HTMLDoc As New MSHTML.HTMLDocument
    Dim VidCats As MSHTML.IHTMLElementCollection
    Dim VidCat As MSHTML.IHTMLElement
    Dim VidCatList As MSHTML.IHTMLElement
    Dim VidCatID As Integer
    Dim NextHref As String
    Dim NextURL As String
    
    XMLReq.Open "GET", WOVidURL, False
    
    'Send the request
    XMLReq.send
    
    If XMLReq.Status <> 200 Then
        MsgBox "Problem" & vbNewLine & XMLReq.Status & " - " & XMLReq.statusText
        Exit Sub
    End If
    
    'Creates the new HTML document
    HTMLDoc.body.innerHTML = XMLReq.responseText
    Set XMLReq = Nothing
    
    Set VidCatList = HTMLDoc.getElementsByClassName("woMenuList")(0)    'Get first element in the collection only
    Set VidCats = VidCatList.getElementsByTagName("a")
    
    Debug.Print VidCats.Length      'How many items are in the collection
    
    For VidCatID = 1 To VidCats.Length - 1
        Set VidCat = VidCats(VidCatID)
        
        NextHref = VidCat.getAttribute("href")
        NextURL = WOVidURL & Mid(NextHref, InStr(NextHref, ":") + 9)
        
        ListVideosOnPage VidCat.innerText, NextURL
        
    Next VidCatID
        
End Sub

Sub ListVideosOnPage(VidCatName As String, VidCatURL As String)

    Dim XMLReq As New MSXML2.XMLHTTP60
    Dim HTMLDoc As New MSHTML.HTMLDocument
    Dim VidRow As MSHTML.IHTMLElement
    Dim VidRows As MSHTML.IHTMLElementCollection
    Dim VidLink As MSHTML.IHTMLElement
    Dim VidPages As MSHTML.IHTMLElementCollection
    Dim i As Integer
    Dim NextHref As String
    Dim NextURL As String
    Dim VidHref As String
    Dim VidURL As String
    
    XMLReq.Open "GET", VidCatURL, False
    
    'Send the request
    XMLReq.send
    
    If XMLReq.Status <> 200 Then
        MsgBox "Problem" & vbNewLine & XMLReq.Status & " - " & XMLReq.statusText
        Exit Sub
    End If
    
    'Creates the new HTML document
    HTMLDoc.body.innerHTML = XMLReq.responseText
    Set XMLReq = Nothing
    
    Worksheets.Add
    Range("A1").Value = VidCatName
    Range("B1").Value = "Video URL"
    Range("A1:B1").Interior.Color = rgbCornflowerBlue
    Range("A1:B1").Font.Color = rgbWhite
    Range("A1:B1").Font.Bold = True
    Range("A2").Select
    
    Set VidPages = HTMLDoc.getElementsByClassName("woPagingItem")
    
    For i = 0 To VidPages.Length - IIf(VidPages.Length > 0, 1, 0)
    
        If i > 0 Then
            NextHref = VidPages(i).getAttribute("href")
            NextURL = WOVidURL & Mid(NextHref, InStr(NextHref, ":") + 9)
            'Debug.Print NextURL
            XMLReq.Open "GET", NextURL, False
    
            'Send the request
            XMLReq.send
            
            If XMLReq.Status <> 200 Then
                MsgBox "Problem" & vbNewLine & XMLReq.Status & " - " & XMLReq.statusText
                Exit Sub
            End If
            
            'Creates the new HTML document
            HTMLDoc.body.innerHTML = XMLReq.responseText
            Set XMLReq = Nothing
            
        End If
            'Can only use 'getElementsByClassName' method on an HTMLDocument object!
            Set VidRows = HTMLDoc.getElementsByClassName("woVideoListRow")
        
            For Each VidRow In VidRows
                Set VidLink = VidRow.getElementsByTagName("a")(0)
                VidHref = VidLink.getAttribute("href")
                VidURL = WOVidURL & Mid(VidHref, InStr(VidHref, ":") + 9)
                ActiveCell.Value = VidLink.innerText
                ActiveCell.Offset(0, 1).Value = VidURL
                
                ActiveCell.Offset(0, 1).Hyperlinks.Add ActiveCell.Offset(0, 1), VidURL
                
                ActiveCell.Offset(1, 0).Select
            Next VidRow
    Next i
    
    Range("A1").Select
    ActiveCell.CurrentRegion.EntireColumn.AutoFit
    
End Sub






































