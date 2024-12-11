Option Explicit

Dim IE, URL
Dim objFSO, objFile
Dim xmlDoc, rootElement, urlElement, locElement
Dim link, links

' Change this to the URL you want to scrape
URL = "https://lostinthecyberabyss.com"

' Array of strings to filter file names
Dim filterArray
filterArray = Array("admin", "test", "debug", "wordpress.org")

' Function to check if a file name contains any filter string
Function IsFiltered(fileName)
    Dim i
    For i = LBound(filterArray) To UBound(filterArray)
        If InStr(LCase(fileName), LCase(filterArray(i))) > 0 Then
            IsFiltered = True
            Exit Function
        End If
    Next
    IsFiltered = False
End Function

' Create Internet Explorer instance
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = False
IE.Navigate URL

' Wait for the page to load
Do While IE.Busy Or IE.readyState <> 4
    WScript.Sleep 100
Loop

' Create XML document with header
Set xmlDoc = CreateObject("MSXML2.DOMDocument")
xmlDoc.appendChild xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

' Create the root <urlset> element with namespaces
Set rootElement = xmlDoc.createElement("urlset")
rootElement.setAttribute "xmlns", "http://www.sitemaps.org/schemas/sitemap/0.9"
xmlDoc.appendChild rootElement

' Collect links and add them as <url> and <loc> elements
Set links = IE.document.getElementsByTagName("a")
For Each link in links
    If link.href <> "" And Not IsFiltered(link.href) Then
        Set urlElement = xmlDoc.createElement("url")
        Set locElement = xmlDoc.createElement("loc")
        locElement.Text = link.href
        urlElement.appendChild locElement
        rootElement.appendChild urlElement
    End If
Next

' Save the XML Sitemap
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("sitemap.xml", True)
objFile.WriteLine xmlDoc.xml
objFile.Close

' Clean up
IE.Quit
Set IE = Nothing
Set objFSO = Nothing
Set objFile = Nothing
Set xmlDoc = Nothing
