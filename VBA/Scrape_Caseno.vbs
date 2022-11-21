Function GetCaseno(histno As String) As String
    Set html = New HTMLDocument
    Dim Result As String
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findNIHSS&histno=" & histno
    'Create connection to get the html
    edge.Get URL
    'still need time out
    'Set html = SaveHTML_Until_Id_Found(URL, "")  WHAT is ID?????
    x = Timer() + 5 '<== capture current time and plus 5 second
    Do Until edge.FindElementsByTag("td").count > 0 Or Timer() > x
        DoEvents
    Loop
    html.body.innerHTML = edge.PageSource
    'start processing data
    GetCaseno = html.getElementsByTagName("td")(1).innerText
    ' Clear all Object
    Set html = Nothing
End Function
孕孕…………