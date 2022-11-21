Sub GetWeight(histno As String)
    On Error GoTo ErrorHandler
    If Not Connected() Then Exit Sub
    Set html = New HTMLDocument
    Dim weight$
    Dim weigthDate$
    caseno = GetCaseno(histno)
    
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findVts&histno=" & histno & "&caseno=" & caseno & "&pbvtype=HWS"
    edge.Get URL
    x = Timer() + 5 '<== capture current time and plus 5 second
    Do Until edge.FindElementsByTag("td").count > 0 Or Timer() > x
        DoEvents
    Loop
    
    'Save data
    html.body.innerHTML = edge.PageSource
    'get weight
    weight = html.getElementsByTagName("td")(2).innerText
    'get weightDate
    weightDate = html.getElementsByTagName("td")(0).innerText
    'Formatting
    weightDate = Mid(weightDate, 5, 4)
    bot.Range("H24").Value = weight & "(" & weightDate & ")"
'    PtList.Range("N2").Value = ">" & weight & "(" & weightDate & ")"
    Set html = Nothing

ErrorHandler:                   ' 錯誤處理用的程式碼
     Application.StatusBar = "錯誤 " & Err.Number & "：" & Err.Description
End Sub
Sub RegexGetSumSheet(patientIndex As Integer)
' NO IN USE, will update at 07:00, too late !
    Dim html As HTMLDocument
    Dim panel As String
    Dim recentBW As String
    Dim text As String
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
    Dim RE_IO As Object
    Set RE_IO = CreateObject("vbscript.regexp")
    Dim histno As String
    histno = PtList.Cells(patientIndex, 2).Value
    
    With RE
        .Pattern = "(體重：)(\d+.\d)(\d+kg)"
        .Global = True
        .IgnoreCase = True
    End With
    With RE_IO
        .Pattern = "I\/O.+:(\d+) \/ (\d+) diff = (\W?\d+)"
        .Global = True
        .IgnoreCase = True
    End With

    Set IE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
    IE.navigate "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPhck&histno=" & histno
    Call CountDown("weight", 2)
    Set html = IE.document
    
    text = html.getElementsByTagName("tr")(0).innerText
    'text = "體重：93.00kg"

    Set allMatches = RE.Execute(text)
    If allMatches.count <> 0 Then
        Result = allMatches.Item(0).submatches.Item(1)
        PtList.Range("U" & patientIndex).Value = ">" & Result & "(" & Format(Date, "mm/dd") & ")"
    Else
        PtList.Range("V" & patientIndex).Value = "沒記辣"
    End If
    
    
    Set allMatches_IO = RE_IO.Execute(text)
    
    If allMatches_IO.count <> 0 Then
        result_input = allMatches_IO.Item(0).submatches.Item(0)
        result_output = allMatches_IO.Item(0).submatches.Item(1)
        result_diff = allMatches_IO.Item(0).submatches.Item(2)
        PtList.Range("V" & patientIndex).Value = "I/O: " & result_input & "/" & result_output & " (" & result_diff & ")"
    Else
        PtList.Range("V" & patientIndex).Value = "沒記辣"
    End If

    ' Clear all Object
    Set html = Nothing
    
End Sub

