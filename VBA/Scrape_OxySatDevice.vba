Sub GetOxySatDevice(histno As String)
    If Not Connected() Then Exit Sub
    Set html = New HTMLDocument
    Dim OXtime$
    Dim Flow$
    Dim Sat$
    Dim Device
    caseno = GetCaseno(histno)
    On Error GoTo ErrorHandler
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findVts&histno=" & histno & "&caseno=" & caseno & "&pbvtype=OXY"
    edge.Get URL
    x = Timer() + 5 '<== capture current time and plus 5 second
    Do Until edge.FindElementsByTag("td").count > 0 Or Timer() > x
        DoEvents
    Loop
    'Save data
    html.body.innerHTML = edge.PageSource
    OXtime = html.getElementsByTagName("td")(0).innerText
    Sat = html.getElementsByTagName("td")(1).innerText
    Device = "RA"
    
    For i = 0 To 40 Step 4
        If Not html.getElementsByTagName("td")(3 + i).innerText = "  " Then
            Flow = html.getElementsByTagName("td")(2 + i).innerText
            Device = html.getElementsByTagName("td")(3 + i).innerText
            Exit For
        Else
        End If
    Next
    OXtime = Mid(OXtime, 11, 5)
    Sat = RTrim(Sat)
    Device = RTrim(Replace(Device, "非侵入性裝置", ""))
    Device = Replace(Device, "侵入性裝置", "MV:")
    Device = Replace(Device, "[", "")
    Device = Replace(Device, "]", "")
    Flow = Replace(Flow, "/min", "")
    bot.Range("E24").Value = Sat
    bot.Range("F24").Value = Device & " " & Flow
    Set html = Nothing

ErrorHandler:                   ' 錯誤處理用的程式碼
     Application.StatusBar = "錯誤 " & Err.Number & "：" & Err.Description
End Sub


