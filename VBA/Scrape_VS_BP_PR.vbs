Sub GetBPPR(histno As String)
'    histno As String
    If Not Connected() Then Exit Sub
    Set html = New HTMLDocument
    Dim vsTime$ '#2(1)
    Dim tempTime$ '#2(1)
    Dim Temp$ '#3(2)
    Dim SBP$ '4#(3)
    Dim DBP$ '#5(4)
    Dim PR$ '#6(5)
    Dim RR$ '#7(6)
    Dim isZero$

    caseno = GetCaseno(histno)
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findTpr&caseno=" & caseno
    edge.Get URL
    x = Timer() + 5 '<== capture current time and plus 5 second
    Do Until edge.FindElementsByTag("td").count > 0 Or Timer() > x
        DoEvents
    Loop
    'Save data
    html.body.innerHTML = edge.PageSource

    vsTime = html.getElementsByTagName("td")(1).innerText
    
    
    For i = 0 To 91 Step 7
        If Not html.getElementsByTagName("td")(2 + i).innerText = "0度C" Then
            tempTime = html.getElementsByTagName("td")(1 + i).innerText
            Temp = html.getElementsByTagName("td")(2 + i).innerText
            Exit For
        End If
    Next
    
    
    For i = 0 To 14 Step 7
        If Not html.getElementsByTagName("td")(3 + i).innerText = "0/mmhg" Then
            vsTime = html.getElementsByTagName("td")(1 + i).innerText
            SBP = html.getElementsByTagName("td")(3 + i).innerText
            DBP = html.getElementsByTagName("td")(4 + i).innerText
            PR = html.getElementsByTagName("td")(5 + i).innerText
            RR = html.getElementsByTagName("td")(6 + i).innerText
            Exit For
        End If
    Next
    
    vsTime = Left(vsTime, 2) & ":" & Right(vsTime, 2)
    tempTime = Left(tempTime, 2) & ":" & Right(tempTime, 2)
    SBP = Left(SBP, Len(SBP) - 4)
    DBP = Replace(DBP, "/mmhg", "")
    PR = Replace(PR, "/min", "bpm")
    Temp = Replace(Temp, "度", "'")
    bot.Range("A24").Value = "[" & vsTime & "]"
    bot.Range("B24").Value = SBP & DBP & vbCrLf & PR & " " & RR
    bot.Range("C24").Value = "[" & tempTime & "]"
    bot.Range("D24").Value = Temp
    Set html = Nothing
End Sub
