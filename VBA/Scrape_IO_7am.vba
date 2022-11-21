Sub GetIOAM(histno As String)
' NO IN USE, will update at 07:00, too late !
    Dim html As HTMLDocument
    Dim panel As String
    Dim recentBW As String
    Dim text As String
    Dim RE_IO As Object
    Set RE_IO = CreateObject("vbscript.regexp")
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPhck&histno=" & histno
    
    With RE_IO
        .Pattern = "I\/O.+:(\d+) \/ (\d+) diff = (\W?\d+)"
        .Global = True
        .IgnoreCase = True
    End With
    
    edge.Get URL
    x = Timer() + 5 '<== capture current time and plus 5 second
    Do Until edge.FindElementsByTag("td").count > 0 Or Timer() > x
        DoEvents
    Loop
    
    'Save data
    text = edge.PageSource
    
    Set allMatches_IO = RE_IO.Execute(text)
    
    If allMatches_IO.count <> 0 Then
        result_input = allMatches_IO.Item(0).submatches.Item(0)
        result_output = allMatches_IO.Item(0).submatches.Item(1)
        result_diff = allMatches_IO.Item(0).submatches.Item(2)
        bot.Range("I24").Value = "I/O: " & result_input & "/" & result_output & " (" & result_diff & ")"
    Else
        bot.Range("I24").Value = "沒記"
    End If

    ' Clear all Object
    Set html = Nothing
    
End Sub

