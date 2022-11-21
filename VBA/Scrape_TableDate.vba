Dim panel$
Dim URL$
Sub GetCHEM(histno As String)
    If Not Connected Then Exit Sub
    Set html = New HTMLDocument
    panel = "CHEM"
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findResd&resdtype=D" _
    & panel & "&resdtmonth=00&histno=" & histno
    'Create connection to get the html
    Set html = SaveHTML_Until_Id_Found(URL, "resdtable")
    Set resdtable = html.getElementById("resdtable")
    Set trs = resdtable.getElementsByTagName("tr")
    Dim k As Integer
    Dim start As Integer
    k = 4
    If trs.Length - 6 < 0 Then k = k + 6 - trs.Length
    start = trs.Length - 6
    If trs.Length - 6 < 0 Then start = 0
    
    For i = start To trs.Length - 2
        Set tds = trs(i).getElementsByTagName("td")
        For j = 0 To tds.Length - 1
            bot.Cells(k + 1, j + 1) = tds(j).innerText
            Application.StatusBar = tds(j).innerText
        Next
        k = k + 1
    Next
    
    DeleteYesterdayData (3)
    
    ' Clear all Object
    Set resdtable = Nothing
    Set trs = Nothing
    Set tds = Nothing
    Set html = Nothing
End Sub
Sub GetCBC(histno$)
    If Not Connected Then Exit Sub
    Set html = New HTMLDocument
    panel = "CBC"
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findResd&resdtype=D" & panel & "&resdtmonth=00&histno=" & histno
    Set html = SaveHTML_Until_Id_Found(URL, "resdtable")
    Set resdtable = html.getElementById("resdtable")
    Set trs = resdtable.getElementsByTagName("tr")
    Dim k As Integer
    Dim start As Integer

    k = 11
    If trs.Length - 6 < 0 Then k = k + 6 - trs.Length
    start = trs.Length - 6
    If trs.Length - 6 < 0 Then start = 0

    For i = start To trs.Length - 2
        Set tds = trs(i).getElementsByTagName("td")
        For j = 0 To tds.Length - 1
            bot.Cells(k + 1, j + 1) = tds(j).innerText
            Application.StatusBar = tds(j).innerText
        Next
        k = k + 1
    Next
    
    DeleteYesterdayData (11)
    ' Clear all Object
    Set resdtable = Nothing
    Set trs = Nothing
    Set tds = Nothing
    Set html = Nothing
End Sub
Sub GetBGAS(histno$)
    If Not Connected Then Exit Sub
    Set html = New HTMLDocument
    panel = "BGAS"
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findResd&resdtype=D" & panel & "&resdtmonth=00&histno=" & histno
    Set html = SaveHTML_Until_Id_Found(URL, "resdtable")
    Set resdtable = html.getElementById("resdtable")
    Set trs = resdtable.getElementsByTagName("tr")
    
    k = 19
    Set tds = trs(trs.Length - 2).getElementsByTagName("td")
    For j = 0 To tds.Length - 1
        bot.Cells(k + 1, j + 1) = tds(j).innerText
        Application.StatusBar = tds(j).innerText
    Next
    ' Clear all Object
    Set resdtable = Nothing
    Set trs = Nothing
    Set tds = Nothing
    Set html = Nothing
End Sub


Sub DeleteYesterdayData(position As Integer)
    Dim today As Integer
    today = Day(Date)
    For i = position + 1 To position + 5
        If today > Mid(bot.Cells(i, 1).Value, 7, 2) Then
           bot.Range("B" & i & ":AX" & i).Value = "-"
        End If
    Next
End Sub
Sub NotTodayAndClear()
    For i = 1 To 5
        If Date - Cells(i, 1).Value > 0 Or Cells(i, 1).Value = Nothing Then
            Range("B" & i & ":AX" & i).Value = "-"
        End If
    Next
End Sub
