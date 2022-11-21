Sub GetPtList()
    If Not Connected() Then Exit Sub
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&srnId=DRWEBAPP&seqNo=008"
    Set html = SaveHTML_Until_Id_Found(URL, "patlist")
    Set patlist = html.getElementById("patlist")
    Set trs = patlist.getElementsByTagName("tr")
    'write patient list into the worksheet
    For i = 1 To trs.Length - 1
        Set tds = trs(i).getElementsByTagName("td")
        For j = 0 To tds.Length - 2
            PtList.Cells(i + 1, j + 1) = tds(j).innerText
            Application.StatusBar = tds(j).innerText
        Next
    Next
    Dim max As Integer
    max = PatientCount()
    For counter = 1 To max
    Set curCell = PtList.Cells(counter, 2)
    curCell.Value = Replace(curCell.Value, "New ", "")
    Next counter
    Application.StatusBar = "Complete"
    With PtList.Range("A1")
        .Value = Date
        .NumberFormatLocal = "m/d;@"
    End With

    Set patlist = Nothing
    Set html = Nothing
    Set trs = Nothing
    Set tds = Nothing
End Sub

