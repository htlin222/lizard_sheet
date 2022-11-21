Sub loadWardBtn()
    Dim wardID As String
    Dim specialty As String
    wardID = start.Range("C13").Value
    specialty = start.Range("C14").Value & "  "
    Call searchByWard(wardID, specialty)
    Application.StatusBar = "Complete"
End Sub
Sub searchByWard(wardID As String, specialty As String)
    If Not Connected() Then Exit Sub
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&srnId=DRWEBAPP&seqNo=008"
    Set html = New HTMLDocument
    Dim wd As Object
    Application.StatusBar = "Ready togo"
    edge.Get URL
    x = Timer() + 5 '<== capture current time and plus 5 second
    Do Until edge.FindElementsById("patlist").count > 0 Or Timer() > x
        DoEvents
    Loop
    
    With edge
        Application.StatusBar = "Send query"
        Set wd = edge.FindElementByName("wd").AsSelect
        wd.SelectByText (wardID)
        Application.StatusBar = "Click search"
        .FindElementById("btn00").Click
    End With
    
    Do Until edge.FindElementsById("bedlist").count > 0 Or Timer() > x
        DoEvents
    Loop
    html.body.innerHTML = edge.PageSource
    Set bedlist = html.getElementById("bedlist")
    If bedlist Is Nothing Then MsgBox ("nothing")
    Set trs = bedlist.getElementsByTagName("tr")
    
    Dim row As Integer
    row = 1
    For i = 1 To trs.Length - 1
        Set tds = trs(i).getElementsByTagName("td")
        ' if no 查 then skip
        If tds(0).innerText = "查" Then
            If tds(5).innerText = specialty Then
                For j = 0 To tds.Length - 1
                    Dim text As String
                    text = tds(j).innerText
                    text = clear_DRG(text)
                    ' if no 查 then skip
                    searchlist.Cells(row + 1, j + 1) = text
                   ' MsgBox tds(j).innerText
                Next
                Dim Combined As String
                Combined = searchlist.Cells(row + 1, 2).Value & vbNewLine & searchlist.Cells(row + 1, 3).Value
                searchlist.Cells(row + 1, 2).Value = Combined
                searchlist.Cells(row + 1, 19).Value = Combined
                row = row + 1
            End If
        End If
    Next
    
    Dim max As Integer
    max = PatientCount()
    For counter = 1 To max
    Set curCell = searchlist.Cells(counter, 2)
    curCell.Value = Replace(curCell.Value, "New ", "")
    Next counter

    Set patlist = Nothing
    Set html = Nothing
    Set trs = Nothing
    Set tds = Nothing
End Sub
Function clear_DRG(Str As String)

Dim RegEx As Object
Set RegEx = CreateObject("VBScript.RegExp")
Str = Replace(Str, Chr(10), "")
With RegEx
    .Pattern = "\[\D+\S+\s+\d+\S+\s+\d+\S+\s+\S+\s+\d+\S+\s+\S+\s+\S+\s+\S+"
    .Global = True      'If FALSE, Replaces only the first matching string'
End With
clear_DRG = RegEx.Replace(Str, "")
End Function

Sub selectAll()
    RowCount = searchlist.Cells(Rows.count, 2).End(xlUp).row
    searchlist.Range(Cells(1, 1), Cells(RowCount, 19)).Select
End Sub

