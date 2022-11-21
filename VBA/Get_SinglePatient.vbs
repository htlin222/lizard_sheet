Sub OnePatient(patientIndex As Integer)
    If Not Connected Then Exit Sub
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=46784901"
    edge.Get URL
    Dim Result As String
    histno = PtList.Cells(patientIndex, 2).Value
    If histno = "" Then Exit Sub
    Application.StatusBar = "Start process of " & PtList.Cells(patientIndex, 3).Value & ", please wait..."
    Call GetCHEM(histno)
    Call GetCBC(histno)
    Call GetBGAS(histno)
'    Call GetOnlyWeight(GetCaseNumber(patientIndex), patientIndex)
    Call GetWeight(histno)
    Call GetBPPR(histno)
    Call GetOxySatDevice(histno)
    Call GetIOAM(histno)
    Application.StatusBar = PtList.Cells(patientIndex, 3).Value & _
    "'s data is completed, remember to logout to release disk memory if the bot is not in use"
    CopyDataFromBotToPtList (patientIndex)
    If edge Is Nothing Then MsgBox ("Oops, there's something wrong")
    bot.Range("5:9,12:16,20:20,24:24").Value = "-" 'Clear Data
End Sub
Sub CopyDataFromBotToPtList(patientIndex As String)
    ' copy data from bot and set font format
    ' set date format
    ' copy data of this patient from bot result
    With PtList.Range("F" & patientIndex & ":O" & patientIndex)
        'From Data in bot to ptlist
        .Value = bot.Range("B2:K2").Value
    End With
    ' add format
    With PtList.Range("F" & patientIndex & ":O" & patientIndex)
        .Font.Size = 9
        .Font.name = "Cascadia Code SemiBold"
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    'PtList.Columns("H:T").AutoFit
End Sub
Function PatientCount() As Integer
    PatientCount = PtList.Cells(Rows.count, 2).End(xlUp).row - 1
End Function
Function SetEndOfDataInPtList() As String
 ' #TODO:U for BW, and V for IO
    SetEndOfData = "V"
End Function
