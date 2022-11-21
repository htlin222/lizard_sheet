Sub PasteSearch()
On Error GoTo ErrorHandler
    searchlist.Range("a1").Select
    searchlist.PasteSpecial Format:="HTML", link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
ErrorHandler:                   ' 錯誤處理用的程式碼
     Application.StatusBar = "格式錯誤 " & Err.Number & "：" & Err.Description
End Sub
Sub OnePatientSearchlist(patientIndex As Integer)
    If Not Connected Then Exit Sub
    Dim Result As String
    histno = searchlist.Cells(patientIndex, 4).Value
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
    CopyDataFromBotToSearchList (patientIndex)
    If edge Is Nothing Then MsgBox ("Oops, there's something wrong")
    bot.Range("5:9,12:16,20:20,24:24").Value = "-" 'Clear Data
End Sub
Sub CopyDataFromBotToSearchList(patientIndex As String)
    ' copy data from bot and set font format
    ' set date format
    ' copy data of this patient from bot result
    With searchlist.Range("I" & patientIndex & ":R" & patientIndex)
        'From Data in bot to ptlist
        .Value = bot.Range("B2:K2").Value
    End With
    ' add format
    With searchlist.Range("I" & patientIndex & ":R" & patientIndex)
        .Font.Size = 10
        .Font.name = "Cascadia Code SemiBold"
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    'PtList.Columns("H:T").AutoFit
End Sub
Function PatientCountSearchList() As Integer
    PatientCountSearchList = searchlist.Cells(Rows.count, 2).End(xlUp).row - 1
End Function
Function SetEndOfDataInPtList() As String
 ' #TODO:U for BW, and V for IO
    SetEndOfData = "V"
End Function

