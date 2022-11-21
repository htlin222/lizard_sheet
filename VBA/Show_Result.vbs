Sub Batch()
'載入數據按鈕
    PtList.Select
    If Not Connected() Then Exit Sub
    Dim max As Integer
    max = PatientCount()
    For counter = 2 To max + 1
        OnePatient (counter)
    Next counter
    Application.StatusBar = "The batch process is complete, remember to log out to release disk memory if the bot is not in use"
End Sub
Sub BatchSearchList()
    If Not Connected() Then Exit Sub
    Dim max As Integer
    max = PatientCountSearchList()
    For counter = 2 To max + 1
        OnePatientSearchlist (counter)
    Next counter
    Application.StatusBar = "The batch process is complete, remember to log out to release disk memory if the bot is not in use"
End Sub
Sub CopyResult(patienIndex As String) '未使用
    PtList.Range("H" & patienIndex & ":T" & patienIndex).Select
    Selection.Copy
End Sub
Sub ClearPtList()
    PtList.Select
    PtList.Range("A2:W50").ClearContents
End Sub
Sub ClearPtData()
    PtList.Select
    PtList.Range("F2:Z50").ClearContents
End Sub
Sub ClearSearchList()
    searchlist.Range("A2:R50").ClearContents
End Sub
Sub ClearSearchData()
    searchlist.Range("I2:R50").ClearContents
End Sub

