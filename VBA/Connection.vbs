Public Cookies As Collection
Public cookieCount As Integer
Public username$
Public password$
Public histno$
Public patientIndex$
Public caseno$
Public edge As Object
Public URL$
'Public loginStatus As Boolean
Sub CheckConnectionBTN()
    If Connected() Then MsgBox "你已連線"
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=46784901"
    With edge
        Application.StatusBar = "Loading VGHTPE Web9 login page..."
        .Get URL

        x = Timer() + 5 '<== capture current time and plus 5 second
    End With
    
End Sub
Sub LoginBTN()
    Call LoginToGetCookie
End Sub
Sub LoginToGetCookie()
    'ShowRibbon False
    'login and click a patient's file, then save cookie as collection
    Application.StatusBar = "Start Login Process"
    Dim html As HTMLDocument
    Set edge = CreateObject("Selenium.EdgeDriver")
    Application.StatusBar = "Create Edge Driver"
    Set Cookies = New Collection
    URL = "https://eip.vghtpe.gov.tw/login.php"
    With edge
        Application.StatusBar = "Loading VGHTPE Web9 login page..."
        .Get URL
        x = Timer() + 5 '<== capture current time and plus 5 second
    End With
End Sub
Sub LogoutBTN()
    start.Range("B3").ClearContents
    start.Range("C3").ClearContents
    Application.StatusBar = "Logged out and disconnected. See you soon. If you want to reload the data, please log in."
    If edge Is Nothing Then Exit Sub
    edge.Quit
    Set edge = Nothing
End Sub
Sub ShowNotLogin()
    Application.StatusBar = "尚未連線，請登入"
    MsgBox "尚未連線，請登入"
End Sub
Function Connected() As Boolean
        'use when edge is create but might not login yet
        If edge Is Nothing Then
            Call ShowNotLogin
            Connected = False
        ElseIf edge.Title = ":::臺北榮民總醫院應用系統入口[Signon Screen]" Then
            Call ShowNotLogin
            Connected = False
        ElseIf edge.Title = "Error.jsp" Then
            Call ShowNotLogin
            Connected = False
        Else
            Application.StatusBar = "已連線"
            Connected = True
        End If
End Function
Function LoggedIn() As Boolean
        'use when edge is create but might not login yet
        If edge Is Nothing Then
            Call ShowNotLogin
            Connected = False
        ElseIf edge.Title = ":::臺北榮民總醫院應用系統入口[Signon Screen]" Then
            Call ShowNotLogin
            Connected = False
        ElseIf edge.Title = "Error.jsp" Then
            Call ShowNotLogin
            Connected = False
        Else
            Application.StatusBar = "已連線"
            Connected = True
        End If
End Function

Sub TestConnection()
    Application.StatusBar = "Start Login Process"
    Dim html As HTMLDocument
    Set edge = CreateObject("Selenium.EdgeDriver")
    Application.StatusBar = "Create Edge Driver"
    URL = "https://stackoverflow.com/"
    With edge
        Application.StatusBar = "Edge Driver was set to Headless"
        .SetCapability "ms:edgeOptions", "{" & """args"":[""headless""]" & "}"
        Application.StatusBar = "Launching Edge Web driver in headless mode....it might take a few seconds, be patient"
        .start
        Application.StatusBar = "Loading VGHTPE Web9 login page..."
    End With
    Set html = SaveHTML_Until_Id_Found(URL, "custom-header--")
    MsgBox html.body.innerHTML
End Sub

Function SaveHTML_Until_Id_Found(URL, Id$) As Object
    Set html = New HTMLDocument
    Application.StatusBar = "Ready togo"
    edge.Get URL
    x = Timer() + 5 '<== capture current time and plus 5 second
    Do Until edge.FindElementsById(Id).count > 0 Or Timer() > x
        DoEvents
    Loop
    html.body.innerHTML = edge.PageSource
    Set SaveHTML_Until_Id_Found = html
End Function

Sub RebuildConnectionByCookie()
    Set edge = CreateObject("Selenium.EdgeDriver")
    URL = "https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient&srnId=DRWEBAPP&seqNo=008"
    edge.Get URL
    Dim cookie As cookie
    edge.Manage.DeleteAllCookies
    Dim count As Integer
    count = 1
    For Each cookie In Cookies
        edge.Mange.AddCookie cookie.name, cookie.Value
'        edge.Manage.AddCookie shCookie.Cells(count, 1), shCoookie.Cells(count, 2)
        count = count + 1
    Next
    edge.Get URL
End Sub
Function hasCookie(name, Cookies As Collection)
    hasCookie = False
    For Each cookie In Cookies
        If cookie.name = name Then hasCookie = True
    Next
End Function
