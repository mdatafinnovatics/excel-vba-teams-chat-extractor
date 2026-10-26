Attribute VB_Name = "Microsoft_Signin"

Option Explicit

Public Const CLIENT_ID As String = "YOUR_CLIENT_ID"
Public Const TENANT_ID As String = "YOUR_TENANT_ID"

Public Const REDIRECT_URI As String = _
"https://login.microsoftonline.com/common/oauth2/nativeclient"

Public GraphAccessToken As String
Public GraphRefreshToken As String
Public PKCE_Verifier As String
Public PKCE_Challenge As String
Public OAuthState As String
Public OAuthAuthorizationCode As String


Public Function IsMac() As Boolean
    IsMac = (InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0)
End Function

Public Function HttpPost(ByVal Url As String, ByVal PostData As String) As String

#If Mac Then

    Dim ScriptText As String

    ScriptText = _
        "do shell script " & Chr(34) & _
        "curl -s -X POST " & _
        "'" & Url & "'" & _
        " -H 'Content-Type: application/x-www-form-urlencoded'" & _
        " --data " & _
        "'" & Replace(PostData, "'", "'\''") & "'" & _
        Chr(34)

    HttpPost = MacScript(ScriptText)

#Else

    Dim http As Object

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "POST", Url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send PostData

    HttpPost = http.ResponseText

#End If

End Function

Public Function HttpGet(ByVal Url As String, _
                        ByVal AccessToken As String) As String

    Dim ResponseText As String

#If Mac Then

    Dim ScriptText As String

    ScriptText = _
        "do shell script " & Chr(34) & _
        "curl -s " & _
        "'" & Url & "'" & _
        " -H 'Authorization: Bearer " & Replace(AccessToken, "'", "'\''") & "'" & _
        " -H 'Accept: application/json'" & _
        " -H 'Cache-Control: no-cache'" & _
        " -H 'Pragma: no-cache'" & _
        Chr(34)

    ResponseText = MacScript(ScriptText)

#Else

    Dim http As Object

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", Url, False

    http.setRequestHeader "Authorization", _
                          "Bearer " & AccessToken

    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Cache-Control", "no-cache"
    http.setRequestHeader "Pragma", "no-cache"

    http.Send

    ResponseText = http.ResponseText

#End If

    If InStr(1, ResponseText, "InvalidAuthenticationToken", vbTextCompare) > 0 _
       Or InStr(1, ResponseText, "token is expired", vbTextCompare) > 0 Then

        If RefreshAccessToken() Then

#If Mac Then

            ScriptText = _
                "do shell script " & Chr(34) & _
                "curl -s " & _
                "'" & Url & "'" & _
                " -H 'Authorization: Bearer " & Replace(GraphAccessToken, "'", "'\''") & "'" & _
                " -H 'Accept: application/json'" & _
                " -H 'Cache-Control: no-cache'" & _
                " -H 'Pragma: no-cache'" & _
                Chr(34)

            ResponseText = MacScript(ScriptText)

#Else

            Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

            http.Open "GET", Url, False

            http.setRequestHeader "Authorization", _
                                  "Bearer " & GraphAccessToken

            http.setRequestHeader "Accept", "application/json"
            http.setRequestHeader "Cache-Control", "no-cache"
            http.setRequestHeader "Pragma", "no-cache"

            http.Send

            ResponseText = http.ResponseText

#End If

        End If

    End If

    HttpGet = ResponseText

End Function

Sub TestOS()
    MsgBox Application.OperatingSystem & vbCrLf & _
           "Is Mac: " & IsMac()
End Sub


Public Sub TestMicrosoftLogin()

    Dim Url As String
    Dim TokenUrl As String
    Dim PostData As String
    Dim TokenPostData As String
    Dim Response As String
    Dim TokenResponse As String

    Dim UserCode As String
    Dim DeviceCode As String
    Dim VerificationUrl As String
    Dim ErrorCode As String

    '-----------------------------------
    ' Step 1 - Request Device Code
    '-----------------------------------

    Url = "https://login.microsoftonline.com/" & _
          TENANT_ID & _
          "/oauth2/v2.0/devicecode"

    PostData = "client_id=" & CLIENT_ID & _
               "&scope=Team.ReadBasic.All%20" & _
               "Channel.ReadBasic.All%20" & _
               "ChannelMessage.Read.All%20" & _
               "offline_access"

    Response = HttpPost(Url, PostData)

    UserCode = GetJsonValue(Response, "user_code")
    DeviceCode = GetJsonValue(Response, "device_code")
    VerificationUrl = GetJsonValue(Response, "verification_uri")

    If DeviceCode = "" Then
        MsgBox "Unable to obtain Microsoft Device Code." & _
               vbCrLf & vbCrLf & Response, vbCritical
        Exit Sub
    End If

    '-----------------------------------
    ' Step 2 - Ask user to login
    '-----------------------------------

    MsgBox "Please open:" & vbCrLf & vbCrLf & _
           VerificationUrl & vbCrLf & vbCrLf & _
           "Enter this code:" & vbCrLf & vbCrLf & _
           UserCode & vbCrLf & vbCrLf & _
           "Complete the Microsoft login and then click OK.", _
           vbInformation, _
           "Microsoft Sign In"

    '-----------------------------------
    ' Step 3 - Token Endpoint
    '-----------------------------------

    TokenUrl = "https://login.microsoftonline.com/" & _
               TENANT_ID & _
               "/oauth2/v2.0/token"

    TokenPostData = _
        "grant_type=urn:ietf:params:oauth:grant-type:device_code" & _
        "&client_id=" & CLIENT_ID & _
        "&device_code=" & DeviceCode

    '-----------------------------------
    ' Step 4 - Poll Microsoft
    '-----------------------------------

    Do

        TokenResponse = HttpPost(TokenUrl, TokenPostData)

        GraphAccessToken = GetJsonValue(TokenResponse, "access_token")

        If GraphAccessToken <> "" Then
            Exit Do
        End If

        ErrorCode = GetJsonValue(TokenResponse, "error")

        If ErrorCode <> "authorization_pending" Then

            MsgBox "Microsoft login failed:" & vbCrLf & vbCrLf & _
                   TokenResponse, vbCritical

            Exit Sub

        End If

        Application.Wait Now + TimeValue("00:00:05")

        DoEvents

    Loop

    MsgBox "Microsoft login successful!" & vbCrLf & vbCrLf & _
           "Access Token received successfully.", _
           vbInformation

End Sub

Public Function GetJsonValue(ByVal JsonText As String, _
                             ByVal KeyName As String) As String

    Dim StartPos As Long
    Dim EndPos As Long
    Dim SearchText As String

    SearchText = """" & KeyName & """:"""

    StartPos = InStr(1, JsonText, SearchText, vbTextCompare)

    If StartPos = 0 Then Exit Function

    StartPos = StartPos + Len(SearchText)
    EndPos = InStr(StartPos, JsonText, """")

    GetJsonValue = Mid(JsonText, StartPos, EndPos - StartPos)

End Function


Public Function GetAuthorizationUrl() As String

    Dim Scopes As String

    PKCE_Verifier = GenerateRandomString(64)
'    ThisWorkbook.Worksheets("Lists").Range("Z4").Value = PKCE_Verifier
    OAuthState = GenerateRandomString(32)
    ThisWorkbook.Worksheets("Lists").Range("Z1").Value = OAuthState

#If Mac Then

    PKCE_Challenge = MacScript( _
        "do shell script " & _
        Chr(34) & _
        "printf %s " & _
        "'" & PKCE_Verifier & "'" & _
        " | openssl dgst -sha256 -binary" & _
        " | openssl base64" & _
        " | tr -d '\n'" & _
        " | tr '+/' '-_'" & _
        " | tr -d '='" & _
        Chr(34))

#Else

    PKCE_Challenge = SHA256_Base64Url_Windows(PKCE_Verifier)

#End If

    Scopes = "Team.ReadBasic.All " & _
             "Channel.ReadBasic.All " & _
             "ChannelMessage.Read.All " & _
             "offline_access"

    GetAuthorizationUrl = _
        "https://login.microsoftonline.com/" & TENANT_ID & _
        "/oauth2/v2.0/authorize" & _
        "?client_id=" & CLIENT_ID & _
        "&response_type=code" & _
        "&redirect_uri=http%3A%2F%2Flocalhost" & _
        "&response_mode=query" & _
        "&scope=" & Replace(Scopes, " ", "%20") & _
        "&state=" & OAuthState & _
        "&code_challenge=" & PKCE_Challenge & _
        "&code_challenge_method=S256"

End Function


Public Sub TestBrowserLogin()

    Dim LoginUrl As String

    LoginUrl = GetAuthorizationUrl()

    ThisWorkbook.FollowHyperlink LoginUrl

End Sub



Public Function GenerateRandomString(ByVal Length As Long) As String

    Dim chars As String
    Dim i As Long
    Dim Result As String

    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"

    Randomize

    For i = 1 To Length
        Result = Result & Mid$(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next i

    GenerateRandomString = Result

End Function

Public Sub TestPKCEValues()

    PKCE_Verifier = GenerateRandomString(64)
    OAuthState = GenerateRandomString(32)

#If Mac Then

    MsgBox "Mac SHA-256 helper will be added separately."

#Else

    PKCE_Challenge = SHA256_Base64Url_Windows(PKCE_Verifier)

    If PKCE_Challenge <> "" Then
        MsgBox "PKCE S256 challenge generated successfully."
    Else
        MsgBox "PKCE challenge generation failed."
    End If

#End If

End Sub

Public Function Base64UrlEncode(ByRef bytes() As Byte) As String

    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim s As String

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")

    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = bytes

    s = xmlNode.Text

    s = Replace(s, "+", "-")
    s = Replace(s, "/", "_")
    s = Replace(s, "=", "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    Base64UrlEncode = s

End Function


Public Function SHA256_Base64Url_Windows(ByVal Text As String) As String

#If Mac Then

    SHA256_Base64Url_Windows = ""

#Else

    Dim Shell As Object
    Dim Exec As Object
    Dim Cmd As String
    Dim Result As String
    Dim SafeText As String

    SafeText = Replace(Text, "'", "''")

    Cmd = "powershell -NoProfile -Command " & _
          """$b=[System.Text.Encoding]::ASCII.GetBytes('" & SafeText & "');" & _
          "$h=[System.Security.Cryptography.SHA256]::Create().ComputeHash($b);" & _
          "[Convert]::ToBase64String($h)"""

    Set Shell = CreateObject("WScript.Shell")
    Set Exec = Shell.Exec(Cmd)

    Result = Trim$(Exec.StdOut.ReadAll)

    Result = Replace(Result, "+", "-")
    Result = Replace(Result, "/", "_")
    Result = Replace(Result, "=", "")
    Result = Replace(Result, vbCr, "")
    Result = Replace(Result, vbLf, "")

    SHA256_Base64Url_Windows = Result

#End If

End Function


Public Sub CaptureAuthorizationCode()

    Dim FullUrl As String
    Dim CodeStart As Long
    Dim CodeEnd As Long
    Dim StateStart As Long
    Dim ReturnedState As String
    Dim AuthCode As String

    FullUrl = ThisWorkbook.Worksheets("Lists").Range("Z2").Value
    
    If FullUrl = "" Then
        MsgBox "Please paste the complete localhost URL into Lists!Z2 first.", vbExclamation
        Exit Sub
    End If

    CodeStart = InStr(1, FullUrl, "code=", vbTextCompare)

    If CodeStart = 0 Then
        MsgBox "Authorization code was not found.", vbCritical
        Exit Sub
    End If

    CodeStart = CodeStart + 5
    CodeEnd = InStr(CodeStart, FullUrl, "&")

    If CodeEnd = 0 Then
        AuthCode = Mid$(FullUrl, CodeStart)
    Else
        AuthCode = Mid$(FullUrl, CodeStart, CodeEnd - CodeStart)
    End If

    StateStart = InStr(1, FullUrl, "state=", vbTextCompare)

    If StateStart > 0 Then
        ReturnedState = Mid$(FullUrl, StateStart + 6)

        If InStr(ReturnedState, "&") > 0 Then
            ReturnedState = Left$(ReturnedState, InStr(ReturnedState, "&") - 1)
        End If
    End If

    OAuthState = ThisWorkbook.Worksheets("Lists").Range("Z1").Value
    
    If ReturnedState <> OAuthState Then
        MsgBox "Security validation failed. State does not match.", vbCritical
        Exit Sub
    End If

    OAuthAuthorizationCode = AuthCode
    ThisWorkbook.Worksheets("Lists").Range("Z3").Value = OAuthAuthorizationCode

    MsgBox "Authorization code captured successfully!", vbInformation

End Sub


Public Function UrlEncode(ByVal Text As String) As String

    Dim i As Long
    Dim Ch As String
    Dim Result As String

    For i = 1 To Len(Text)
        Ch = Mid$(Text, i, 1)

        Select Case AscW(Ch)
            Case 48 To 57, 65 To 90, 97 To 122
                Result = Result & Ch
            Case 45, 46, 95, 126
                Result = Result & Ch
            Case Else
                Result = Result & "%" & Right$("0" & hex(AscW(Ch)), 2)
        End Select
    Next i

    UrlEncode = Result

End Function


Public Sub ExchangeCodeForToken()

    Dim TokenUrl As String
    Dim PostData As String
    Dim Response As String
    Dim ErrorCode As String

    OAuthAuthorizationCode = ThisWorkbook.Worksheets("Lists").Range("Z3").Value
'    PKCE_Verifier = ThisWorkbook.Worksheets("Lists").Range("Z4").Value
    
    If OAuthAuthorizationCode = "" Then
        MsgBox "Authorization code is missing. Please run the login process again.", vbExclamation
        Exit Sub
    End If

    If PKCE_Verifier = "" Then
        MsgBox "PKCE verifier is missing. Please run the login process again.", vbExclamation
        Exit Sub
    End If

    TokenUrl = "https://login.microsoftonline.com/" & _
               TENANT_ID & _
               "/oauth2/v2.0/token"

    PostData = _
        "client_id=" & UrlEncode(CLIENT_ID) & _
        "&grant_type=authorization_code" & _
        "&code=" & UrlEncode(OAuthAuthorizationCode) & _
        "&redirect_uri=" & UrlEncode("http://localhost") & _
        "&code_verifier=" & UrlEncode(PKCE_Verifier) & _
        "&scope=" & UrlEncode( _
            "Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All offline_access")

    Response = HttpPost(TokenUrl, PostData)

    GraphAccessToken = GetJsonValue(Response, "access_token")
    GraphRefreshToken = GetJsonValue(Response, "refresh_token")

    ThisWorkbook.Worksheets("Lists").Range("Z6").Value = GraphAccessToken
    ThisWorkbook.Worksheets("Lists").Range("Z10").Value = GraphRefreshToken
    
    If GraphAccessToken <> "" Then

        MsgBox "Microsoft Graph access token received successfully!", vbInformation

    Else

        ErrorCode = GetJsonValue(Response, "error")

        MsgBox "Token request failed." & vbCrLf & vbCrLf & _
               Response, vbCritical

    End If

End Sub



Public Sub TestGetTeams()

    Dim Url As String
    Dim Response As String

    GraphAccessToken = ThisWorkbook.Worksheets("Lists").Range("Z6").Value

    If GraphAccessToken = "" Then
        MsgBox "Access token is missing. Please login again.", vbExclamation
        Exit Sub
    End If

    Url = "https://graph.microsoft.com/v1.0/me/joinedTeams"

    Response = HttpGet(Url, GraphAccessToken)

    'Temporarily store the raw response for testing
    ThisWorkbook.Worksheets("Lists").Range("Z5").Value = Response

    MsgBox "Teams response received." & vbCrLf & _
           "Please check Lists!Z5.", vbInformation

End Sub


Public Sub TestGetChannels()

    Dim TeamID As String
    Dim Url As String
    Dim Response As String

    'Use the Team ID returned from joinedTeams
    TeamID = ThisWorkbook.Worksheets("Lists").Range("B2").Value

    If TeamID = "" Then
        MsgBox "Team ID is missing in Lists!B2.", vbExclamation
        Exit Sub
    End If

    GraphAccessToken = ThisWorkbook.Worksheets("Lists").Range("Z6").Value

    If GraphAccessToken = "" Then
        MsgBox "Access token is missing. Please login again.", vbExclamation
        Exit Sub
    End If

    Url = "https://graph.microsoft.com/v1.0/teams/" & TeamID & "/channels"

    Response = HttpGet(Url, GraphAccessToken)

    ThisWorkbook.Worksheets("Lists").Range("Z7").Value = Response

    MsgBox "Channel response received. Please check Lists!Z7.", vbInformation

End Sub


Public Sub PopulateTeams()

    Dim wsLists As Worksheet
    Dim wsMessages As Worksheet

    Dim Url As String
    Dim Response As String

    Dim Json As Object
    Dim TeamsCollection As Object
    Dim Team As Variant

    Dim RowNum As Long
    Dim LastRow As Long

    Set wsLists = ThisWorkbook.Worksheets("Lists")
    Set wsMessages = ThisWorkbook.Worksheets("Messages")

    'Get current access token
    GraphAccessToken = wsLists.Range("Z6").Value

    If GraphAccessToken = "" Then

        'Try to refresh automatically
        If Not RefreshAccessToken() Then
            MsgBox "Microsoft login is required.", vbExclamation
            Exit Sub
        End If

    End If

    'Get fresh Teams directly from Microsoft Graph
    Url = "https://graph.microsoft.com/v1.0/me/joinedTeams"

    Response = HttpGet(Url, GraphAccessToken)

    'Optional debug copy
    wsLists.Range("Z5").Value = Response

    Set Json = JsonConverter.ParseJson(Response)

    If Json Is Nothing Then
        MsgBox "Unable to read Teams response.", vbCritical
        Exit Sub
    End If

    If Not Json.Exists("value") Then
        MsgBox "Microsoft Graph returned an error:" & vbCrLf & vbCrLf & Response, vbCritical
        Exit Sub
    End If

    Set TeamsCollection = Json.Item("value")

    'Clear previous Teams
    wsLists.Range("A2:B1000").ClearContents

    RowNum = 2

    For Each Team In TeamsCollection

        wsLists.Cells(RowNum, "A").Value = Team.Item("displayName")
        wsLists.Cells(RowNum, "B").Value = Team.Item("id")

        RowNum = RowNum + 1

    Next Team

    LastRow = RowNum - 1

    If LastRow < 2 Then
        MsgBox "No Teams were found for this user.", vbExclamation
        Exit Sub
    End If

    'Create Team dropdown
    With wsMessages.Range("P3").Validation

        .Delete

        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=Lists!$A$2:$A$" & LastRow

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

End Sub

Public Sub PopulateChannels()

    Dim wsLists As Worksheet
    Dim wsMessages As Worksheet

    Dim SelectedTeam As String
    Dim TeamID As String

    Dim LastTeamRow As Long
    Dim i As Long

    Dim Url As String
    Dim Response As String

    Dim Json As Object
    Dim ChannelsCollection As Object
    Dim Channel As Variant

    Dim RowNum As Long
    Dim LastChannelRow As Long

    Set wsLists = ThisWorkbook.Worksheets("Lists")
    Set wsMessages = ThisWorkbook.Worksheets("Messages")

    SelectedTeam = Trim$(wsMessages.Range("P3").Value)

    If SelectedTeam = "" Then
        MsgBox "Please select a Team first.", vbExclamation
        Exit Sub
    End If

    'Find Team ID
    LastTeamRow = wsLists.Cells(wsLists.rows.Count, "A").End(xlUp).Row

    For i = 2 To LastTeamRow

        If wsLists.Cells(i, "A").Value = SelectedTeam Then

            TeamID = wsLists.Cells(i, "B").Value
            Exit For

        End If

    Next i

    If TeamID = "" Then
        MsgBox "Team ID could not be found.", vbCritical
        Exit Sub
    End If

    'Get access token
    GraphAccessToken = wsLists.Range("Z6").Value

    If GraphAccessToken = "" Then
        MsgBox "Access token is missing. Please login again.", vbExclamation
        Exit Sub
    End If

    'Call Graph
    Url = "https://graph.microsoft.com/v1.0/teams/" & TeamID & "/channels"

    Response = HttpGet(Url, GraphAccessToken)

    'Optional raw response for testing
    wsLists.Range("Z7").Value = Response

    Set Json = JsonConverter.ParseJson(Response)
    
    If Json Is Nothing Then
        MsgBox "Could not parse the channel response.", vbCritical
        Exit Sub
    End If
    
    If Not Json.Exists("value") Then
        MsgBox "Microsoft Graph returned an error:" & vbCrLf & vbCrLf & Response, vbCritical
        Exit Sub
    End If
    
    Set ChannelsCollection = Json.Item("value")


    'Clear old Channel list
    wsLists.Range("D2:E1000").ClearContents

    RowNum = 2

    For Each Channel In ChannelsCollection

        wsLists.Cells(RowNum, "D").Value = Channel.Item("displayName")
        wsLists.Cells(RowNum, "E").Value = Channel.Item("id")

        RowNum = RowNum + 1

    Next Channel

    LastChannelRow = RowNum - 1

    If LastChannelRow < 2 Then
        MsgBox "No channels were found for the selected Team.", vbExclamation
        Exit Sub
    End If

    'Create dropdown
    With wsMessages.Range("P4").Validation

        .Delete

        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=Lists!$D$2:$D$" & LastChannelRow

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

    wsMessages.Range("P4").ClearContents

'    MsgBox "Channel dropdown created successfully.", vbInformation

End Sub


Public Function GraphDateTime(ByVal ISODate As String) As Date

    If Len(ISODate) < 19 Then Exit Function

    GraphDateTime = DateSerial( _
        CInt(Mid$(ISODate, 1, 4)), _
        CInt(Mid$(ISODate, 6, 2)), _
        CInt(Mid$(ISODate, 9, 2))) + _
        TimeSerial( _
        CInt(Mid$(ISODate, 12, 2)), _
        CInt(Mid$(ISODate, 15, 2)), _
        CInt(Mid$(ISODate, 18, 2)))

End Function


Public Sub ImportChannelMessages()

    Dim wsL As Worksheet
    Dim wsM As Worksheet

    Dim SelectedTeam As String
    Dim SelectedChannel As String
    Dim TeamID As String
    Dim ChannelID As String

    Dim FromDate As Date
    Dim ToDate As Date

    Dim i As Long
    Dim LastRow As Long
    Dim OutputRow As Long

    Dim Url As String
    Dim Response As String
    Dim NextLink As String

    Dim Json As Object
    Dim MsgCollection As Object
    Dim Msg As Variant

    Dim Replies As Object
    Dim Reply As Variant

    Dim MsgDateTime As Date
    Dim Sender As String
    Dim MessageText As String
    Dim MessageID As String

    Dim FromObj As Object
    Dim UserObj As Object
    Dim BodyObj As Object

    Dim MsgRng As Range
    
    Set wsL = ThisWorkbook.Worksheets("Lists")
    Set wsM = ThisWorkbook.Worksheets("Messages")

    SelectedTeam = Trim$(wsM.Range("P3").Value)
    SelectedChannel = Trim$(wsM.Range("P4").Value)

    If SelectedTeam = "" Then
        MsgBox "Please select a Team.", vbExclamation
        Exit Sub
    End If

    If SelectedChannel = "" Then
        MsgBox "Please select a Channel.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(wsM.Range("P5").Value) Then
        MsgBox "Please enter a valid From Date.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(wsM.Range("P6").Value) Then
        MsgBox "Please enter a valid To Date.", vbExclamation
        Exit Sub
    End If

    FromDate = CDate(wsM.Range("P5").Value)
    ToDate = CDate(wsM.Range("P6").Value)

    If FromDate > ToDate Then
        MsgBox "From Date cannot be greater than To Date.", vbExclamation
        Exit Sub
    End If

    'Find Team ID
    LastRow = wsL.Cells(wsL.rows.Count, "A").End(xlUp).Row

    For i = 2 To LastRow
        If wsL.Cells(i, "A").Value = SelectedTeam Then
            TeamID = wsL.Cells(i, "B").Value
            Exit For
        End If
    Next i

    If TeamID = "" Then
        MsgBox "Team ID not found.", vbCritical
        Exit Sub
    End If

    'Find Channel ID
    LastRow = wsL.Cells(wsL.rows.Count, "D").End(xlUp).Row

    For i = 2 To LastRow
        If wsL.Cells(i, "D").Value = SelectedChannel Then
            ChannelID = wsL.Cells(i, "E").Value
            Exit For
        End If
    Next i

    If ChannelID = "" Then
        MsgBox "Channel ID not found.", vbCritical
        Exit Sub
    End If

    GraphAccessToken = wsL.Range("Z6").Value

    If GraphAccessToken = "" Then
        MsgBox "Access token is missing. Please login again.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    wsM.Range("A2:J" & wsM.rows.Count).ClearContents
    OutputRow = 2

    Url = "https://graph.microsoft.com/v1.0/teams/" & _
          TeamID & "/channels/" & _
          ChannelID & _
          "/messages?$top=50&$expand=replies"

    Do While Url <> ""

        Response = HttpGet(Url, GraphAccessToken)

        Set Json = JsonConverter.ParseJson(Response)
        Set MsgCollection = Json.Item("value")

        For Each Msg In MsgCollection

            '-----------------------------
            'ROOT POST
            '-----------------------------
            MsgDateTime = 0
            Sender = ""
            MessageText = ""
            MessageID = ""

            If Msg.Exists("createdDateTime") Then
                MsgDateTime = GraphDateTime(CStr(Msg.Item("createdDateTime")))
            End If

            If Msg.Exists("id") Then
                MessageID = CStr(Msg.Item("id"))
            End If

            If MsgDateTime >= FromDate And _
               MsgDateTime < ToDate + 1 Then

                Set FromObj = Nothing
                Set UserObj = Nothing
                Set BodyObj = Nothing

                On Error Resume Next

                Set FromObj = Msg.Item("from")

                If Not FromObj Is Nothing Then
                    Set UserObj = FromObj.Item("user")
                End If

                If Not UserObj Is Nothing Then
                    Sender = CStr(UserObj.Item("displayName"))
                End If

                Set BodyObj = Msg.Item("body")

                If Not BodyObj Is Nothing Then
                    MessageText = CleanTeamsHTML( _
                        CStr(BodyObj.Item("content")))
                End If

                On Error GoTo 0

                wsM.Cells(OutputRow, "A").Value = SelectedTeam
                wsM.Cells(OutputRow, "B").Value = SelectedChannel
                wsM.Cells(OutputRow, "C").Value = dateValue(MsgDateTime)
                wsM.Cells(OutputRow, "D").Value = TimeValue(MsgDateTime)
                wsM.Cells(OutputRow, "E").Value = Sender
                wsM.Cells(OutputRow, "F").Value = MessageText
                wsM.Cells(OutputRow, "G").Value = "Post"
                wsM.Cells(OutputRow, "H").Value = ""
                wsM.Cells(OutputRow, "I").Value = MessageID
                wsM.Cells(OutputRow, "J").Value = MsgDateTime

                OutputRow = OutputRow + 1

            End If

            '-----------------------------
            'EXPANDED REPLIES
            '-----------------------------
            Set Replies = Nothing

            On Error Resume Next
            Set Replies = Msg.Item("replies")
            On Error GoTo 0

            If Not Replies Is Nothing Then

                For Each Reply In Replies

                    MsgDateTime = 0
                    Sender = ""
                    MessageText = ""

                    If Reply.Exists("createdDateTime") Then
                        MsgDateTime = GraphDateTime( _
                            CStr(Reply.Item("createdDateTime")))
                    End If

                    If MsgDateTime >= FromDate And _
                       MsgDateTime < ToDate + 1 Then

                        Set FromObj = Nothing
                        Set UserObj = Nothing
                        Set BodyObj = Nothing

                        On Error Resume Next

                        Set FromObj = Reply.Item("from")

                        If Not FromObj Is Nothing Then
                            Set UserObj = FromObj.Item("user")
                        End If

                        If Not UserObj Is Nothing Then
                            Sender = CStr(UserObj.Item("displayName"))
                        End If

                        Set BodyObj = Reply.Item("body")

                        If Not BodyObj Is Nothing Then
                            MessageText = CleanTeamsHTML( _
                                CStr(BodyObj.Item("content")))
                        End If

                        On Error GoTo 0

                        wsM.Cells(OutputRow, "A").Value = SelectedTeam
                        wsM.Cells(OutputRow, "B").Value = SelectedChannel
                        wsM.Cells(OutputRow, "C").Value = dateValue(MsgDateTime)
                        wsM.Cells(OutputRow, "D").Value = TimeValue(MsgDateTime)
                        wsM.Cells(OutputRow, "E").Value = Sender
                        wsM.Cells(OutputRow, "F").Value = MessageText
                        wsM.Cells(OutputRow, "G").Value = "Reply"
                        wsM.Cells(OutputRow, "H").Value = MessageID
                        wsM.Cells(OutputRow, "I").Value = _
                            CStr(Reply.Item("id"))
                        wsM.Cells(OutputRow, "J").Value = MsgDateTime

                        OutputRow = OutputRow + 1

                    End If

                Next Reply

            End If

        Next Msg

        NextLink = ""

        If Json.Exists("@odata.nextLink") Then
            NextLink = CStr(Json.Item("@odata.nextLink"))
        End If

        Url = NextLink

    Loop

    wsM.Columns("C").NumberFormat = "dd-mmm-yyyy"
    wsM.Columns("D").NumberFormat = "hh:mm:ss"
    wsM.Columns("J").NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox OutputRow - 2 & _
           " messages imported successfully.", _
           vbInformation

End Sub


Public Sub ImportRepliesForMessage( _
    ByVal TeamID As String, _
    ByVal ChannelID As String, _
    ByVal ParentMessageID As String, _
    ByVal SelectedTeam As String, _
    ByVal SelectedChannel As String, _
    ByVal FromDate As Date, _
    ByVal ToDate As Date, _
    ByRef OutputRow As Long)

    Dim wsM As Worksheet
    Dim Url As String
    Dim Response As String
    Dim NextLink As String

    Dim Json As Object
    Dim RepliesCollection As Object
    Dim Reply As Variant

    Dim ReplyDateTime As Date
    Dim Sender As String
    Dim MessageText As String
    Dim MessageID As String

    Dim FromObj As Object
    Dim UserObj As Object
    Dim BodyObj As Object

    Set wsM = ThisWorkbook.Worksheets("Messages")

    Url = "https://graph.microsoft.com/v1.0/teams/" & _
          TeamID & "/channels/" & _
          ChannelID & "/messages/" & _
          ParentMessageID & "/replies?$top=50"

    Do While Url <> ""

        Response = HttpGet(Url, GraphAccessToken)

        Set Json = JsonConverter.ParseJson(Response)
        Set RepliesCollection = Json.Item("value")

        For Each Reply In RepliesCollection

            ReplyDateTime = 0
            Sender = ""
            MessageText = ""
            MessageID = ""

            If Reply.Exists("createdDateTime") Then
                ReplyDateTime = GraphDateTime( _
                    CStr(Reply.Item("createdDateTime")))
            End If

            If ReplyDateTime >= FromDate And _
               ReplyDateTime < ToDate + 1 Then

                'Sender
                Set FromObj = Nothing
                Set UserObj = Nothing

                On Error Resume Next

                Set FromObj = Reply.Item("from")

                If Not FromObj Is Nothing Then
                    Set UserObj = FromObj.Item("user")
                End If

                If Not UserObj Is Nothing Then
                    Sender = UserObj.Item("displayName")
                End If

                On Error GoTo 0

                'Body
                Set BodyObj = Nothing

                On Error Resume Next
                Set BodyObj = Reply.Item("body")

                If Not BodyObj Is Nothing Then
                    MessageText = CleanTeamsHTML(CStr(BodyObj.Item("content")))
                End If

                On Error GoTo 0

                If Reply.Exists("id") Then
                    MessageID = Reply.Item("id")
                End If

                'Write reply
                wsM.Cells(OutputRow, "A").Value = SelectedTeam
                wsM.Cells(OutputRow, "B").Value = SelectedChannel
                wsM.Cells(OutputRow, "C").Value = dateValue(ReplyDateTime)
                wsM.Cells(OutputRow, "D").Value = TimeValue(ReplyDateTime)
                wsM.Cells(OutputRow, "E").Value = Sender
                wsM.Cells(OutputRow, "F").Value = MessageText

                wsM.Cells(OutputRow, "G").Value = "Reply"
                wsM.Cells(OutputRow, "H").Value = ParentMessageID
                wsM.Cells(OutputRow, "I").Value = MessageID
                wsM.Cells(OutputRow, "J").Value = ReplyDateTime

                OutputRow = OutputRow + 1

            End If

        Next Reply

        NextLink = ""

        If Json.Exists("@odata.nextLink") Then
            NextLink = Json.Item("@odata.nextLink")
        End If

        Url = NextLink

    Loop

End Sub


Public Function CleanTeamsHTML(ByVal Text As String) As String

    Dim s As String

    s = Text

    s = Replace(s, "<br>", vbLf)
    s = Replace(s, "<br/>", vbLf)
    s = Replace(s, "<br />", vbLf)
    s = Replace(s, "</p>", vbLf)

    Do While InStr(s, "<") > 0 And InStr(s, ">") > 0
        s = Left$(s, InStr(s, "<") - 1) & _
            Mid$(s, InStr(s, ">") + 1)
    Loop

    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    
    s = Trim$(s)

    CleanTeamsHTML = s

End Function



Public Sub TestChannelDetails()

    Dim wsL As Worksheet
    Dim wsM As Worksheet
    Dim TeamID As String
    Dim ChannelID As String
    Dim SelectedTeam As String
    Dim SelectedChannel As String
    Dim i As Long
    Dim LastRow As Long
    Dim Url As String
    Dim Response As String

    Set wsL = ThisWorkbook.Worksheets("Lists")
    Set wsM = ThisWorkbook.Worksheets("Messages")

    SelectedTeam = wsM.Range("P3").Value
    SelectedChannel = wsM.Range("P4").Value

    LastRow = wsL.Cells(wsL.rows.Count, "A").End(xlUp).Row

    For i = 2 To LastRow
        If wsL.Cells(i, "A").Value = SelectedTeam Then
            TeamID = wsL.Cells(i, "B").Value
            Exit For
        End If
    Next i

    LastRow = wsL.Cells(wsL.rows.Count, "D").End(xlUp).Row

    For i = 2 To LastRow
        If wsL.Cells(i, "D").Value = SelectedChannel Then
            ChannelID = wsL.Cells(i, "E").Value
            Exit For
        End If
    Next i

    GraphAccessToken = wsL.Range("Z6").Value

    Url = "https://graph.microsoft.com/v1.0/teams/" & _
          TeamID & "/channels/" & ChannelID

    Response = HttpGet(Url, GraphAccessToken)

    wsL.Range("Z9").Value = Response

    MsgBox "Channel details saved in Lists!Z9."

End Sub


Public Function RefreshAccessToken() As Boolean

    Dim TokenUrl As String
    Dim PostData As String
    Dim Response As String
    Dim NewAccessToken As String
    Dim NewRefreshToken As String

    GraphRefreshToken = ThisWorkbook.Worksheets("Lists").Range("Z10").Value

    If GraphRefreshToken = "" Then
        RefreshAccessToken = False
        Exit Function
    End If

    TokenUrl = "https://login.microsoftonline.com/" & _
               TENANT_ID & _
               "/oauth2/v2.0/token"

    PostData = _
        "client_id=" & UrlEncode(CLIENT_ID) & _
        "&grant_type=refresh_token" & _
        "&refresh_token=" & UrlEncode(GraphRefreshToken) & _
        "&scope=" & UrlEncode( _
            "Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All offline_access")

    Response = HttpPost(TokenUrl, PostData)

    NewAccessToken = GetJsonValue(Response, "access_token")
    NewRefreshToken = GetJsonValue(Response, "refresh_token")

    If NewAccessToken <> "" Then

        GraphAccessToken = NewAccessToken
        ThisWorkbook.Worksheets("Lists").Range("Z6").Value = GraphAccessToken

        'Microsoft may return a new refresh token
        If NewRefreshToken <> "" Then
            GraphRefreshToken = NewRefreshToken
            ThisWorkbook.Worksheets("Lists").Range("Z10").Value = GraphRefreshToken
        End If

        RefreshAccessToken = True

    Else

        RefreshAccessToken = False

    End If

End Function


Public Sub SignInToMicrosoft()

    Dim LoginUrl As String

    'Create NEW PKCE verifier, challenge and state
    LoginUrl = GetAuthorizationUrl()

#If Mac Then

    'Remove any old redirect URL
    ThisWorkbook.Worksheets("Lists").Range("Z2").ClearContents
    ThisWorkbook.Worksheets("Lists").Range("Z3").ClearContents

    MsgBox _
        "The Microsoft login page will now open." & vbCrLf & vbCrLf & _
        "Complete the Microsoft sign-in." & vbCrLf & _
        "When the localhost page appears, copy the COMPLETE URL." & vbCrLf & vbCrLf & _
        "Then paste it into Lists!Z2 and run COMPLETE SIGN IN.", _
        vbInformation, _
        "Microsoft Sign In"

    ThisWorkbook.FollowHyperlink LoginUrl

    'IMPORTANT:
    'Stop here on Mac.
    'Do NOT exchange the authorization code yet.
    Exit Sub

#Else

    'Windows login process
    ThisWorkbook.FollowHyperlink LoginUrl

    frmMicrosoftLogin.Show

    If frmMicrosoftLogin.LoginCancelled Then
        Unload frmMicrosoftLogin
        Exit Sub
    End If

    Unload frmMicrosoftLogin

    CaptureAuthorizationCode
    ExchangeCodeForToken
    PopulateTeams

    MsgBox "Microsoft sign-in completed successfully.", vbInformation

#End If

End Sub



Public Sub CompleteMacSignIn()

#If Mac Then

    If Trim$(ThisWorkbook.Worksheets("Lists").Range("Z2").Value) = "" Then

        MsgBox _
            "Please paste the COMPLETE localhost URL into Lists!Z2 first.", _
            vbExclamation

        Exit Sub

    End If

    'Extract authorization code and validate state
    CaptureAuthorizationCode

    If OAuthAuthorizationCode = "" Then
        MsgBox "Authorization code could not be captured.", vbCritical
        Exit Sub
    End If

    'Exchange authorization code using the SAME PKCE verifier
    ExchangeCodeForToken

    GraphAccessToken = _
        ThisWorkbook.Worksheets("Lists").Range("Z6").Value

    If GraphAccessToken = "" Then
        MsgBox "Microsoft sign-in failed.", vbCritical
        Exit Sub
    End If

    'Load fresh Teams
    PopulateTeams

    MsgBox _
        "Microsoft sign-in completed successfully.", _
        vbInformation

#Else

    MsgBox "This procedure is only required on Mac.", vbInformation

#End If

End Sub

























