Attribute VB_Name = "ApiConnection"
Option Explicit

Private XmlHttp As Object
Private Endpoint As String
Private Parameters As String

Public Sub ApiEndpointBackupCharfiles()
    Endpoint = ApiUrlServer & "/api/v1/charfiles/backupcharfiles"
    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointBackupCuentas()
    Endpoint = ApiUrlServer & "/api/v1/accounts/backupaccountfiles"
    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointBackupLogs()
    Endpoint = ApiUrlServer & "/api/v1/logs/backuplogs"
    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointSendWelcomeEmail(ByVal UserName As String, ByVal Password As String, ByVal Email As String)
    Endpoint = ApiUrlServer & "/api/v1/emails/welcome"
    Parameters = "username=" & UserName & "&password=" & Password & "&emailTo=" & Email
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendLoginAccountEmail(ByVal Email As String, ByVal LastIpsUsed As String, ByVal CurrentIp As String)
    Endpoint = ApiUrlServer & "/api/v1/emails/loginAccount"
    Parameters = "emailTo=" & Email & "&lastIpsUsed=" & LastIpsUsed & "&currentIp=" & CurrentIp
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendResetPasswordAccountEmail(ByVal Email As String, ByVal NewPassword As String)
    Endpoint = ApiUrlServer & "/api/v1/emails/resetAccountPassword"
    Parameters = "newPassword=" & NewPassword & "&emailTo=" & Email
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendUserConnectedMessageDiscord(ByVal UserName As String, ByVal Desc As String, ByVal EsCriminal As Boolean, ByVal Clase As String)
    Endpoint = ApiUrlServer & "/api/v1/discord/sendConnectedMessage"
    Parameters = "userName=" & UserName & "&desc=" & Desc & "&esCriminal=" & EsCriminal & "&clase=" & Clase
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendHappyHourStartedMessageDiscord(ByVal Message As String)
    Endpoint = ApiUrlServer & "/api/v1/discord/sendHappyHourStartMessage"
    Parameters = "message=" & Message
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendHappyHourEndedMessageDiscord(ByVal Message As String)
    Endpoint = ApiUrlServer & "/api/v1/discord/sendHappyHourEndMessage"
    Parameters = "message=" & Message
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendHappyHourModifiedMessageDiscord(ByVal Message As String)
    Endpoint = ApiUrlServer & "/api/v1/discord/sendHappyHourModifiedMessage"
    Parameters = "message=" & Message
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendNewGuildCreatedMessageDiscord(ByVal Message As String, ByVal Desc As String, ByVal GuildName As String, ByVal Site As String)
    Endpoint = ApiUrlServer & "/api/v1/discord/sendNewGuildCreated"
    Parameters = "message=" & Message & "&desc=" & Desc & "&guildname=" & GuildName & "&site=" & Site
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendCustomCharacterMessageDiscord(ByVal Chat As String, ByVal Name As String, ByVal Desc As String)
    Endpoint = ApiUrlServer & "/api/v1/discord/sendCustomCharacterMessageDiscord"
    Parameters = "userName=" & Name & "&desc=" & Desc & "&chat=" & Chat
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendWorldSaveMessageDiscord()
    Endpoint = ApiUrlServer & "/api/v1/discord/sendWorldSaveMessage"
    Parameters = ""
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendCreateNewCharacterMessageDiscord(ByVal Name As String)
    Endpoint = ApiUrlServer & "/api/v1/discord/sendCreatedNewCharacterMessage"
    Parameters = "name=" & Name
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Public Sub ApiEndpointSendServerDataToApiToShowOnlineUsers()
    Endpoint = "https://api.argentumonline.org/api/v1/servers/sendUsersOnline"
    Parameters = "serverName=" & NombreServidor & "&quantityUsers=" & LastUser & "&ip=" & IpPublicaServidor & "&port=" & Puerto
    Call SendPOSTRequest(Endpoint, Parameters)
End Sub

Private Sub SendPOSTRequest(ByVal Endpoint As String, ByVal Parameters As String)
On Error GoTo ErrorHandler
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    XmlHttp.Open "POST", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send CStr(Parameters)
    Set XmlHttp = Nothing
ErrorHandler:
    If Err.Number <> 0 Then
        Call LogError("Error POST endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.description)
    End If
End Sub

Private Sub SendGETRequest(ByVal Endpoint As String)
On Error GoTo ErrorHandler
    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    XmlHttp.Open "GET", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.send
    Set XmlHttp = Nothing
ErrorHandler:
    If Err.Number <> 0 Then
        Call LogError("Error GET endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.description)
    End If
End Sub
