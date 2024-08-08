Attribute VB_Name = "Logs"
Option Explicit

Public Sub LogBan(ByVal BannedIndex As Integer, ByVal Userindex As Integer, ByVal Motivo As String)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(Userindex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile
End Sub

Public Sub LogBanFromName(ByVal BannedName As String, ByVal Userindex As Integer, ByVal Motivo As String)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(Userindex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile
End Sub

Public Sub LogServerStartTime()
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & time & " server iniciado " & GetVersionOfTheServer()
    Close #n
End Sub

Public Sub LogCriticEvent(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogIndex(ByVal index As Integer, ByVal Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogError(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    Exit Sub

ErrorHandler:

End Sub

Public Sub LogRetos(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\Retos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogDatabaseError(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
        nfile = FreeFile
    Open App.Path & "\logs\database.log" For Append Shared As #nfile
        Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    Exit Sub
    Debug.Print Desc
ErrorHandler:

End Sub

Public Sub LogStatic(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogTarea(Desc As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile(1)
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogClanes(ByVal Str As String)
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Str
    Close #nfile
End Sub

Public Sub LogIP(ByVal Str As String)
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Str
    Close #nfile
End Sub

Public Sub LogDesarrollo(ByVal Str As String)
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Str
    Close #nfile
End Sub

Public Sub LogGM(Nombre As String, texto As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogAsesinato(texto As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogHackAttemp(texto As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogCheating(texto As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub LogAntiCheat(texto As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
    nfile = FreeFile
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, ""
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub
