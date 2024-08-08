VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultasPopulares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ARCHIVOMAILS = "\logs\votaron.dat"
Private Const ARCHIVOCONFIG = "\dat\consultas.dat"
Private pEncuestaActualNum As Integer
Private pEncuestaActualTex As String
Private pNivelRequerido    As Integer
Private pOpciones()        As Integer

Public Property Let Numero(ByVal NumEncuesta As Integer)
    pEncuestaActualNum = NumEncuesta
End Property

Public Property Get Numero() As Integer
    Numero = pEncuestaActualNum
End Property

Public Property Let texto(ByVal Descripcion As String)
    pEncuestaActualTex = Descripcion
End Property

Public Property Get texto() As String
    texto = pEncuestaActualTex
End Property

Public Sub LoadData()
    Dim CantOpciones As Integer
    Dim i            As Integer
    pEncuestaActualNum = val(GetVar(App.Path & ARCHIVOCONFIG, "INIT", "ConsultaActual"))
    pEncuestaActualTex = GetVar(App.Path & ARCHIVOCONFIG, "INIT", "ConsultaActualTexto")
    pNivelRequerido = GetVar(App.Path & ARCHIVOCONFIG, "INIT", "NivelRequerido")
    If pEncuestaActualNum > 0 Then
        CantOpciones = val(GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "CANTOPCIONES"))
        ReDim pOpciones(1 To CantOpciones) As Integer
        For i = 1 To CantOpciones
            pOpciones(i) = val(GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i))
        Next i
    End If
End Sub

Public Function doVotar(ByVal Userindex As Integer, ByVal opcion As Integer) As String
    On Error GoTo ErrorHandler
    Dim YaVoto   As Boolean
    Dim sufragio As Integer
    If pEncuestaActualNum = 0 Then
        doVotar = "No hay consultas populares abiertas"
        Exit Function
    End If
    If (UserList(Userindex).Stats.ELV >= pNivelRequerido) Then
        If (OpcionValida(opcion)) Then
            YaVoto = PersonajeCantidadVotos(UserList(Userindex).Name) >= pEncuestaActualNum
            If Not YaVoto Then
                If Not MailYaVoto(UserList(Userindex).Email) Then
                    sufragio = CLng(val(GetVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion)))
                    sufragio = sufragio + 1
                    Call WriteVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion, Str(sufragio))
                    doVotar = "Tu voto ha sido computado. Opcion: " & opcion
                    Call MarcarPjComoQueYaVoto(Userindex, pEncuestaActualNum)
                    Call MarcarMailComoQueYaVoto(UserList(Userindex).Email)
                Else
                    Call MarcarPjComoQueYaVoto(Userindex, pEncuestaActualNum)
                    doVotar = "Este email ya voto en la consulta: " & pEncuestaActualTex
                End If
            Else
                doVotar = "Este personaje ya voto en la consulta: " & pEncuestaActualTex
            End If
        Else
            doVotar = "Esa no es una opcion para votar"
        End If
    Else
        doVotar = "Para votar en esta consulta debes ser nivel " & pNivelRequerido & " o superior"
    End If
    Exit Function
ErrorHandler:
    Call LogError("Error en ConsultasPopularse.doVotar: " & Err.description)
End Function

Public Function SendInfoEncuesta(ByVal Userindex As Integer) As String
    Dim i As Integer
    Call WriteConsoleMsg(Userindex, "CONSULTA POPULAR NUMERO " & pEncuestaActualNum, FontTypeNames.FONTTYPE_GUILD)
    Call WriteConsoleMsg(Userindex, pEncuestaActualTex, FontTypeNames.FONTTYPE_GUILD)
    Call WriteConsoleMsg(Userindex, " Opciones de voto: ", FontTypeNames.FONTTYPE_GUILDMSG)
    For i = 1 To UBound(pOpciones)
        Call WriteConsoleMsg(Userindex, "(Opcion " & i & "): " & GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i), FontTypeNames.FONTTYPE_GUILDMSG)
    Next i
    Call WriteConsoleMsg(Userindex, " Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1. Tu voto no podra ser cambiado.", FontTypeNames.FONTTYPE_VENENO)
End Function

Private Sub MarcarPjComoQueYaVoto(ByVal Userindex As Integer, ByVal NumeroEncuesta As Integer)
    If Not Database_Enabled Then
        Call MarcarPjComoQueYaVotoCharfile(Userindex, NumeroEncuesta)
    Else
        Call MarcarPjComoQueYaVotoDatabase(Userindex, NumeroEncuesta)
    End If
End Sub

Private Function PersonajeCantidadVotos(ByVal UserName As String) As Integer
    If Not Database_Enabled Then
        PersonajeCantidadVotos = PersonajeCantidadVotosCharfile(UserName)
    Else
        PersonajeCantidadVotos = PersonajeCantidadVotosDatabase(UserName)
    End If
End Function

Private Function MailYaVoto(ByVal Email As String) As Boolean
    Dim ArchN As Integer
    Dim Tmp   As String
    MailYaVoto = False
    ArchN = FreeFile()
    Open App.Path & ARCHIVOMAILS For Input As #ArchN
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        If Email = Tmp Then
            MailYaVoto = True
            Close #ArchN
            Exit Function
        End If
    Loop
    Close #ArchN
End Function

Private Sub MarcarMailComoQueYaVoto(ByVal Email As String)
    Dim ArchN As Integer
    ArchN = FreeFile()
    Open App.Path & ARCHIVOMAILS For Append As #ArchN
    Print #ArchN, Email
    Close #ArchN
End Sub

Private Function OpcionValida(ByVal opcion As Integer) As Boolean
    OpcionValida = opcion > 0 And opcion <= UBound(pOpciones)
End Function
