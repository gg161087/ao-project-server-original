Attribute VB_Name = "modGuilds"
Option Explicit

Private GUILDINFOFILE             As String
Private Const MAX_GUILDS          As Integer = 1000
Public CANTIDADDECLANES           As Integer
Public guilds(1 To MAX_GUILDS)    As clsClan
Private Const CANTIDADMAXIMACODEX As Byte = 8
Public Const MAXASPIRANTES        As Byte = 10
Private Const MAXANTIFACCION      As Byte = 5

Public Enum ALINEACION_GUILD
    ALINEACION_LEGION = 1
    ALINEACION_CRIMINAL = 2
    ALINEACION_NEUTRO = 3
    ALINEACION_CIUDA = 4
    ALINEACION_ARMADA = 5
    ALINEACION_MASTER = 6
End Enum

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum

Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum

Public Sub LoadGuildsDB()
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando guildsinfo.inf."
    Dim CantClanes As String
    Dim i          As Integer
    Dim TempStr    As String
    Dim Alin       As ALINEACION_GUILD
    GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"
    CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    If IsNumeric(CantClanes) Then
        CANTIDADDECLANES = CInt(CantClanes)
    Else
        CANTIDADDECLANES = 0
    End If
    For i = 1 To CANTIDADDECLANES
        Set guilds(i) = New clsClan
        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
        Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
        Call guilds(i).Inicializar(TempStr, i, Alin)
    Next i
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo guildsinfo.inf."
End Sub

Public Function m_ConectarMiembroAClan(ByVal Userindex As Integer, ByVal GuildIndex As Integer) As Boolean
    Dim NuevaA As Boolean
    Dim News   As String
    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function
    If m_EstadoPermiteEntrar(Userindex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(Userindex)
        UserList(Userindex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(Userindex, True, NuevaA)
        If NuevaA Then News = News & "El clan tiene nueva alineacion."
    End If
End Function

Public Function m_ValidarPermanencia(ByVal Userindex As Integer, ByVal SumaAntifaccion As Boolean, ByRef CambioAlineacion As Boolean) As Boolean
    Dim GuildIndex As Integer
    m_ValidarPermanencia = True
    GuildIndex = UserList(Userindex).GuildIndex
    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    If Not m_EstadoPermiteEntrar(Userindex, GuildIndex) Then
        If m_EsGuildLeader(UserList(Userindex).Name, GuildIndex) Then
            Call LogClanes(UserList(Userindex).Name & ", lider de " & guilds(GuildIndex).GuildName & " hizo bajar la alienacion de su clan.")
            CambioAlineacion = True
            Do
                Call UpdateGuildMembers(GuildIndex)
            Loop Until m_EstadoPermiteEntrar(Userindex, GuildIndex)
        Else
            Call LogClanes(UserList(Userindex).Name & " de " & guilds(GuildIndex).GuildName & " es expulsado en validar permanencia.")
            m_ValidarPermanencia = False
            If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
            CambioAlineacion = guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            Call LogClanes(UserList(Userindex).Name & " de " & guilds(GuildIndex).GuildName & IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alineacion. MAXANT:" & CambioAlineacion)
            Call m_EcharMiembroDeClan(-1, UserList(Userindex).Name)
            If CambioAlineacion Then
                Call UpdateGuildMembers(GuildIndex)
            End If
        End If
    End If
End Function

Private Sub UpdateGuildMembers(ByVal GuildIndex As Integer)
    Dim GuildMembers() As String
    Dim TotalMembers   As Integer
    Dim MemberIndex    As Long
    Dim Sale           As Boolean
    Dim MemberName     As String
    Dim Userindex      As Integer
    Dim Reenlistadas   As Integer
    If guilds(GuildIndex).CambiarAlineacion(BajarGrado(GuildIndex)) Then
        GuildMembers = guilds(GuildIndex).GetMemberList()
        TotalMembers = UBound(GuildMembers)
        For MemberIndex = 0 To TotalMembers
            MemberName = GuildMembers(MemberIndex)
            Userindex = NameIndex(MemberName)
            If Userindex > 0 Then
                Sale = Not m_EstadoPermiteEntrar(Userindex, GuildIndex)
            Else
                Sale = Not m_EstadoPermiteEntrarChar(MemberName, GuildIndex)
            End If
            If Sale Then
                If m_EsGuildLeader(MemberName, GuildIndex) Then
                    If Userindex > 0 Then
                        If UserList(Userindex).Faccion.ArmadaReal <> 0 Then
                            Call ExpulsarFaccionReal(Userindex)
                            UserList(Userindex).Faccion.Reenlistadas = UserList(Userindex).Faccion.Reenlistadas - 1
                        ElseIf UserList(Userindex).Faccion.FuerzasCaos <> 0 Then
                            Call ExpulsarFaccionCaos(Userindex)
                            UserList(Userindex).Faccion.Reenlistadas = UserList(Userindex).Faccion.Reenlistadas - 1
                        End If
                    Else
                        If PersonajeExiste(MemberName) Then
                            Call KickUserFacciones(MemberName)
                            Reenlistadas = GetUserReenlists(MemberName)
                            Call SaveUserReenlists(MemberName, IIf(Reenlistadas > 1, Reenlistadas - 1, Reenlistadas))
                        End If
                    End If
                Else
                    Call m_EcharMiembroDeClan(-1, MemberName)
                End If
            End If
        Next MemberIndex
    Else
        guilds(GuildIndex).PuntosAntifaccion = 0
    End If
End Sub

Private Function BajarGrado(ByVal GuildIndex As Integer) As ALINEACION_GUILD
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_ARMADA
            BajarGrado = ALINEACION_CIUDA

        Case ALINEACION_LEGION
            BajarGrado = ALINEACION_CRIMINAL

        Case Else
            BajarGrado = ALINEACION_NEUTRO
    End Select
End Function

Public Sub m_DesconectarMiembroDelClan(ByVal Userindex As Integer, ByVal GuildIndex As Integer)
    If UserList(Userindex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(Userindex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
    Dim Userindex As Integer
    Dim GI        As Integer
    m_EcharMiembroDeClan = 0
    Userindex = NameIndex(Expulsado)
    If Userindex > 0 Then
        GI = UserList(Userindex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                Call guilds(GI).DesConectarMiembro(Userindex)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(Userindex).GuildIndex = 0
                Call RefreshCharStatus(Userindex)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        GI = GetUserGuildIndex(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If
End Function

Public Sub ActualizarWebSite(ByVal Userindex As Integer, ByRef Web As String)
    Dim GI As Integer
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then Exit Sub
    Call guilds(GI).SetURL(Web)
End Sub

Public Sub ChangeCodexAndDesc(ByRef Desc As String, ByRef codex() As String, ByVal GuildIndex As Integer)
    Dim i As Long
    If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    With guilds(GuildIndex)
        Call .SetDesc(Desc)
        For i = 0 To UBound(codex())
            Call .SetCodex(i, codex(i))
        Next i
        For i = i To CANTIDADMAXIMACODEX
            Call .SetCodex(i, vbNullString)
        Next i
    End With
End Sub

Public Sub ActualizarNoticias(ByVal Userindex As Integer, ByRef Datos As String)
    Dim GI As Integer
    With UserList(Userindex)
        GI = .GuildIndex
        If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
        If Not m_EsGuildLeader(.Name, GI) Then Exit Sub
        Call guilds(GI).SetGuildNews(Datos)
        Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & " ha actualizado las noticias del clan!"))
    End With
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef Desc As String, ByRef GuildName As String, ByRef URL As String, ByRef codex() As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
    Dim CantCodex   As Integer
    Dim i           As Integer
    Dim DummyString As String
    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If
    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan invalido."
        Exit Function
    End If
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function
    End If
    CantCodex = UBound(codex()) + 1
    If CANTIDADDECLANES < UBound(guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        Set guilds(CANTIDADDECLANES) = New clsClan
        With guilds(CANTIDADDECLANES)
            Call .Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
            Call .InicializarNuevoClan(UserList(FundadorIndex).Name)
            For i = 1 To CantCodex
                Call .SetCodex(i, codex(i - 1))
            Next i
            Call .SetDesc(Desc)
            Call .SetGuildNews("Clan creado con alineacion: " & Alineacion2String(Alineacion))
            Call .SetLeader(UserList(FundadorIndex).Name)
            Call .SetURL(URL)
            Call .AceptarNuevoMiembro(UserList(FundadorIndex).Name)
            Call .ConectarMiembro(FundadorIndex)
        End With
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call RefreshCharStatus(FundadorIndex)
        For i = 1 To CANTIDADDECLANES - 1
            Call guilds(i).ProcesarFundacionDeOtroClan
        Next i
    Else
        refError = "No hay mas slots para fundar clanes. Consulte a un administrador."
        Exit Function
    End If
    CrearNuevoClan = True
End Function

Public Sub SendGuildNews(ByVal Userindex As Integer)
    Dim GuildIndex As Integer
    Dim i          As Integer
    Dim go         As Integer
    GuildIndex = UserList(Userindex).GuildIndex
    If GuildIndex = 0 Then Exit Sub
    Dim enemies() As String
    With guilds(GuildIndex)
        If .CantidadEnemys Then
            ReDim enemies(0 To .CantidadEnemys - 1) As String
        Else
            ReDim enemies(0)
        End If
        Dim allies() As String
        If .CantidadAllies Then
            ReDim allies(0 To .CantidadAllies - 1) As String
        Else
            ReDim allies(0)
        End If
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
        go = 0
        While i > 0
            enemies(go) = guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
            go = go + 1
        Wend
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
        go = 0
        While i > 0
            allies(go) = guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
        Wend
        Call WriteGuildNews(Userindex, .GetGuildNews, enemies, allies)
        If .EleccionesAbiertas Then
            Call WriteConsoleMsg(Userindex, "Hoy es la votacion para elegir un nuevo lider para el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(Userindex, "La eleccion durara 24 horas, se puede votar a cualquier miembro del clan.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(Userindex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(Userindex, "Solo se computara un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Then Exit Function
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If
    If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then
                Exit Function
            End If
        End If
    End If
    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).GetLeader) <> UCase$(Nombre)
End Function

Public Function PuedeFundarUnClan(ByVal Userindex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
    If UserList(Userindex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    If UserList(Userindex).Stats.ELV < 25 Or UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
        refError = "Para fundar un clan debes ser nivel 25 y tener 90 skills en liderazgo."
        Exit Function
    End If
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            If UserList(Userindex).Faccion.ArmadaReal <> 1 Then
                refError = "Para fundar un clan real debes ser miembro del ejercito real."
                Exit Function
            End If

        Case ALINEACION_GUILD.ALINEACION_CIUDA
            If criminal(Userindex) Then
                refError = "Para fundar un clan de ciudadanos no debes ser criminal."
                Exit Function
            End If

        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            If Not criminal(Userindex) Then
                refError = "Para fundar un clan de criminales no debes ser ciudadano."
                Exit Function
            End If

        Case ALINEACION_GUILD.ALINEACION_LEGION
            If UserList(Userindex).Faccion.FuerzasCaos <> 1 Then
                refError = "Para fundar un clan del mal debes pertenecer a la legion oscura."
                Exit Function
            End If

        Case ALINEACION_GUILD.ALINEACION_MASTER
            If UserList(Userindex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                refError = "Para fundar un clan sin alineacion debes ser un dios."
                Exit Function
            End If

        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            If UserList(Userindex).Faccion.ArmadaReal <> 0 Or UserList(Userindex).Faccion.FuerzasCaos <> 0 Then
                refError = "Para fundar un clan neutro no debes pertenecer a ninguna faccion."
                Exit Function
            End If
    End Select
    PuedeFundarUnClan = True
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
    Dim Promedio As Long
    Dim ELV      As Integer
    Dim f        As Byte
    m_EstadoPermiteEntrarChar = False
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace(Personaje, ".", vbNullString)
    End If
    If PersonajeExiste(Personaje) Then
        Promedio = GetUserPromedio(Personaje)
        Select Case guilds(GuildIndex).Alineacion
            Case ALINEACION_GUILD.ALINEACION_ARMADA
                If Promedio >= 0 Then
                    ELV = GetUserLevel(Personaje)
                    If ELV >= 25 Then
                        f = UserBelongsToRoyalArmy(Personaje)
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f, True)
                End If

            Case ALINEACION_GUILD.ALINEACION_CIUDA
                m_EstadoPermiteEntrarChar = Promedio >= 0

            Case ALINEACION_GUILD.ALINEACION_CRIMINAL
                m_EstadoPermiteEntrarChar = Promedio < 0

            Case ALINEACION_GUILD.ALINEACION_NEUTRO
                m_EstadoPermiteEntrarChar = (UserBelongsToRoyalArmy(Personaje) = False And UserBelongsToChaosLegion(Personaje) = False)

            Case ALINEACION_GUILD.ALINEACION_LEGION
                If Promedio < 0 Then
                    ELV = GetUserLevel(Personaje)
                    If ELV >= 25 Then
                        f = UserBelongsToChaosLegion(Personaje)
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f, True)
                End If
                
            Case Else
                m_EstadoPermiteEntrarChar = True
        End Select
    End If
End Function

Private Function m_EstadoPermiteEntrar(ByVal Userindex As Integer, ByVal GuildIndex As Integer) As Boolean
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Not criminal(Userindex) And IIf(UserList(Userindex).Stats.ELV >= 25, UserList(Userindex).Faccion.ArmadaReal <> 0, True)
        
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = criminal(Userindex) And IIf(UserList(Userindex).Stats.ELV >= 25, UserList(Userindex).Faccion.FuerzasCaos <> 0, True)
        
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            m_EstadoPermiteEntrar = UserList(Userindex).Faccion.ArmadaReal = 0 And UserList(Userindex).Faccion.FuerzasCaos = 0
        
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            m_EstadoPermiteEntrar = Not criminal(Userindex)
        
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = criminal(Userindex)
        
        Case Else
            m_EstadoPermiteEntrar = True
    End Select
End Function

Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    Select Case S
        Case "Neutral"
            String2Alineacion = ALINEACION_NEUTRO

        Case "Del Mal"
            String2Alineacion = ALINEACION_LEGION

        Case "Real"
            String2Alineacion = ALINEACION_ARMADA

        Case "Game Masters"
            String2Alineacion = ALINEACION_MASTER

        Case "Legal"
            String2Alineacion = ALINEACION_CIUDA

        Case "Criminal"
            String2Alineacion = ALINEACION_CRIMINAL
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            Alineacion2String = "Neutral"

        Case ALINEACION_GUILD.ALINEACION_LEGION
            Alineacion2String = "Del Mal"

        Case ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Real"

        Case ALINEACION_GUILD.ALINEACION_MASTER
            Alineacion2String = "Game Masters"

        Case ALINEACION_GUILD.ALINEACION_CIUDA
            Alineacion2String = "Legal"

        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"
    End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    Select Case Relacion
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "A"

        Case RELACIONES_GUILD.GUERRA
            Relacion2String = "G"

        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"

        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "?"
    End Select
End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
    Select Case UCase$(Trim$(S))
        Case vbNullString, "P"
            String2Relacion = RELACIONES_GUILD.PAZ

        Case "G"
            String2Relacion = RELACIONES_GUILD.GUERRA

        Case "A"
            String2Relacion = RELACIONES_GUILD.ALIADOS

        Case Else
            String2Relacion = RELACIONES_GUILD.PAZ
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i   As Integer
    cad = LCase$(cad)
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            GuildNameValido = False
            Exit Function
        End If
    Next i
    GuildNameValido = True
End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
    Dim i As Integer
    YaExiste = False
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        YaExiste = (UCase$(guilds(i).GuildName) = GuildName)
        If YaExiste Then Exit Function
    Next i
End Function

Public Function HasFound(ByRef UserName As String) As Boolean
    Dim i    As Long
    Dim Name As String
    Name = UCase$(UserName)
    For i = 1 To CANTIDADDECLANES
        HasFound = (UCase$(guilds(i).Fundador) = Name)
        If HasFound Then Exit Function
    Next i
End Function

Public Function v_AbrirElecciones(ByVal Userindex As Integer, ByRef refError As String) As Boolean
    Dim GuildIndex As Integer
    v_AbrirElecciones = False
    GuildIndex = UserList(Userindex).GuildIndex
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningun clan."
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GuildIndex) Then
        refError = "No eres el lider de tu clan"
        Exit Function
    End If
    If guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya estan abiertas."
        Exit Function
    End If
    v_AbrirElecciones = True
    Call guilds(GuildIndex).AbrirElecciones
End Function

Public Function v_UsuarioVota(ByVal Userindex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
    Dim GuildIndex As Integer
    Dim list()     As String
    Dim i          As Long
    v_UsuarioVota = False
    GuildIndex = UserList(Userindex).GuildIndex
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningun clan."
        Exit Function
    End If
    With guilds(GuildIndex)
        If Not .EleccionesAbiertas Then
            refError = "No hay elecciones abiertas en tu clan."
            Exit Function
        End If
        list = .GetMemberList()
        For i = 0 To UBound(list())
            If UCase$(Votado) = list(i) Then Exit For
        Next i
        If i > UBound(list()) Then
            refError = Votado & " no pertenece al clan."
            Exit Function
        End If
        If .YaVoto(UserList(Userindex).Name) Then
            refError = "Ya has votado, no puedes cambiar tu voto."
            Exit Function
        End If
        Call .ContabilizarVoto(UserList(Userindex).Name, Votado)
        v_UsuarioVota = True
    End With
End Function

Public Sub v_RutinaElecciones()
    On Error GoTo ErrorHandler
    Dim i As Integer
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To CANTIDADDECLANES
        If Not guilds(i) Is Nothing Then
            If guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & " es el nuevo lider de " & guilds(i).GuildName & ".", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If
proximo:
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Elecciones revisadas.", FontTypeNames.FONTTYPE_SERVER))
    Exit Sub
ErrorHandler:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.description)
    Resume proximo
End Sub

Public Function GuildIndex(ByRef GuildName As String) As Integer
    Dim i As Integer
    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        If UCase$(guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function m_ListaDeMiembrosOnline(ByVal Userindex As Integer, ByVal GuildIndex As Integer) As String
    Dim i    As Integer
    Dim priv As PlayerType
    priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
    If UserList(Userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
        priv = priv Or PlayerType.Dios Or PlayerType.Admin
    End If
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            If i <> Userindex And (UserList(i).flags.Privilegios And priv) Then
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
            End If
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function

Public Function PrepareGuildsList() As String()
    Dim tStr() As String
    Dim i      As Long
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        For i = 1 To CANTIDADDECLANES
            If LenB(guilds(i).GuildName) <> 0 Then
                If guilds(i).GuildName <> "CLAN CERRADO" Then
                    tStr(i - 1) = guilds(i).GuildName
                End If
            End If
        Next i
    End If
    PrepareGuildsList = tStr
End Function

Public Sub SendGuildDetails(ByVal Userindex As Integer, ByRef GuildName As String)
    Dim codex(CANTIDADMAXIMACODEX - 1) As String
    Dim GI                             As Integer
    Dim i                              As Long
    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Sub
    With guilds(GI)
        For i = 1 To CANTIDADMAXIMACODEX
            codex(i - 1) = .GetCodex(i)
        Next i
        Call Protocol.WriteGuildDetails(Userindex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, .GetURL, .CantidadDeMiembros, .EleccionesAbiertas, Alineacion2String(.Alineacion), .CantidadEnemys, .CantidadAllies, .PuntosAntifaccion & "/" & CStr(MAXANTIFACCION), codex, .GetDesc)
    End With
End Sub

Public Sub SendGuildLeaderInfo(ByVal Userindex As Integer)
    Dim GI              As Integer
    Dim guildList()     As String
    Dim MemberList()    As String
    Dim aspirantsList() As String
    With UserList(Userindex)
        GI = .GuildIndex
        guildList = PrepareGuildsList()
        If GI <= 0 Or GI > CANTIDADDECLANES Then
            Call WriteGuildList(Userindex, guildList)
            Exit Sub
        End If
        MemberList = guilds(GI).GetMemberList()
        If Not m_EsGuildLeader(.Name, GI) Then
            Call WriteGuildMemberInfo(Userindex, guildList, MemberList)
            Exit Sub
        End If
        aspirantsList = guilds(GI).GetAspirantes()
        Call WriteGuildLeaderInfo(Userindex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList)
    End With
End Sub

Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()
    End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
    r_Iterador_ProximaPropuesta = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)
    End If
End Function

Public Function GMEscuchaClan(ByVal Userindex As Integer, ByVal GuildName As String) As Integer
    Dim GI As Integer
    If LenB(GuildName) = 0 And UserList(Userindex).EscucheClan <> 0 Then
        Call WriteConsoleMsg(Userindex, "Dejas de escuchar a : " & guilds(UserList(Userindex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(Userindex).EscucheClan).DesconectarGM (Userindex)
        Exit Function
    End If
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        If UserList(Userindex).EscucheClan <> 0 Then
            If UserList(Userindex).EscucheClan = GI Then
                Call WriteConsoleMsg(Userindex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = GI
                Exit Function
            Else
                Call WriteConsoleMsg(Userindex, "Dejas de escuchar a : " & guilds(UserList(Userindex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(Userindex).EscucheClan).DesconectarGM (Userindex)
            End If
        End If
        Call guilds(GI).ConectarGM(Userindex)
        Call WriteConsoleMsg(Userindex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(Userindex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(Userindex, "Error, el clan no existe.", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0
    End If
End Function

Public Sub GMDejaDeEscucharClan(ByVal Userindex As Integer, ByVal GuildIndex As Integer)
    UserList(Userindex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(Userindex)
End Sub

Public Function r_DeclararGuerra(ByVal Userindex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
    Dim GI  As Integer
    Dim GIG As Integer
    r_DeclararGuerra = 0
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function
    End If
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function
    End If
    GIG = GuildIndex(GuildGuerra)
    If guilds(GI).GetRelacion(GIG) = GUERRA Then
        refError = "Tu clan ya esta en guerra con " & GuildGuerra & "."
        Exit Function
    End If
    If GI = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan."
        Exit Function
    End If
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)
    r_DeclararGuerra = GIG
End Function

Public Function r_AceptarPropuestaDePaz(ByVal Userindex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
    Dim GI  As Integer
    Dim GIG As Integer
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function
    End If
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function
    End If
    GIG = GuildIndex(GuildPaz)
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function
    End If
    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No estas en guerra con ese clan."
        Exit Function
    End If
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar."
        Exit Function
    End If
    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
    r_AceptarPropuestaDePaz = GIG
End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal Userindex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
    Dim GI  As Integer
    Dim GIG As Integer
    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function
    End If
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function
    End If
    GIG = GuildIndex(GuildPro)
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function
    End If
    If Not guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function
    End If
    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG
End Function

Public Function r_RechazarPropuestaDePaz(ByVal Userindex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
    Dim GI  As Integer
    Dim GIG As Integer
    r_RechazarPropuestaDePaz = 0
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function
    End If
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function
    End If
    GIG = GuildIndex(GuildPro)
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function
    End If
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function
    End If
    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG
End Function

Public Function r_AceptarPropuestaDeAlianza(ByVal Userindex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
    Dim GI  As Integer
    Dim GIG As Integer
    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function
    End If
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningun clan."
        Exit Function
    End If
    GIG = GuildIndex(GuildAllie)
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function
    End If
    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = "No estas en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function
    End If
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function
    End If
    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
    r_AceptarPropuestaDeAlianza = GIG
End Function

Public Function r_ClanGeneraPropuesta(ByVal Userindex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
    Dim OtroClanGI As Integer
    Dim GI         As Integer
    r_ClanGeneraPropuesta = False
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function
    End If
    OtroClanGI = GuildIndex(OtroClan)
    If OtroClanGI = GI Then
        refError = "No puedes declarar relaciones con tu propio clan."
        Exit Function
    End If
    If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe."
        Exit Function
    End If
    If guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function
    End If
    If Tipo = RELACIONES_GUILD.PAZ Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No estas en guerra con " & OtroClan
            Exit Function
        End If
    ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
    ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function
        End If
    End If
    Call guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True
End Function

Public Function r_VerPropuesta(ByVal Userindex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
    Dim OtroClanGI As Integer
    Dim GI         As Integer
    r_VerPropuesta = vbNullString
    refError = vbNullString
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningun clan."
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan."
        Exit Function
    End If
    OtroClanGI = GuildIndex(OtroGuild)
    If Not guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada."
        Exit Function
    End If
    r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, Tipo)
End Function

Public Function r_ListaDePropuestas(ByVal Userindex As Integer, ByVal Tipo As RELACIONES_GUILD) As String()
    Dim GI            As Integer
    Dim i             As Integer
    Dim proposalCount As Integer
    Dim proposals()   As String
    GI = UserList(Userindex).GuildIndex
    If GI > 0 And GI <= CANTIDADDECLANES Then
        With guilds(GI)
            proposalCount = .CantidadPropuestas(Tipo)
            If proposalCount > 0 Then
                ReDim proposals(proposalCount - 1) As String
            Else
                ReDim proposals(0) As String
            End If
            For i = 0 To proposalCount - 1
                proposals(i) = guilds(.Iterador_ProximaPropuesta(Tipo)).GuildName
            Next i
        End With
    End If
    r_ListaDePropuestas = proposals
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal Guild As Integer, ByRef Detalles As String)
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    Call SaveUserGuildRejectionReason(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    a_ObtenerRechazoDeChar = GetUserGuildRejectionReason(Aspirante)
    Call SaveUserGuildRejectionReason(Aspirante, vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal Userindex As Integer, ByRef Nombre As String, ByRef refError As String) As Boolean
    Dim GI           As Integer
    Dim NroAspirante As Integer
    a_RechazarAspirante = False
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningun clan"
        Exit Function
    End If
    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan."
        Exit Function
    End If
    Call guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
    a_RechazarAspirante = True
End Function

Public Function a_DetallesAspirante(ByVal Userindex As Integer, ByRef Nombre As String) As String
    Dim GI           As Integer
    Dim NroAspirante As Integer
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        Exit Function
    End If
    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
End Function

Public Sub SendDetallesPersonaje(ByVal Userindex As Integer, ByVal Personaje As String)
    On Error GoTo ErrorHandler
    Dim GI     As Integer
    Dim NroAsp As Integer
    Dim list() As String
    Dim i      As Long
    GI = UserList(Userindex).GuildIndex
    Personaje = UCase$(Personaje)
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Call Protocol.WriteConsoleMsg(Userindex, "No perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        Call Protocol.WriteConsoleMsg(Userindex, "No eres el lider de tu clan.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)
    End If
    NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
    If NroAsp = 0 Then
        list = guilds(GI).GetMemberList()
        For i = 0 To UBound(list())
            If Personaje = list(i) Then Exit For
        Next i
        If i > UBound(list()) Then
            Call Protocol.WriteConsoleMsg(Userindex, "El personaje no es ni aspirante ni miembro del clan.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If Not Database_Enabled Then
        Call SendCharacterInfoCharfile(Userindex, Personaje)
    Else
        Call SendCharacterInfoDatabase(Userindex, Personaje)
    End If
    Exit Sub
ErrorHandler:
    If Not PersonajeExiste(Personaje) Then
        Call LogError("El usuario " & UserList(Userindex).Name & " (" & Userindex & " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
    Else
        Call LogError("[" & Err.Number & "] " & Err.description & " En la rutina SendDetallesPersonaje, por el usuario " & UserList(Userindex).Name & " (" & Userindex & " ), pidiendo informacion sobre el personaje " & Personaje)
    End If
End Sub

Public Function a_NuevoAspirante(ByVal Userindex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
    Dim ViejoSolicitado   As String
    Dim ViejoGuildINdex   As Integer
    Dim ViejoNroAspirante As Integer
    Dim NuevoGuildIndex   As Integer
    a_NuevoAspirante = False
    If UserList(Userindex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro."
        Exit Function
    End If
    If EsNewbie(Userindex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function
    End If
    NuevoGuildIndex = GuildIndex(clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe, avise a un administrador."
        Exit Function
    End If
    If Not m_EstadoPermiteEntrar(Userindex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineacion " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If
    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contactate con un miembro para que procese las solicitudes."
        Exit Function
    End If
    ViejoSolicitado = GetUserGuildAspirant(UserList(Userindex).Name)
    If LenB(ViejoSolicitado) <> 0 Then
        ViejoGuildINdex = ViejoSolicitado
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(Userindex).Name)
            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(Userindex).Name, ViejoNroAspirante)
            End If
        Else
        End If
    End If
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(Userindex).Name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal Userindex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
    Dim GI           As Integer
    Dim NroAspirante As Integer
    Dim AspiranteUI  As Integer
    a_AceptarAspirante = False
    GI = UserList(Userindex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningun clan"
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(Userindex).Name, GI) Then
        refError = "No eres el lider de tu clan"
        Exit Function
    End If
    NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan."
        Exit Function
    End If
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan de alineacion " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = Aspirante & " no puede entrar a un clan de alineacion " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf GetUserGuildIndex(Aspirante) Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    End If
    Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call guilds(GI).AceptarNuevoMiembro(Aspirante)
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)
    End If
    a_AceptarAspirante = True
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildName = guilds(GuildIndex).GuildName
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildLeader = guilds(GuildIndex).GetLeader
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildFounder = guilds(GuildIndex).Fundador
End Function

Public Function GetUserGuildMember(ByVal UserName As String) As String
    If Not Database_Enabled Then
        GetUserGuildMember = GetUserGuildMemberCharfile(UserName)
    Else
        GetUserGuildMember = GetUserGuildMemberDatabase(UserName)
    End If
End Function

Public Function GetUserGuildAspirant(ByVal UserName As String) As Integer
    If Not Database_Enabled Then
        GetUserGuildAspirant = GetUserGuildAspirantCharfile(UserName)
    Else
        GetUserGuildAspirant = GetUserGuildAspirantDatabase(UserName)
    End If
End Function

Public Function GetUserGuildRejectionReason(ByVal UserName As String) As String
    If Not Database_Enabled Then
        GetUserGuildRejectionReason = GetUserGuildRejectionReasonCharfile(UserName)
    Else
        GetUserGuildRejectionReason = GetUserGuildRejectionReasonDatabase(UserName)
    End If
End Function

Public Function GetUserGuildPedidos(ByVal UserName As String) As String
    If Not Database_Enabled Then
        GetUserGuildPedidos = GetUserGuildPedidosCharfile(UserName)
    Else
        GetUserGuildPedidos = GetUserGuildPedidosDatabase(UserName)
    End If
End Function

Public Sub SaveUserGuildRejectionReason(ByVal UserName As String, ByVal Reason As String)
    If Not Database_Enabled Then
        Call SaveUserGuildRejectionReasonCharfile(UserName, Reason)
    Else
        Call SaveUserGuildRejectionReasonDatabase(UserName, Reason)
    End If
End Sub

Public Sub SaveUserGuildIndex(ByVal UserName As String, ByVal GuildIndex As Integer)
    If Not Database_Enabled Then
        Call SaveUserGuildIndexCharfile(UserName, GuildIndex)
    Else
        Call SaveUserGuildIndexDatabase(UserName, GuildIndex)
    End If
End Sub

Public Sub SaveUserGuildAspirant(ByVal UserName As String, ByVal AspirantIndex As Integer)
    If Not Database_Enabled Then
        Call SaveUserGuildAspirantCharfile(UserName, AspirantIndex)
    Else
        Call SaveUserGuildAspirantDatabase(UserName, AspirantIndex)
    End If
End Sub

Public Sub SaveUserGuildMember(ByVal UserName As String, ByVal guilds As String)
    If Not Database_Enabled Then
        Call SaveUserGuildMemberCharfile(UserName, guilds)
    Else
        Call SaveUserGuildMemberDatabase(UserName, guilds)
    End If
End Sub

Public Sub SaveUserGuildPedidos(ByVal UserName As String, ByVal Pedidos As String)
    If Not Database_Enabled Then
        Call SaveUserGuildPedidosCharfile(UserName, Pedidos)
    Else
        Call SaveUserGuildPedidosDatabase(UserName, Pedidos)
    End If
End Sub
