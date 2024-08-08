Attribute VB_Name = "General"
Option Explicit

#If False Then
    Dim X, Y, Map, K, errHandler, obj, index, n, Email As Variant
#End If

Global LeerNPCs As clsIniManager
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Function DarCuerpoDesnudo(ByVal Userindex As Integer, Optional ByVal Mimetizado As Boolean = False) As Integer
    With UserList(Userindex)
        Select Case .Genero
            Case eGenero.Hombre
                Select Case .raza
                    Case eRaza.Humano
                        DarCuerpoDesnudo = 21

                    Case eRaza.Drow
                        DarCuerpoDesnudo = 32

                    Case eRaza.Elfo
                        DarCuerpoDesnudo = 210

                    Case eRaza.Gnomo
                        DarCuerpoDesnudo = 222

                    Case eRaza.Enano
                        DarCuerpoDesnudo = 53
                End Select

            Case eGenero.Mujer
                Select Case .raza
                    Case eRaza.Humano
                        DarCuerpoDesnudo = 39

                    Case eRaza.Drow
                        DarCuerpoDesnudo = 40

                    Case eRaza.Elfo
                        DarCuerpoDesnudo = 259

                    Case eRaza.Gnomo
                        DarCuerpoDesnudo = 260

                    Case eRaza.Enano
                        DarCuerpoDesnudo = 60
                End Select
        End Select
        If Mimetizado Then
            .CharMimetizado.body = DarCuerpoDesnudo
        Else
            .Char.body = DarCuerpoDesnudo
        End If
        .flags.Desnudo = 1
    End With
End Function

Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)
    End If
End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        With MapData(Map, X, Y)
            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                (.Graphic(1) >= 12439 And .Graphic(1) <= 12454) Or _
                (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
                (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
                .Graphic(2) = 0 Then
                HayAgua = True
            Else
                HayAgua = False
            End If
        End With
    Else
        HayAgua = False
    End If
End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False
        End If
    Else
        HayLava = False
    End If
End Function

Function HaySacerdote(ByVal Userindex As Integer) As Boolean
    Dim X As Integer, Y As Integer
    With UserList(Userindex)
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then
                    If Npclist(MapData(.Pos.Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Then
                        If Distancia(.Pos, Npclist(MapData(.Pos.Map, X, Y).NpcIndex).Pos) < 5 Then
                            HaySacerdote = True
                            Exit Function
                        End If
                    End If
                End If
            Next X
        Next Y
    End With
    HaySacerdote = False
End Function

Sub EnviarSpawnList(ByVal Userindex As Integer)
    Dim K          As Long
    Dim npcNames() As String
    ReDim npcNames(1 To UBound(SpawnList)) As String
    For K = 1 To UBound(SpawnList)
        npcNames(K) = SpawnList(K).NpcName
    Next K
    Call WriteSpawnList(Userindex, npcNames())
End Sub

Public Function GetVersionOfTheServer() As String
    GetVersionOfTheServer = GetVar(App.Path & "\Server.ini", "INIT", "VersionTagRelease")
End Function

Sub Main()
    On Error Resume Next
    ChDir App.Path
    ChDrive App.Path
    Call LoadMotd
    Call BanIpCargar
    frmCargando.Show
    frmCargando.Label1(2).Caption = "Cargando constantes..."
    Call LoadConstants
    DoEvents
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call LoadArrays
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    Call LoadSini
    Call CargarCiudades
    Call CargaApuestas
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    Call LoadOBJData
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herreria"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpinteria"
    Call LoadObjCarpintero
    frmCargando.Label1(2).Caption = "Cargando Objetos del Artesano"
    Call LoadObjArtesano
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance
    frmCargando.Label1(2).Caption = "Cargando ArmadurasFaccionarias.dat"
    Call LoadArmadurasFaccion
    frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
    Call LoadPretorianData
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando Backup"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    Call InitializeAreas
    Call LoadArenas
    Call generateMatrix(MATRIX_INITIAL_MAP)
    Call ResetUsersConnections
    Call SocketConfig
    Call InitMainTimers
    Unload frmCargando
    LogServerStartTime
    If HideMe Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    tInicioServer = GetTickCount()
    IpPublicaServidor = frmMain.Inet1.OpenURL("http://ip1.dynupdate.no-ip.com")
    frmMain.lblIp.Tag = IpPublicaServidor & ":" & Puerto
    MundoSeleccionado = GetVar(App.Path & "\Dat\Map.dat", "INIT", "MapPath")
    NombreServidor = GetVar(App.Path & "\Server.ini", "INIT", "Nombre")
    frmMain.Caption = GetVersionOfTheServer() & " - Mundo Seleccionado: " & MundoSeleccionado & " - " & NombreServidor
    DescripcionServidor = GetVar(App.Path & "\Server.ini", "INIT", "Descripcion")
    frmMain.txtRecordOnline.Text = RecordUsuariosOnline
    If Not ClanPretoriano(ePretorianType.Default).SpawnClan(MAPA_PRETORIANO, PRETORIANO_X, PRETORIANO_Y, ePretorianType.Default) Then
        Call LogError("No se pudo invocar al Clan Pretoriano.")
    End If
    If ConexionAPI Then
        ApiNodeJsTaskId = Shell("cmd /c cd " & ApiPath & " && npm start")
    End If
End Sub

Private Sub LoadConstants()
    On Error Resume Next
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"
    CharPath = App.Path & "\Charfile\"
    AccountPath = App.Path & "\Account\"
    If LenB(Dir$(AccountPath, vbDirectory)) = 0 Then
        Call MkDir(AccountPath)
    End If
    If LenB(Dir$(CharPath, vbDirectory)) = 0 Then
        Call MkDir(CharPath)
    End If
    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasion en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apunalar) = "Apunalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    ListaPeces(1) = PECES_POSIBLES.PESCADO1
    ListaPeces(2) = PECES_POSIBLES.PESCADO2
    ListaPeces(3) = PECES_POSIBLES.PESCADO3
    ListaPeces(4) = PECES_POSIBLES.PESCADO4
    ListaPeces(5) = PECES_POSIBLES.PESCADO5
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    Set Ayuda = New cCola
    Set Denuncias = New cCola
    Denuncias.MaxLenght = MAX_DENOUNCES
    MaxUsers = 0
    Set WSAPISock2Usr = New Collection
    Protocol.InitAuxiliarBuffer
    Set aClon = New clsAntiMassClon
    Set TrashCollector = New Collection
End Sub

Private Sub LoadArrays()
    On Error Resume Next
    Call LoadRecords
    Call LoadGuildsDB
    Call LoadQuests
    Call CargarForbidenWords
End Sub

Private Sub ResetUsersConnections()
    On Error Resume Next
    Dim LoopC As Long
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
End Sub

Private Sub InitMainTimers()
    On Error Resume Next
    With frmMain
        .AutoSave.Enabled = True
        .GameTimer.Enabled = True
        .PacketResend.Enabled = True
        .TIMER_AI.Enabled = True
        .Auditoria.Enabled = True
        .TimerEnviarDatosServer.Enabled = True
    End With
    LastGameTick = GetTickCount
End Sub

Private Sub SocketConfig()
    On Error Resume Next
    Call SecurityIp.InitIpTables(1000)
    If LastSockListen >= 0 Then
        Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
    End If
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
    If SockListen <> -1 Then
        Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen)
    Else
        Call MsgBox("Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly)
    End If
    frmMain.txtStatus.Text = Date & " " & time & " - Escuchando conexiones entrantes ..."
End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
    FileExist = LenB(Dir$(File, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
    Dim i          As Long
    Dim lastPos    As Long
    Dim CurrentPos As Long
    Dim delimiter  As String * 1
    delimiter = Chr$(SepASCII)
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
    MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()
    frmMain.txtNumUsers.Text = NumUsers
End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
    Dim Arg As String
    Dim i   As Integer
    For i = 1 To 33
        Arg = ReadField(i, cad, 44)
        If LenB(Arg) = 0 Then Exit Function
    Next i
    ValidInputNP = True
End Function

Sub Restart()
    On Error Resume Next
    If frmMain.Visible Then frmMain.txtStatus.Text = "Reiniciando."
    Dim LoopC As Long
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next
    Call Statistics.Initialize
    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC
    ReDim UserList(1 To MaxUsers) As User
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    LastUser = 0
    NumUsers = 0
    Call FreeNPCs
    Call FreeCharIndexes
    Call LoadSini
    Call ResetForums
    Call LoadOBJData
    Call LoadMapData
    Call CargarHechizos
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " servidor reiniciado correctamente. - Escuchando conexiones entrantes ..."
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & time & " servidor reiniciado."
    Close #n
    If HideMe Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
End Sub

Public Function Intemperie(ByVal Userindex As Integer) As Boolean
    With UserList(Userindex)
        If MapInfo(.Pos.Map).Zona <> Dungeon Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.BAJOTECHO And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.CASA And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.ZONASEGURA Then _
                Intemperie = True
        Else
            Intemperie = False
        End If
    End With
    If IsArena(Userindex) Then Intemperie = False
End Function

Public Sub EfectoLluvia(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    If UserList(Userindex).flags.UserLogged Then
        If Intemperie(Userindex) Then
            Dim modifi As Long
            modifi = Porcentaje(UserList(Userindex).Stats.MaxSta, 3)
            Call QuitarSta(Userindex, modifi)
        End If
    End If
    Exit Sub
ErrorHandler:
    LogError ("Error en EfectoLluvia")
End Sub

Public Sub TiempoInvocacion(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        With UserList(Userindex)
            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia - DeltaTick
                    If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia <= 0 Then Call MuereNpc(.MascotasIndex(i), 0)
                End If
            End If
        End With
    Next i
End Sub

Public Sub EfectoFrio(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    Dim modifi As Integer
    With UserList(Userindex)
        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + DeltaTick
        Else
            If MapInfo(.Pos.Map).Terreno = Nieve Then
                Call WriteConsoleMsg(Userindex, "Estas muriendo de frio, abrigate o moriras!!", FontTypeNames.FONTTYPE_INFO)
                modifi = Porcentaje(.Stats.MaxHp, 5)
                .Stats.MinHp = .Stats.MinHp - modifi
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(Userindex, "Has muerto de frio!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(Userindex)
                End If
                Call WriteUpdateHP(Userindex)
            Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
                Call QuitarSta(Userindex, modifi)
                Call WriteUpdateSta(Userindex)
            End If
            .Counters.Frio = 0
        End If
    End With
End Sub

Public Sub EfectoLava(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    With UserList(Userindex)
        If .Counters.Lava < IntervaloFrio Then
            .Counters.Lava = .Counters.Lava + DeltaTick
        Else
            If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
                Call WriteConsoleMsg(Userindex, "Quitate de la lava, te estas quemando!!", FontTypeNames.FONTTYPE_INFO)
                .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(Userindex, "Has muerto quemado!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(Userindex)
                End If
                Call WriteUpdateHP(Userindex)
            End If
            .Counters.Lava = 0
        End If
    End With
End Sub

Public Sub EfectoEstadoAtacable(ByVal Userindex As Integer)
    If Not IntervaloEstadoAtacable(Userindex) Then
        UserList(Userindex).flags.AtacablePor = 0
        If Not UserList(Userindex).flags.Seguro Then
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOn)
        End If
        Call RefreshCharStatus(Userindex)
    End If
End Sub

Public Sub TravelingEffect(ByVal Userindex As Integer)
    If IntervaloGoHome(Userindex) Then
        Call HomeArrival(Userindex)
    End If
End Sub

Public Sub EfectoMimetismo(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    Dim Barco As ObjData
    With UserList(Userindex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + DeltaTick
        Else
            Call WriteConsoleMsg(Userindex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    Call ToggleBoatBody(Userindex)
                Else
                    .Char.body = iFragataFantasmal
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco
                End If
            Else
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            End If
            With .Char
                Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            .flags.Ignorado = False
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    With UserList(Userindex)
        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + DeltaTick
        Else
            .Counters.Invisibilidad = 0
            .flags.invisible = 0
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                If Not .flags.Navegando = 1 Then
                    Call SetInvisible(Userindex, .Char.CharIndex, False)
                End If
            End If
        End If
    End With
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
        End If
    End With
End Sub

Public Sub EfectoCegueEstu(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    With UserList(Userindex)
        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - DeltaTick
            If .Counters.Ceguera <= 0 Then
                If .flags.Estupidez = 1 Then
                    .flags.Estupidez = 0
                    Call WriteDumbNoMore(Userindex)
                Else
                    Call WriteBlindNoMore(Userindex)
                End If
            End If
        End If
    End With
End Sub

Public Sub EfectoParalisisUser(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    With UserList(Userindex)
        If .Counters.Paralisis > 0 Then
            Dim CasterIndex As Integer
            CasterIndex = .flags.ParalizedByIndex
            If .Stats.MaxMAN = 0 Then
                If CasterIndex <> 0 Then
                    If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
                        Call RemoveParalisis(Userindex)
                        Exit Sub
                    ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
                        Call RemoveParalisis(Userindex)
                        Exit Sub
                    ElseIf .Counters.Paralisis > IntervaloParalizadoReducido Then
                        If Not InVisionRangeAndMap(Userindex, UserList(CasterIndex).Pos) Then
                            .Counters.Paralisis = IntervaloParalizadoReducido
                            Exit Sub
                        End If
                    End If
                Else
                    CasterIndex = .flags.ParalizedByNpcIndex
                    If CasterIndex <> 0 Then
                        If .Counters.Paralisis > IntervaloParalizadoReducido Then
                            If Not InVisionRangeAndMap(Userindex, Npclist(CasterIndex).Pos) Then
                                .Counters.Paralisis = IntervaloParalizadoReducido
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
            .Counters.Paralisis = .Counters.Paralisis - DeltaTick
            If .Counters.Paralisis <= 0 Then
                Call RemoveParalisis(Userindex)
            End If
        End If
    End With
End Sub

Public Sub RemoveParalisis(ByVal Userindex As Integer)
    With UserList(Userindex)
        .flags.Paralizado = 0
        .flags.Inmovilizado = 0
        .flags.ParalizedBy = vbNullString
        .flags.ParalizedByIndex = 0
        .flags.ParalizedByNpcIndex = 0
        .Counters.Paralisis = 0
        Call WriteParalizeOK(Userindex)
    End With
End Sub

Public Sub RecStamina(ByVal Userindex As Integer, ByVal DeltaTick As Single, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
    With UserList(Userindex)
        Dim massta As Integer
        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + DeltaTick
            Else
                EnviarStats = True
                .Counters.STACounter = 0
                If .flags.Desnudo Then Exit Sub
                massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
                .Stats.MinSta = .Stats.MinSta + massta
                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta
                End If
            End If
        End If
    End With
End Sub

Public Sub EfectoVeneno(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    Dim n As Integer
    With UserList(Userindex)
        If .Counters.Veneno < IntervaloVeneno Then
            .Counters.Veneno = .Counters.Veneno + DeltaTick
        Else
            Call WriteConsoleMsg(Userindex, "Estas envenenado, si no te curas moriras.", FontTypeNames.FONTTYPE_VENENO)
            .Counters.Veneno = 0
            n = RandomNumber(1, 5)
            .Stats.MinHp = .Stats.MinHp - n
            If .Stats.MinHp < 1 Then Call UserDie(Userindex)
            Call WriteUpdateHP(Userindex)
        End If
    End With
End Sub

Public Sub DuracionPociones(ByVal Userindex As Integer, ByVal DeltaTick As Single)
    With UserList(Userindex)
        If .flags.DuracionEfecto > 0 Then
            .flags.DuracionEfecto = .flags.DuracionEfecto - DeltaTick
            If .flags.DuracionEfecto <= 0 Then
                .flags.TomoPocion = False
                Dim loopX As Integer
                For loopX = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(loopX) = .Stats.UserAtributosBackUP(loopX)
                Next loopX
                Call WriteUpdateStrenghtAndDexterity(Userindex)
            End If
        End If
    End With
End Sub

Public Sub HambreYSed(ByVal Userindex As Integer, ByVal DeltaTick As Single, ByRef fenviarAyS As Boolean)
    With UserList(Userindex)
        If Not .flags.Privilegios And PlayerType.User Then Exit Sub
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + DeltaTick
            Else
                .Counters.AGUACounter = 0
                .Stats.MinAGU = .Stats.MinAGU - 10
                If .Stats.MinAGU <= 0 Then
                    .Stats.MinAGU = 0
                    .flags.Sed = 1
                End If
                fenviarAyS = True
            End If
        End If
        If .Stats.MinHam > 0 Then
            If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + DeltaTick
            Else
                .Counters.COMCounter = 0
                .Stats.MinHam = .Stats.MinHam - 10
                If .Stats.MinHam <= 0 Then
                    .Stats.MinHam = 0
                    .flags.Hambre = 1
                End If
                fenviarAyS = True
            End If
        End If
    End With
End Sub

Public Sub Sanar(ByVal Userindex As Integer, ByVal DeltaTick As Single, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
    With UserList(Userindex)
        Dim mashit As Integer
        If .Stats.MinHp < .Stats.MaxHp Then
            If .Counters.HPCounter < Intervalo Then
                .Counters.HPCounter = .Counters.HPCounter + DeltaTick
            Else
                mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
                .Counters.HPCounter = 0
                .Stats.MinHp = .Stats.MinHp + mashit
                If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                Call WriteConsoleMsg(Userindex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
                EnviarStats = True
            End If
        End If
    End With
End Sub

Public Sub CargaNpcsDat(Optional ByVal ForzarActualizacionNpcsExistentes As Boolean = False)
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando NPCs.dat."
    Set LeerNPCs = New clsIniManager
    Call LeerNPCs.Initialize(DatPath & "NPCs.dat")
    Call CargarSpawnList
    If ForzarActualizacionNpcsExistentes Then
        Dim i As Long
        For i = 1 To MAXNPCS
            If Npclist(i).flags.NPCActive Then
                Call ReloadNPCByIndex(i)
            End If
            DoEvents
        Next i
    End If
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo NPCs.dat."
End Sub

Public Function ReiniciarAutoUpdate() As Double
    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)
End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    Call ES.DoBackUp
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")
    Unload frmMain
End Sub
 
Sub GuardarUsuarios()
    haciendoBK = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, False)
        End If
    Next i
    Call SaveRecords
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False
End Sub

Sub SaveUser(ByVal Userindex As Integer, Optional ByVal SaveTimeOnline As Boolean = True)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        If .Clase = 0 Or .Stats.ELV = 0 Then
            Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
            Exit Sub
        End If
        If .flags.Mimetizado = 1 Then
            .Char.body = .CharMimetizado.body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            .flags.Ignorado = False
        End If
        Dim Prom As Long
        Prom = (-.Reputacion.AsesinoRep) + (-.Reputacion.BandidoRep) + .Reputacion.BurguesRep + (-.Reputacion.LadronesRep) + .Reputacion.NobleRep + .Reputacion.PlebeRep
        Prom = Prom / 6
        .Reputacion.Promedio = Prom
        If Not Database_Enabled Then
            Call SaveUserToCharfile(Userindex, SaveTimeOnline)
        Else
            Call SaveUserToDatabase(Userindex, SaveTimeOnline)
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en SaveUser - Userindex: " & Userindex)
End Sub

Sub LoadUser(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    If Not Database_Enabled Then
        Call LoadUserFromCharfile(Userindex)
    Else
        Call LoadUserFromDatabase(Userindex)
    End If
    With UserList(Userindex)
        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = IntervaloParalizado
        End If
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
        End If
        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1
        End If
        If .Invent.EscudoEqpSlot > 0 Then
            .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
        End If
        If .Invent.CascoEqpSlot > 0 Then
            .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
        End If
        If .Invent.BarcoSlot > 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
        End If
        If .Invent.MunicionEqpSlot > 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
        End If
        If .Invent.AnilloEqpSlot > 0 Then
            .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex
        End If
        If .Invent.MonturaObjIndex > 0 Then
            .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaObjIndex).ObjIndex
        End If
        If .flags.Muerto = 0 Then
            .Char = .OrigChar
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
            .Char.heading = eHeading.SOUTH
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en LoadUser: " & UserList(Userindex).Name & " - " & Err.Number & " - " & Err.description)
End Sub

Public Sub FreeNPCs()
    Dim LoopC As Long
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub

Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal index As Integer, ByVal SoundIndex As Integer)
    Call SendData(Destino, index, PrepareMessagePlayWave(SoundIndex, UserList(index).Pos.X, UserList(index).Pos.Y))
End Sub


Public Function Tilde(ByRef data As String) As String
    Dim temp As String
    temp = UCase$(data)
    If InStr(1, temp, "Á") Then temp = Replace$(temp, "Á", "A")
    If InStr(1, temp, "e") Then temp = Replace$(temp, "e", "E")
    If InStr(1, temp, "Í") Then temp = Replace$(temp, "Í", "I")
    If InStr(1, temp, "Ó") Then temp = Replace$(temp, "Ó", "O")
    If InStr(1, temp, "U") Then temp = Replace$(temp, "U", "U")
    Tilde = temp
End Function

Public Sub CloseServer()
    If ConexionAPI Then
        Shell ("taskkill /PID " & ApiNodeJsTaskId)
    End If
    End
End Sub

Public Function GetTickCount() As Long
    GetTickCount = timeGetTime And &H7FFFFFFF
End Function
