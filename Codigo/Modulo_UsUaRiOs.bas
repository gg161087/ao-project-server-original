Attribute VB_Name = "UsUaRiOs"
Option Explicit

Public Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
    Dim DaExp       As Integer
    Dim EraCriminal As Boolean
    DaExp = CInt(UserList(VictimIndex).Stats.ELV) * 2
    With UserList(AttackerIndex)
        .Stats.Exp = .Stats.Exp + DaExp
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
            If UserList(VictimIndex).flags.AtacablePor <> AttackerIndex Then
                EraCriminal = criminal(AttackerIndex)
                With .Reputacion
                    If Not criminal(VictimIndex) Then
                        .AsesinoRep = .AsesinoRep + vlASESINO * 2
                        If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                        .BurguesRep = 0
                        .NobleRep = 0
                        .PlebeRep = 0
                    Else
                        .NobleRep = .NobleRep + vlNoble
                        If .NobleRep > MAXREP Then .NobleRep = MAXREP
                    End If
                End With
                Dim EsCriminal As Boolean
                EsCriminal = criminal(AttackerIndex)
                If EraCriminal <> EsCriminal Then
                    Call RefreshCharStatus(AttackerIndex)
                End If
            End If
        End If
        Call WriteMultiMessage(AttackerIndex, eMessages.HaveKilledUser, VictimIndex, DaExp)
        Call WriteMultiMessage(VictimIndex, eMessages.UserKill, AttackerIndex)
        Call LogAsesinato(.Name & " asesino a " & UserList(VictimIndex).Name)
    End With
End Sub

Public Sub RevivirUsuario(ByVal Userindex As Integer)
    With UserList(Userindex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion)
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        If .flags.Navegando = 1 Then
            Call ToggleBoatBody(Userindex)
        Else
            Call DarCuerpoDesnudo(Userindex)
            .Char.Head = .OrigChar.Head
        End If
        If .flags.Traveling Then
            .flags.Traveling = 0
            .Counters.goHome = 0
            Call WriteMultiMessage(Userindex, eMessages.CancelHome)
        End If
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(Userindex)
    End With
End Sub

Public Sub ToggleBoatBody(ByVal Userindex As Integer)
    Dim Ropaje        As Integer
    Dim EsFaccionario As Boolean
    Dim NewBody       As Integer
    With UserList(Userindex)
        .Char.Head = 0
        If .Invent.BarcoObjIndex = 0 Then Exit Sub
        Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
        If criminal(Userindex) Then
            EsFaccionario = esCaos(Userindex)
            Select Case Ropaje
                Case iBarca
                    If EsFaccionario Then
                        NewBody = iBarcaCaos
                    Else
                        NewBody = iBarcaPk
                    End If
                
                Case iGalera
                    If EsFaccionario Then
                        NewBody = iGaleraCaos
                    Else
                        NewBody = iGaleraPk
                    End If
                    
                Case iGaleon
                    If EsFaccionario Then
                        NewBody = iGaleonCaos
                    Else
                        NewBody = iGaleonPk
                    End If

                Case iFragataFantasmal
                    NewBody = iFragataFantasmal
            End Select
        Else
            EsFaccionario = esArmada(Userindex)
            If .flags.AtacablePor <> 0 Then
                Select Case Ropaje
                    Case iBarca
                        If EsFaccionario Then
                            NewBody = iBarcaRealAtacable
                        Else
                            NewBody = iBarcaCiudaAtacable
                        End If
                    
                    Case iGalera
                        If EsFaccionario Then
                            NewBody = iGaleraRealAtacable
                        Else
                            NewBody = iGaleraCiudaAtacable
                        End If
                        
                    Case iGaleon
                        If EsFaccionario Then
                            NewBody = iGaleonRealAtacable
                        Else
                            NewBody = iGaleonCiudaAtacable
                        End If

                    Case iFragataFantasmal
                        NewBody = iFragataFantasmal
                End Select
            Else
                Select Case Ropaje
                    Case iBarca
                        If EsFaccionario Then
                            NewBody = iBarcaReal
                        Else
                            NewBody = iBarcaCiuda
                        End If
                    
                    Case iGalera
                        If EsFaccionario Then
                            NewBody = iGaleraReal
                        Else
                            NewBody = iGaleraCiuda
                        End If
                        
                    Case iGaleon
                        If EsFaccionario Then
                            NewBody = iGaleonReal
                        Else
                            NewBody = iGaleonCiuda
                        End If

                    Case iFragataFantasmal
                        NewBody = iFragataFantasmal
                End Select
            End If
        End If
        .Char.body = NewBody
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
    End With
End Sub

Public Sub ToggleMonturaBody(ByVal Userindex As Integer)
    With UserList(Userindex)
        If .Invent.MonturaObjIndex = 0 Then Exit Sub
        .Char.body = ObjData(.Invent.MonturaObjIndex).Ropaje
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
    End With
End Sub

Public Sub ChangeUserChar(ByVal Userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)
    With UserList(Userindex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
        Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageHeadingChange(heading, .CharIndex))
    End With
End Sub

Public Function GetWeaponAnim(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Integer
    Dim Tmp As Integer
    With UserList(Userindex)
        Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
        If Tmp > 0 Then
            If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
                GetWeaponAnim = Tmp
                Exit Function
            End If
        End If
        GetWeaponAnim = ObjData(ObjIndex).WeaponAnim
    End With
End Function

Public Sub EnviarFama(ByVal Userindex As Integer)
    Dim L As Long
    With UserList(Userindex).Reputacion
        L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
        L = Round(L / 6)
        .Promedio = L
    End With
    Call WriteFame(Userindex)
End Sub

Public Sub EraseUserChar(ByVal Userindex As Integer, ByVal IsAdminInvisible As Boolean)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        CharList(.Char.CharIndex) = 0
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If
        If IsAdminInvisible Then
            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
        End If
        Call QuitarUser(Userindex, .Pos.Map)
        MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex = 0
        .Char.CharIndex = 0
    End With
    NumChars = NumChars - 1
    Exit Sub
ErrorHandler:
    Dim UserName  As String
    Dim CharIndex As Integer
    If Userindex > 0 Then
        UserName = UserList(Userindex).Name
        CharIndex = UserList(Userindex).Char.CharIndex
    End If
    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description & ". User: " & UserName & "(UI: " & Userindex & " - CI: " & CharIndex & ")")
End Sub

Public Sub RefreshCharStatus(ByVal Userindex As Integer)
    Dim ClanTag   As String
    Dim NickColor As Byte
    Dim NuevaA    As Boolean
    Dim GI        As Integer
    Dim tStr      As String
    With UserList(Userindex)
        If .GuildIndex > 0 Then
            ClanTag = modGuilds.GuildName(.GuildIndex)
            ClanTag = " <" & ClanTag & ">"
        End If
        NickColor = GetNickColor(Userindex)
        If .showName Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageUpdateTagAndStatus(Userindex, NickColor, .Name & ClanTag))
        Else
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageUpdateTagAndStatus(Userindex, NickColor, vbNullString))
        End If
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = iFragataFantasmal
            Else
                Call ToggleBoatBody(Userindex)
            End If
            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
        GI = .GuildIndex
        If GI > 0 Then
            NuevaA = False
            If Not modGuilds.m_ValidarPermanencia(Userindex, True, NuevaA) Then
                Call WriteConsoleMsg(Userindex, "Has sido expulsado del clan. El clan ha sumado un punto de antifaccion!", FontTypeNames.FONTTYPE_GUILD)
            End If
            If NuevaA Then
                Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("El clan ha pasado a tener alineacion " & modGuilds.GuildAlignment(GI) & "!", FontTypeNames.FONTTYPE_GUILD))
                tStr = modGuilds.GuildName(GI)
                Call LogClanes("El clan " & tStr & " cambio de alineacion!")
            End If
        End If
    End With
End Sub

Public Function GetNickColor(ByVal Userindex As Integer) As Byte
    With UserList(Userindex)
        If criminal(Userindex) Then
            GetNickColor = eNickColor.ieCriminal
        Else
            GetNickColor = eNickColor.ieCiudadano
        End If
        If .flags.AtacablePor > 0 Then GetNickColor = GetNickColor Or eNickColor.ieAtacable
    End With
End Function

Public Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ButIndex As Boolean = False)
    On Error GoTo ErrorHandler
    Dim CharIndex  As Integer
    Dim ClanTag    As String
    Dim NickColor  As Byte
    Dim UserName   As String
    Dim Privileges As Byte
    With UserList(Userindex)
        If InMapBounds(Map, X, Y) Then
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = Userindex
            End If
            If toMap Then MapData(Map, X, Y).Userindex = Userindex
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = modGuilds.GuildName(.GuildIndex)
                End If
                NickColor = GetNickColor(Userindex)
                Privileges = .flags.Privilegios
                If .showName Then
                    UserName = .Name
                    If .flags.EnConsulta Then
                        UserName = UserName & " " & TAG_CONSULT_MODE
                    Else
                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                            If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"
                        Else
                            If (.flags.invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) And .flags.Navegando = 0 Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else
                                If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"
                            End If
                        End If
                    End If
                End If
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, UserName, NickColor, Privileges)
            Else
                Call AgregarUser(Userindex, .Pos.Map, ButIndex)
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    Call CloseSocket(Userindex)
End Sub

Public Sub CheckUserLevel(ByVal Userindex As Integer, Optional ByVal PrintInConsole As Boolean = True)
    On Error GoTo ErrorHandler
    Dim Pts              As Integer
    Dim AumentoHIT       As Integer
    Dim AumentoMANA      As Integer
    Dim AumentoSTA       As Integer
    Dim AumentoHP        As Integer
    Dim WasNewbie        As Boolean
    Dim Promedio         As Double
    Dim aux              As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI               As Integer
    WasNewbie = EsNewbie(Userindex)
    With UserList(Userindex)
        Do While .Stats.Exp >= .Stats.ELU
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Sub
            End If
            Call Statistics.UserLevelUp(Userindex)
            If PrintInConsole Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
                Call WriteConsoleMsg(Userindex, "Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
            End If
            If .Stats.ELV = 1 Then
                Pts = 10
            Else
                Pts = Pts + 5
            End If
            .Stats.ELV = .Stats.ELV + 1
            .Stats.Exp = .Stats.Exp - .Stats.ELU
            If .Stats.ELV < 15 Then
                .Stats.ELU = .Stats.ELU * 1.4
            ElseIf .Stats.ELV < 21 Then
                .Stats.ELU = .Stats.ELU * 1.35
            ElseIf .Stats.ELV < 26 Then
                .Stats.ELU = .Stats.ELU * 1.3
            ElseIf .Stats.ELV < 35 Then
                .Stats.ELU = .Stats.ELU * 1.2
            ElseIf .Stats.ELV < 40 Then
                .Stats.ELU = .Stats.ELU * 1.3
            Else
                .Stats.ELU = .Stats.ELU * 1.375
            End If
            Promedio = ModVida(.Clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
            aux = RandomNumber(0, 100)
            If Promedio - Int(Promedio) = 0.5 Then
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 1.5
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio - 0.5
                Else
                    AumentoHP = Promedio - 1.5
                End If
            Else
                DistVida(1) = DistribucionEnteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
                DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 2
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 1
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio
                ElseIf aux <= DistVida(4) Then
                    AumentoHP = Promedio - 1
                Else
                    AumentoHP = Promedio - 2
                End If
            End If
        
            Select Case .Clase
                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Pirat
                    AumentoHIT = 3
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Thief
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mage
                    AumentoHIT = 1
                    AumentoMANA = 2.8 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                
                Case eClass.Worker
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTTrabajador
                
                Case eClass.Cleric
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druid
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bard
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                    
                Case eClass.Bandit
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3 * 2
                    AumentoSTA = AumentoStBandido
                
                Case Else
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
            End Select
            .Stats.MaxHp = .Stats.MaxHp + AumentoHP
            If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then .Stats.MaxHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then .Stats.MaxHIT = STAT_MAXHIT_OVER36
            End If
            .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then .Stats.MinHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then .Stats.MinHIT = STAT_MAXHIT_OVER36
            End If
            If PrintInConsole Then
                If AumentoHP > 0 Then
                    Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
                End If
                If AumentoSTA > 0 Then
                    Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoSTA & " puntos de energia.", FontTypeNames.FONTTYPE_INFO)
                End If
                If AumentoMANA > 0 Then
                    Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoMANA & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)
                End If
                If AumentoHIT > 0 Then
                    Call WriteConsoleMsg(Userindex, "Tu golpe maximo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(Userindex, "Tu golpe minimo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
            .Stats.MinHp = .Stats.MaxHp
            Call mdParty.ActualizarSumaNivelesElevados(Userindex)
            If .Stats.ELV = 25 Then
                GI = .GuildIndex
                If GI > 0 Then
                    If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                        Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
                        If PrintInConsole Then
                            Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                            Call WriteConsoleMsg(Userindex, "Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearas! Por esta razon, hasta tanto no te enlistes en la faccion bajo la cual tu clan esta alineado, estaras excluido del mismo.", FontTypeNames.FONTTYPE_GUILD)
                        End If
                    End If
                End If
            End If
        Loop
        If Not EsNewbie(Userindex) And WasNewbie Then
            Call QuitarNewbieObj(Userindex)
            If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
                Call WarpUserChar(Userindex, 1, 50, 50, True)
                If PrintInConsole Then
                    Call WriteConsoleMsg(Userindex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        If Pts > 0 Then
            Call WriteLevelUp(Userindex, Pts)
            .Stats.SkillPts = .Stats.SkillPts + Pts
            If PrintInConsole Then
                Call WriteConsoleMsg(Userindex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Call SaveUser(Userindex, True)
    Call WriteUpdateUserStats(Userindex)
    Exit Sub
ErrorHandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)
End Sub

Public Function PuedeAtravesarAgua(ByVal Userindex As Integer) As Boolean
    PuedeAtravesarAgua = UserList(Userindex).flags.Navegando = 1 Or UserList(Userindex).flags.Vuela = 1
End Function

Sub MoveUserChar(ByVal Userindex As Integer, ByVal nHeading As eHeading)
    Dim nPos          As WorldPos
    Dim sailing       As Boolean
    Dim CasperIndex   As Integer
    Dim CasperHeading As eHeading
    Dim isAdminInvi   As Boolean
    sailing = PuedeAtravesarAgua(Userindex)
    nPos = UserList(Userindex).Pos
    Call HeadtoPos(nHeading, nPos)
    isAdminInvi = (UserList(Userindex).flags.AdminInvisible = 1)
    If MoveToLegalPos(UserList(Userindex).Pos.Map, nPos.X, nPos.Y, sailing, Not sailing) Then
        If UserList(Userindex).flags.Equitando And _
           (MapData(UserList(Userindex).Pos.Map, nPos.X, nPos.Y).trigger = eTrigger.CASA Or _
           MapData(UserList(Userindex).Pos.Map, nPos.X, nPos.Y).trigger = eTrigger.BAJOTECHO Or _
           MapInfo(UserList(Userindex).Pos.Map).Zona = Dungeon) Then _

            Exit Sub
        End If
        If MapInfo(UserList(Userindex).Pos.Map).NumUsers > 1 Then
            CasperIndex = MapData(UserList(Userindex).Pos.Map, nPos.X, nPos.Y).Userindex
            If CasperIndex > 0 Then
                If Not isAdminInvi Then
                    If TriggerZonaPelea(Userindex, CasperIndex) = TRIGGER6_PROHIBE Then
                        If UserList(CasperIndex).flags.SeguroResu = False Then
                            UserList(CasperIndex).flags.SeguroResu = True
                            Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)
                        End If
                    End If
                    With UserList(CasperIndex)
                        CasperHeading = InvertHeading(nHeading)
                        Call HeadtoPos(CasperHeading, .Pos)
                        If Not .flags.AdminInvisible = 1 Then Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y))
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                        .Char.heading = CasperHeading
                        MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex = CasperIndex
                    End With
                    Call Areas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                End If
            End If
            If Not isAdminInvi Then Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageCharacterMove(UserList(Userindex).Char.CharIndex, nPos.X, nPos.Y))
        End If
        If Not (isAdminInvi And (CasperIndex <> 0)) Then
            Dim oldUserIndex As Integer
            With UserList(Userindex)
                oldUserIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex
                If oldUserIndex = Userindex Then
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex = 0
                End If
                .Pos = nPos
                .Char.heading = nHeading
                MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex = Userindex
                If HaySacerdote(Userindex) Then Call AccionParaSacerdote(Userindex)
                Call DoTileEvents(Userindex, .Pos.Map, .Pos.X, .Pos.Y)
            End With
            Call Areas.CheckUpdateNeededUser(Userindex, nHeading)
        Else
            Call WritePosUpdate(Userindex)
        End If
    Else
        Call WritePosUpdate(Userindex)
    End If
    If UserList(Userindex).Counters.Trabajando Then UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando - 1
    If UserList(Userindex).Counters.Ocultando Then UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando - 1
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST

        Case eHeading.WEST
            InvertHeading = EAST

        Case eHeading.SOUTH
            InvertHeading = NORTH

        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function

Sub ChangeUserInv(ByVal Userindex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
    UserList(Userindex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(Userindex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim LoopC As Long
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            If LoopC > LastChar Then LastChar = LoopC
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    NextOpenUser = LoopC
End Function

Public Sub LiberarSlot(ByVal Userindex As Integer)
    With UserList(Userindex)
        .ConnID = -1
        .ConnIDValida = False
    End With
    If Userindex = LastUser Then
        Do While (LastUser > 0) And (UserList(LastUser).ConnID = -1)
            LastUser = LastUser - 1
            If LastUser = 0 Then Exit Do
        Loop
    End If
End Sub

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    Dim GuildI As Integer
    With UserList(Userindex)
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Mana: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energia: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
        End If
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If

        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        GuildI = .GuildIndex
        If GuildI > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        #If ConUpTime Then
            Dim TempDate As Date
            Dim TempSecs As Long
            Dim TempStr  As String
            TempDate = Now - .LogOnTime
            TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
            TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
            Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If
        If .flags.Traveling = 1 Then
            Call WriteConsoleMsg(sendIndex, "Tiempo restante para llegar a tu hogar: " & GetHomeArrivalTime(Userindex) & " segundos.", FontTypeNames.FONTTYPE_INFO)
        End If
        Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.Gld & "  Posicion: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.Map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    With UserList(Userindex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.Clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & (.Counters.Pena / 40), FontTypeNames.FONTTYPE_INFO)
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejercito real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingreso en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legion oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingreso en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejercito real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legion oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        Call WriteConsoleMsg(sendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    On Error Resume Next
    Dim j As Long
    With UserList(Userindex)
        Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        For j = 1 To .CurrentInventorySlots
            If .Invent.Object(j).ObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    End With
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    On Error Resume Next
    Dim j As Integer
    Call WriteConsoleMsg(sendIndex, UserList(Userindex).Name, FontTypeNames.FONTTYPE_INFO)
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(Userindex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(Userindex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Boolean
    If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "" & UserList(Userindex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
    Dim EraCriminal As Boolean
    Npclist(NpcIndex).flags.AttackedBy = UserList(Userindex).Name
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(Userindex).flags.NPCAtacado
    UserList(Userindex).flags.NPCAtacado = NpcIndex
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(Userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(Userindex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(Userindex).Name Then
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(Userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> Userindex Then
            Call AllMascotasAtacanUser(Userindex, Npclist(NpcIndex).MaestroUser)
        End If
    End If
    If EsMascotaCiudadano(NpcIndex, Userindex) Then
        Call VolverCriminal(Userindex)
        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1
    Else
        EraCriminal = criminal(Userindex)
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                Call VolverCriminal(Userindex)
            End If
        ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
            UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlCAZADOR / 2
            If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then UserList(Userindex).Reputacion.PlebeRep = MAXREP
        End If
        If Npclist(NpcIndex).MaestroUser <> Userindex Then
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
        End If
        If EraCriminal And Not criminal(Userindex) Then
            Call VolverCiudadano(Userindex)
        End If
    End If
End Sub

Public Function PuedeApunalar(ByVal Userindex As Integer) As Boolean
    Dim WeaponIndex As Integer
    With UserList(Userindex)
        WeaponIndex = .Invent.WeaponEqpObjIndex
        If WeaponIndex > 0 Then
            If ObjData(WeaponIndex).Apunala = 1 Then
                PuedeApunalar = .Stats.UserSkills(eSkill.Apunalar) >= MIN_APUNALAR Or .Clase = eClass.Assasin
            End If
        End If
    End With
End Function

Public Function PuedeAcuchillar(ByVal Userindex As Integer) As Boolean
    Dim WeaponIndex As Integer
    With UserList(Userindex)
        If .Clase = eClass.Pirat Then
            WeaponIndex = .Invent.WeaponEqpObjIndex
            If WeaponIndex > 0 Then
                PuedeAcuchillar = (ObjData(WeaponIndex).Acuchilla = 1)
            End If
        End If
    End With
End Function

Sub SubirSkill(ByVal Userindex As Integer, ByVal Skill As Integer, ByVal Acerto As Boolean)
    With UserList(Userindex)
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            If .Counters.AsignedSkills < 10 Then
                If Not .flags.UltimoMensaje = 7 Then
                    Call WriteConsoleMsg(Userindex, "Para poder entrenar un skill debes asignar los 10 skills iniciales.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 7
                End If
                Exit Sub
            End If
            With .Stats
                If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
                Dim Lvl As Integer
                Lvl = .ELV
                If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
                If .UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
                If Acerto Then
                    .ExpSkills(Skill) = .ExpSkills(Skill) + EXP_ACIERTO_SKILL
                Else
                    .ExpSkills(Skill) = .ExpSkills(Skill) + EXP_FALLO_SKILL
                End If
                If .ExpSkills(Skill) >= .EluSkills(Skill) Then
                    .UserSkills(Skill) = .UserSkills(Skill) + 1
                    Call WriteConsoleMsg(Userindex, "Has mejorado tu skill " & SkillsNames(Skill) & " en un punto! Ahora tienes " & .UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                    .Exp = .Exp + 50
                    If .Exp > MAXEXP Then .Exp = MAXEXP
                    Call WriteConsoleMsg(Userindex, "Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteUpdateExp(Userindex)
                    Call CheckUserLevel(Userindex)
                    Call CheckEluSkill(Userindex, Skill, False)
                End If
            End With
        End If
    End With
End Sub

Public Sub UserDie(ByVal Userindex As Integer, Optional ByVal AttackerIndex As Integer = 0)
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim aN          As Integer
    Dim iSoundDeath As Integer
    With UserList(Userindex)
        If .Genero = eGenero.Mujer Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_MUJER_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_MUJER
            End If
        Else
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE
            End If
        End If
        Call ReproducirSonido(SendTarget.ToPCArea, Userindex, iSoundDeath)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1
        .Counters.Trabajando = 0
        If TriggerZonaPelea(Userindex, Userindex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOn)
        Else
            .flags.SeguroResu = False
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOff)
        End If
        aN = .flags.AtacadoPorNpc
        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString
        End If
        aN = .flags.NPCAtacado
        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString
            End If
        End If
        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        Call PerdioNpc(Userindex, False)
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)
        End If
        If .flags.AtacablePor > 0 Then
            .flags.AtacablePor = 0
            Call RefreshCharStatus(Userindex)
        End If
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(Userindex)
        End If
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(Userindex)
        End If
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(Userindex)
        End If
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(Userindex)
        End If
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            Call SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, False)
        End If
        If TriggerZonaPelea(Userindex, Userindex) <> eTrigger6.TRIGGER6_PERMITE Then
            If DropItemsAlMorir Then
                If MapInfo(.Pos.Map).Pk Then
                    If .Invent.MochilaEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.MochilaEqpSlot)
                    End If
                    If Not EsNewbie(Userindex) Then
                        Call TirarTodo(Userindex)
                    Else
                        Call TirarTodosLosItemsNoNewbies(Userindex)
                    End If
                End If
            End If
        End If
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.ArmourEqpSlot)
        End If
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
        End If
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.CascoEqpSlot)
        End If
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(Userindex, .Invent.AnilloEqpSlot)
        End If
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.MunicionEqpSlot)
        End If
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
        End If
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0
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
        If .flags.TomoPocion = True Then
            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
        End If
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            .Char.body = iFragataFantasmal
        End If
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                Call MuereNpc(.MascotasIndex(i), 0)
            Else
                .MascotasType(i) = 0
            End If
        Next i
        .NroMascotas = 0
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateUserStats(Userindex)
        Call WriteUpdateStrenghtAndDexterity(Userindex)
        If .PartyIndex > 0 Then
            Call mdParty.ObtenerExito(Userindex, .Stats.ELV * -10 * mdParty.CantMiembros(Userindex), .Pos.Map, .Pos.X, .Pos.Y)
        End If
        Call LimpiarComercioSeguro(Userindex)
        Dim Mapa As Integer
        Mapa = .Pos.Map
        Dim MapaTelep As Integer
        MapaTelep = MapInfo(Mapa).OnDeathGoTo.Map
        If MapaTelep <> 0 Then
            Call WriteConsoleMsg(Userindex, "Tu estado no te permite permanecer en el mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WarpUserChar(Userindex, MapaTelep, MapInfo(Mapa).OnDeathGoTo.X, MapInfo(Mapa).OnDeathGoTo.Y, True, True)
        End If
        If AttackerIndex <> 0 Then
            If .flags.SlotReto > 0 Then
                Call Retos.UserDieFight(Userindex, AttackerIndex, False)
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripcion: " & Err.description)
End Sub

Public Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
    If EsNewbie(Muerto) Then Exit Sub
    With UserList(Atacante)
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        If criminal(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name
                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
            End If
            If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
                .Faccion.Reenlistadas = 200
            End If
        Else
            If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                .flags.LastCiudMatado = UserList(Muerto).Name
                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
            End If
        End If
        If .Stats.UsuariosMatados < MAXUSERMATADOS Then .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1
    End With
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef obj As obj, ByRef PuedeAgua As Boolean, ByRef PuedeTierra As Boolean)
    On Error GoTo ErrorHandler
    Dim Found As Boolean
    Dim LoopC As Integer
    Dim tX    As Long
    Dim tY    As Long
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    LoopC = 1
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, True) Then
        If Not HayObjeto(Pos.Map, nPos.X, nPos.Y, obj.ObjIndex, obj.Amount) Then
            Found = True
        End If
    End If
    If Not Found Then
        While (Not Found) And LoopC <= 16
            If RhombLegalTilePos(Pos, tX, tY, LoopC, obj.ObjIndex, obj.Amount, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True
            End If
            LoopC = LoopC + 1
        Wend
    End If
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en Tilelibre. Error: " & Err.Number & " - " & Err.description)
End Sub

Sub WarpUserChar(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal FX As Boolean, Optional ByVal Teletransported As Boolean)
    Dim OldMap As Integer
    Dim OldX   As Integer
    Dim OldY   As Integer
    With UserList(Userindex)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        OldMap = .Pos.Map
        OldX = .Pos.X
        OldY = .Pos.Y
        Call EraseUserChar(Userindex, .flags.AdminInvisible = 1)
        If OldMap <> Map Then
            Call WriteChangeMap(Userindex, Map, MapInfo(.Pos.Map).MapVersion)
            If .flags.Privilegios And PlayerType.User Then
                Dim AhoraVisible As Boolean
                Dim WasInvi      As Boolean
                If MapInfo(Map).InviSinEfecto > 0 And .flags.invisible = 1 Then
                    .flags.invisible = 0
                    .Counters.Invisibilidad = 0
                    AhoraVisible = True
                    WasInvi = True
                End If
                If MapInfo(Map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
                    AhoraVisible = True
                    .flags.Oculto = 0
                    .Counters.TiempoOculto = 0
                End If
                If AhoraVisible Then
                    Call SetInvisible(Userindex, .Char.CharIndex, False)
                    If WasInvi Then
                        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
            If MapInfo(Map).MusicMp3 <> vbNullString Then
                Call WritePlayMp3(Userindex, MapInfo(Map).MusicMp3)
            Else
                Call WritePlayMidi(Userindex, val(ReadField(1, MapInfo(Map).Music, 45)))
            End If
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
            Dim nextMap, previousMap As Boolean
            nextMap = IIf(distanceToCities(Map).distanceToCity(.Hogar) >= 0, True, False)
            previousMap = IIf(distanceToCities(.Pos.Map).distanceToCity(.Hogar) >= 0, True, False)
            If previousMap And nextMap Then
            ElseIf previousMap And Not nextMap Then
                .flags.lastMap = .Pos.Map
            ElseIf Not previousMap And nextMap Then
                .flags.lastMap = 0
            ElseIf Not previousMap And Not nextMap Then
                .flags.lastMap = .flags.lastMap
            End If
            Call WriteRemoveAllDialogs(Userindex)
        End If
        .Pos.X = X
        .Pos.Y = Y
        .Pos.Map = Map
        Call MakeUserChar(True, Map, Userindex, Map, X, Y)
        Call WriteUserCharIndexInServer(Userindex)
        Call DoTileEvents(Userindex, Map, X, Y)
        If Teletransported Then
            If .flags.Traveling = 1 Then
                .flags.Traveling = 0
                .Counters.goHome = 0
                Call WriteMultiMessage(Userindex, eMessages.CancelHome)
            End If
        End If
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_WARP, X, Y))
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        End If
        If .NroMascotas Then Call WarpMascotas(Userindex)
        Call IntervaloPermiteSerAtacado(Userindex, True)
        Call PerdioNpc(Userindex, False)
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                If .flags.Navegando = 0 Then
                    .flags.Navegando = 1
                    Call WriteNavigateToggle(Userindex)
                End If
            Else
                If .flags.Navegando = 1 Then
                    .flags.Navegando = 0
                    Call WriteNavigateToggle(Userindex)
                End If
            End If
        End If
    End With
End Sub

Private Sub WarpMascotas(ByVal Userindex As Integer)
    Dim i                As Integer
    Dim petType          As Integer
    Dim PetRespawn       As Boolean
    Dim PetTiempoDeVida  As Integer
    Dim NroPets          As Integer
    Dim InvocadosMatados As Integer
    Dim canWarp          As Boolean
    Dim index            As Integer
    Dim iMinHP           As Integer
    NroPets = UserList(Userindex).NroMascotas
    canWarp = (MapInfo(UserList(Userindex).Pos.Map).Pk = True)
    For i = 1 To MAXMASCOTAS
        index = UserList(Userindex).MascotasIndex(i)
        If index > 0 Then
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(Userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                petType = 0
            Else
                petType = UserList(Userindex).MascotasType(i)
                iMinHP = Npclist(index).Stats.MinHp
                Call QuitarNPC(index)
                UserList(Userindex).MascotasType(i) = petType
            End If
        ElseIf UserList(Userindex).MascotasType(i) > 0 Then
            PetRespawn = True
            petType = UserList(Userindex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0
        End If
        If petType > 0 And canWarp Then
            Dim SpawnPos As WorldPos
            SpawnPos.Map = UserList(Userindex).Pos.Map
            SpawnPos.X = UserList(Userindex).Pos.X + RandomNumber(-3, 3)
            SpawnPos.Y = UserList(Userindex).Pos.Y + RandomNumber(-3, 3)
            index = SpawnNpc(petType, SpawnPos, False, PetRespawn)
            If index = 0 Then
                Call WriteConsoleMsg(Userindex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Else
                UserList(Userindex).MascotasIndex(i) = index
                Npclist(index).Stats.MinHp = IIf(iMinHP = 0, Npclist(index).Stats.MinHp, iMinHP)
                Npclist(index).MaestroUser = Userindex
                Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(index)
            End If
        End If
    Next i
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(Userindex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    If Not canWarp Then
        Call WriteConsoleMsg(Userindex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    UserList(Userindex).NroMascotas = NroPets
End Sub

Public Sub WarpMascota(ByVal Userindex As Integer, ByVal PetIndex As Integer)
    Dim petType   As Integer
    Dim NpcIndex  As Integer
    Dim iMinHP    As Integer
    Dim TargetPos As WorldPos
    With UserList(Userindex)
        TargetPos.Map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
        NpcIndex = .MascotasIndex(PetIndex)
        petType = .MascotasType(PetIndex)
        iMinHP = Npclist(NpcIndex).Stats.MinHp
        Call QuitarNPC(NpcIndex)
        .MascotasType(PetIndex) = petType
        .NroMascotas = .NroMascotas + 1
        NpcIndex = SpawnNpc(petType, TargetPos, False, False)
        If NpcIndex = 0 Then
            Call WriteConsoleMsg(Userindex, "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
        Else
            .MascotasIndex(PetIndex) = NpcIndex
            With Npclist(NpcIndex)
                .Stats.MinHp = IIf(iMinHP = 0, .Stats.MinHp, iMinHP)
                .MaestroUser = Userindex
                .Movement = TipoAI.SigueAmo
                .Target = 0
                .TargetNPC = 0
            End With
            Call FollowAmo(NpcIndex)
        End If
    End With
End Sub

Sub Cerrar_Usuario(ByVal Userindex As Integer)
    Dim isNotVisible As Boolean
    Dim HiddenPirat  As Boolean
    With UserList(Userindex)
        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.Map).Pk, IntervaloCerrarConexion, 0)
            isNotVisible = (.flags.Oculto Or .flags.invisible)
            If isNotVisible Then
                .flags.invisible = 0
                If .flags.Oculto Then
                    If .flags.Navegando = 1 Then
                        If .Clase = eClass.Pirat Then
                            ' Pierde la apariencia de fragata fantasmal
                            Call ToggleBoatBody(Userindex)
                            Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                            HiddenPirat = True
                        End If
                    End If
                End If
                .flags.Oculto = 0
                If Not HiddenPirat Then Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                If .flags.Navegando = 0 Then
                    Call SetInvisible(Userindex, .Char.CharIndex, False)
                End If
            End If
            If .flags.Traveling = 1 Then
                Call WriteMultiMessage(Userindex, eMessages.CancelHome)
                .flags.Traveling = 0
                .Counters.goHome = 0
            End If
            Call WriteConsoleMsg(Userindex, "Cerrando...Se cerrara el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub CancelExit(ByVal Userindex As Integer)
    If UserList(Userindex).Counters.Saliendo Then
        If UserList(Userindex).ConnIDValida Then
            UserList(Userindex).Counters.Saliendo = False
            UserList(Userindex).Counters.Salir = 0
            Call WriteConsoleMsg(Userindex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            UserList(Userindex).Counters.Salir = IIf((UserList(Userindex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(Userindex).Pos.Map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

Sub VolverCriminal(ByVal Userindex As Integer)
    With UserList(Userindex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            .Reputacion.BurguesRep = 0
            .Reputacion.NobleRep = 0
            .Reputacion.PlebeRep = 0
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO
            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(Userindex)
            If .flags.AtacablePor > 0 Then .flags.AtacablePor = 0
        End If
    End With
    Call RefreshCharStatus(Userindex)
End Sub

Sub VolverCiudadano(ByVal Userindex As Integer)
    With UserList(Userindex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        .Reputacion.LadronesRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO
        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
    End With
    Call RefreshCharStatus(Userindex)
End Sub

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function

Public Sub SetInvisible(ByVal Userindex As Integer, ByVal userCharIndex As Integer, ByVal invisible As Boolean)
    Dim sndNick As String
    With UserList(Userindex)
        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, Userindex, PrepareMessageSetInvisible(userCharIndex, invisible))
        sndNick = .Name
        If invisible Then
            sndNick = sndNick & " " & TAG_USER_INVISIBLE
        Else
            If .GuildIndex > 0 Then
                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
            End If
        End If
        Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, Userindex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))
    End With
End Sub

Public Sub SetConsulatMode(ByVal Userindex As Integer)
    Dim sndNick As String
    With UserList(Userindex)
        sndNick = .Name
        If .flags.EnConsulta Then
            sndNick = sndNick & " " & TAG_CONSULT_MODE
        Else
            If .GuildIndex > 0 Then
                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
            End If
        End If
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))
    End With
End Sub

Public Function IsArena(ByVal Userindex As Integer) As Boolean
    IsArena = (TriggerZonaPelea(Userindex, Userindex) = TRIGGER6_PERMITE)
End Function

Public Sub PerdioNpc(ByVal Userindex As Integer, Optional ByVal CheckPets As Boolean = True)
    Dim PetCounter As Long
    Dim PetIndex   As Integer
    Dim NpcIndex   As Integer
    With UserList(Userindex)
        NpcIndex = .flags.OwnedNpc
        If NpcIndex > 0 Then
            If CheckPets Then
                If .NroMascotas > 0 Then
                    For PetCounter = 1 To MAXMASCOTAS
                        PetIndex = .MascotasIndex(PetCounter)
                        If PetIndex > 0 Then
                            If Npclist(PetIndex).TargetNPC = NpcIndex Then
                                Call FollowAmo(PetIndex)
                            End If
                        End If
                    Next PetCounter
                End If
            End If
            Npclist(NpcIndex).Owner = 0
            .flags.OwnedNpc = 0
        End If
    End With
End Sub

Public Sub ApropioNpc(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
    With UserList(Userindex)
        If EsGm(Userindex) Then Exit Sub
        Dim Mapa As Integer
        Mapa = .Pos.Map
        If MapData(Mapa, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then Exit Sub
        If MapInfo(Mapa).Pk = False Then Exit Sub
        If MapInfo(Mapa).RoboNpcsPermitido = 1 Then Exit Sub
        If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        Npclist(NpcIndex).Owner = Userindex
        .flags.OwnedNpc = NpcIndex
    End With
    Call IntervaloPerdioNpc(Userindex, True)
End Sub

Public Function GetDireccion(ByVal Userindex As Integer, ByVal OtherUserIndex As Integer) As String
    Dim X As Integer
    Dim Y As Integer
    X = UserList(Userindex).Pos.X - UserList(OtherUserIndex).Pos.X
    Y = UserList(Userindex).Pos.Y - UserList(OtherUserIndex).Pos.Y
    If X = 0 And Y > 0 Then
        GetDireccion = "Sur"
    ElseIf X = 0 And Y < 0 Then
        GetDireccion = "Norte"
    ElseIf X > 0 And Y = 0 Then
        GetDireccion = "Este"
    ElseIf X < 0 And Y = 0 Then
        GetDireccion = "Oeste"
    ElseIf X > 0 And Y < 0 Then
        GetDireccion = "NorEste"
    ElseIf X < 0 And Y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf X > 0 And Y > 0 Then
        GetDireccion = "SurEste"
    ElseIf X < 0 And Y > 0 Then
        GetDireccion = "SurOeste"
    End If
End Function

Public Function SameFaccion(ByVal Userindex As Integer, ByVal OtherUserIndex As Integer) As Boolean
    SameFaccion = (esCaos(Userindex) And esCaos(OtherUserIndex)) Or (esArmada(Userindex) And esArmada(OtherUserIndex))
End Function

Public Function FarthestPet(ByVal Userindex As Integer) As Integer
    On Error GoTo ErrorHandler
    Dim PetIndex      As Integer
    Dim Distancia     As Integer
    Dim OtraDistancia As Integer
    With UserList(Userindex)
        If .NroMascotas = 0 Then Exit Function
        For PetIndex = 1 To MAXMASCOTAS
            If .MascotasIndex(PetIndex) > 0 Then
                If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                    Else
                        OtraDistancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex
                        End If
                    End If
                End If
            End If
        Next PetIndex
    End With
    Exit Function
ErrorHandler:
    Call LogError("Error en FarthestPet")
End Function

Public Sub CheckEluSkill(ByVal Userindex As Integer, ByVal Skill As Byte, ByVal Allocation As Boolean)
    With UserList(Userindex).Stats
        If .UserSkills(Skill) < MAXSKILLPOINTS Then
            If Allocation Then
                .ExpSkills(Skill) = 0
            Else
                .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)
            End If
            .EluSkills(Skill) = ELU_SKILL_INICIAL * 1.05 ^ .UserSkills(Skill)
        Else
            .ExpSkills(Skill) = 0
            .EluSkills(Skill) = 0
        End If
    End With
End Sub

Public Function HasEnoughItems(ByVal Userindex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long) As Boolean
    Dim Slot          As Long
    Dim ItemInvAmount As Long
    With UserList(Userindex)
        For Slot = 1 To .CurrentInventorySlots
            If .Invent.Object(Slot).ObjIndex = ObjIndex Then
                ItemInvAmount = ItemInvAmount + .Invent.Object(Slot).Amount
            End If
        Next Slot
    End With
    HasEnoughItems = Amount <= ItemInvAmount
End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, ByVal Userindex As Integer) As Long
    Dim Slot As Byte
    For Slot = 1 To MAX_OFFER_SLOTS
        If UserList(Userindex).ComUsu.Objeto(Slot) = ObjIndex Then
            TotalOfferItems = TotalOfferItems + UserList(Userindex).ComUsu.cant(Slot)
        End If
    Next Slot
End Function

Public Function getMaxInventorySlots(ByVal Userindex As Integer) As Byte
    If UserList(Userindex).Invent.MochilaEqpObjIndex > 0 Then
        getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(UserList(Userindex).Invent.MochilaEqpObjIndex).MochilaType * SLOTS_PER_ROW_INVENTORY
    Else
        getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
    End If
End Function

Public Sub goHome(ByVal Userindex As Integer)
    Dim Distance As Long
    Dim Tiempo   As Long
    With UserList(Userindex)
        If .flags.Muerto = 1 Then
            If .flags.lastMap = 0 Then
                Distance = distanceToCities(.Pos.Map).distanceToCity(.Hogar)
            Else
                Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY
            End If
            Tiempo = (Distance + 1) * 20
            If Tiempo > 60 Then
                Tiempo = 60
            End If
            Call IntervaloGoHome(Userindex, Tiempo * 1000, True)
            Call WriteMultiMessage(Userindex, eMessages.Home, Distance, Tiempo, , MapInfo(Ciudades(.Hogar).Map).Name)
        Else
            Call WriteConsoleMsg(Userindex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With
End Sub

Public Function ToogleToAtackable(ByVal Userindex As Integer, ByVal OwnerIndex As Integer, Optional ByVal StealingNpc As Boolean = True) As Boolean
    Dim AtacablePor As Integer
    With UserList(Userindex)
        If MapInfo(.Pos.Map).Pk = False Then
            Call WriteConsoleMsg(Userindex, "No puedes robar npcs en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        AtacablePor = .flags.AtacablePor
        If AtacablePor > 0 Then
            If StealingNpc Then
                If AtacablePor <> OwnerIndex Then
                    Call WriteConsoleMsg(Userindex, "No puedes atacar otra criatura con dueno hasta que haya terminado tu castigo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            Else
                Call IntervaloEstadoAtacable(Userindex, True)
                ToogleToAtackable = True
                Exit Function
            End If
        End If
        .flags.AtacablePor = OwnerIndex
        Call RefreshCharStatus(Userindex)
        Call IntervaloEstadoAtacable(Userindex, True)
        ToogleToAtackable = True
    End With
End Function

Public Sub setHome(ByVal Userindex As Integer, ByVal newHome As eCiudad, ByVal NpcIndex As Integer)
    If newHome < eCiudad.cUllathorpe Or newHome > eCiudad.cLastCity - 1 Then Exit Sub
    If UserList(Userindex).Hogar <> newHome Then
        UserList(Userindex).Hogar = newHome
        Call WriteChatOverHead(Userindex, "Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    Else
        Call WriteChatOverHead(Userindex, "Ya eres miembro de nuestra humilde comunidad!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    End If
End Sub

Public Function GetHomeArrivalTime(ByVal Userindex As Integer) As Integer
    Dim TActual As Long
    TActual = GetTickCount()
    With UserList(Userindex)
        GetHomeArrivalTime = (.Counters.goHome - TActual) * 0.001
    End With
End Function

Public Sub HomeArrival(ByVal Userindex As Integer)
    Dim tX   As Integer
    Dim tY   As Integer
    Dim tMap As Integer
    With UserList(Userindex)
        If .flags.Navegando = 1 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .flags.Navegando = 0
            Call WriteNavigateToggle(Userindex)
        End If
        tX = Ciudades(.Hogar).X
        tY = Ciudades(.Hogar).Y
        tMap = Ciudades(.Hogar).Map
        Call FindLegalPos(Userindex, tMap, tX, tY)
        Call WarpUserChar(Userindex, tMap, tX, tY, True)
        Call WriteMultiMessage(Userindex, eMessages.FinishHome)
        .flags.Traveling = 0
        .Counters.goHome = 0
    End With
End Sub

