Attribute VB_Name = "SistemaCombate"
Option Explicit

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal Userindex As Integer) As Long
    PoderEvasionEscudo = (UserList(Userindex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList(Userindex).Clase).Escudo) / 2
End Function

Private Function PoderEvasion(ByVal Userindex As Integer) As Long
    Dim lTemp As Long
    With UserList(Userindex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).Evasion
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueArma(ByVal Userindex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    With UserList(Userindex)
        If .Stats.UserSkills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        End If
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal Userindex As Integer) As Long
    Dim PoderAtaqueTemp  As Long
    Dim SkillProyectiles As Integer
    With UserList(Userindex)
        SkillProyectiles = .Stats.UserSkills(eSkill.Proyectiles)
        If SkillProyectiles < 31 Then
            PoderAtaqueTemp = SkillProyectiles * ModClase(.Clase).AtaqueProyectiles
        ElseIf SkillProyectiles < 61 Then
            PoderAtaqueTemp = (SkillProyectiles + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        ElseIf SkillProyectiles < 91 Then
            PoderAtaqueTemp = (SkillProyectiles + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (SkillProyectiles + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        End If
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueWrestling(ByVal Userindex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    Dim WrestlingSkill  As Integer
    With UserList(Userindex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        If WrestlingSkill < 31 Then
            PoderAtaqueTemp = WrestlingSkill * ModClase(.Clase).AtaqueWrestling
        ElseIf WrestlingSkill < 61 Then
            PoderAtaqueTemp = (WrestlingSkill + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueWrestling
        ElseIf WrestlingSkill < 91 Then
            PoderAtaqueTemp = (WrestlingSkill + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueWrestling
        Else
            PoderAtaqueTemp = (WrestlingSkill + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueWrestling
        End If
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal Userindex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim PoderAtaque As Long
    Dim Arma        As Integer
    Dim Skill       As eSkill
    Dim ProbExito   As Long
    Dim MunicionObjIndex    As Integer
    Arma = UserList(Userindex).Invent.WeaponEqpObjIndex
    If Arma > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(Userindex)
            Skill = eSkill.Proyectiles
            MunicionObjIndex = UserList(Userindex).Invent.MunicionEqpObjIndex
            If Not (UserList(Userindex).Clase = eClass.Hunter And UserList(Userindex).flags.Oculto = 1) Then
                If MunicionObjIndex <> 0 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageProyectil(Userindex, UserList(Userindex).Char.CharIndex, Npclist(NpcIndex).Char.CharIndex, ObjData(UserList(Userindex).Invent.MunicionEqpObjIndex).GrhIndex))
                End If
                If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Acuchilla = 1 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageProyectil(Userindex, UserList(Userindex).Char.CharIndex, Npclist(NpcIndex).Char.CharIndex, ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).GrhIndex))
                End If
            End If
        Else
            PoderAtaque = PoderAtaqueArma(Userindex)
            Skill = eSkill.Armas
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(Userindex)
        Skill = eSkill.Wrestling
    End If
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    If UserImpactoNpc Then
        Call SubirSkill(Userindex, Skill, True)
    Else
        Call SubirSkill(Userindex, Skill, False)
    End If
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Boolean
    Dim Rechazo           As Boolean
    Dim ProbRechazo       As Long
    Dim ProbExito         As Long
    Dim UserEvasion       As Long
    Dim NpcPoderAtaque    As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas     As Long
    Dim SkillDefensa      As Long
    UserEvasion = PoderEvasion(Userindex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(Userindex)
    SkillTacticas = UserList(Userindex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(Userindex).Stats.UserSkills(eSkill.Defensa)
    If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Else
                ProbRechazo = 10
            End If
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_ESCUDO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                Call WriteMultiMessage(Userindex, eMessages.BlockedWithShieldUser)
                Call SubirSkill(Userindex, eSkill.Defensa, True)
            Else
                Call SubirSkill(Userindex, eSkill.Defensa, False)
            End If
        End If
    End If
End Function

Public Function CalcularDano(ByVal Userindex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
    Dim DanoArma    As Long
    Dim DanoUsuario As Long
    Dim Arma        As ObjData
    Dim ModifClase  As Single
    Dim proyectil   As ObjData
    Dim DanoMaxArma As Long
    Dim DanoMinArma As Long
    Dim ObjIndex    As Integer
    Dim matoDragon  As Boolean
    matoDragon = False
    With UserList(Userindex)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DanoProyectiles
                    DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DanoMaxArma = Arma.MaxHIT
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DanoArma = DanoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    End If
                Else
                    ModifClase = ModClase(.Clase).DanoArmas
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        If Npclist(NpcIndex).NPCtype = DRAGON Then
                            DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                            DanoMaxArma = Arma.MaxHIT
                            matoDragon = True
                        Else
                            DanoArma = 1
                            DanoMaxArma = 1
                        End If
                    Else
                        DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DanoMaxArma = Arma.MaxHIT
                    End If
                End If
            Else
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DanoProyectiles
                    DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DanoMaxArma = Arma.MaxHIT
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DanoArma = DanoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    End If
                Else
                    ModifClase = ModClase(.Clase).DanoArmas
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        ModifClase = ModClase(.Clase).DanoArmas
                        DanoArma = 1 ' Si usa la espada mataDragones dano es 1
                        DanoMaxArma = 1
                    Else
                        DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DanoMaxArma = Arma.MaxHIT
                    End If
                End If
            End If
        Else
            ModifClase = ModClase(.Clase).DanoWrestling
            DanoMinArma = 4
            DanoMaxArma = 9
            ObjIndex = .Invent.AnilloEqpObjIndex
            If ObjIndex > 0 Then
                If ObjData(ObjIndex).Guante = 1 Then
                    DanoMinArma = DanoMinArma + ObjData(ObjIndex).MinHIT
                    DanoMaxArma = DanoMaxArma + ObjData(ObjIndex).MaxHIT
                End If
            End If
            DanoArma = RandomNumber(DanoMinArma, DanoMaxArma)
        End If
        DanoUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        If matoDragon Then
            CalcularDano = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
        Else
            CalcularDano = (3 * DanoArma + ((DanoMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + DanoUsuario) * ModifClase
        End If
    End With
End Function

Public Sub UserDanoNpc(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
    Dim dano                                 As Long
    Dim DanoBase                             As Long
    Dim PI                                   As Integer
    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    Dim Text                                 As String
    Dim i                                    As Integer
    Dim BoatIndex                            As Integer
    DanoBase = CalcularDano(Userindex, NpcIndex)
    If UserList(Userindex).flags.Navegando = 1 Then
        BoatIndex = UserList(Userindex).Invent.BarcoObjIndex
        If BoatIndex > 0 Then
            DanoBase = DanoBase + RandomNumber(ObjData(BoatIndex).MinHIT, ObjData(BoatIndex).MaxHIT)
        End If
    End If
    With Npclist(NpcIndex)
        dano = DanoBase - .Stats.def
        If dano < 0 Then dano = 0
        Call WriteMultiMessage(Userindex, eMessages.UserHitNPC, dano)
        Call CalcularDarExp(Userindex, NpcIndex, dano)
        .Stats.MinHp = .Stats.MinHp - dano
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
        If .Stats.MinHp > 0 Then
            If PuedeApunalar(Userindex) Then
                If UserList(Userindex).Clase <> eClass.Assasin Then
                    DanoBase = dano
                End If
                Call DoApunalar(Userindex, NpcIndex, 0, DanoBase)
            End If
            Call DoGolpeCritico(Userindex, NpcIndex, 0, dano)
            If PuedeAcuchillar(Userindex) Then
                Call DoAcuchillar(Userindex, NpcIndex, 0, dano)
            End If
        End If
        If .Stats.MinHp <= 0 Then
            If .NPCtype = DRAGON Then
                If UserList(Userindex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                    Call QuitarObjetos(EspadaMataDragonesIndex, 1, Userindex)
                End If
                If .Stats.MaxHp > 100000 Then
                    Text = UserList(Userindex).Name & " mato un dragon"
                    PI = UserList(Userindex).PartyIndex
                    If PI > 0 Then
                        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline())
                        Text = Text & " estando en party "
                        For i = 1 To PARTY_MAXMEMBERS
                            If MembersOnline(i) > 0 Then
                                Text = Text & UserList(MembersOnline(i)).Name & ", "
                            End If
                        Next i
                        Text = Left$(Text, Len(Text) - 2) & ")"
                    End If
                    Call LogDesarrollo(Text & ".")
                End If
            End If
            For i = 1 To MAXMASCOTAS
                If UserList(Userindex).MascotasIndex(i) > 0 Then
                    If Npclist(UserList(Userindex).MascotasIndex(i)).TargetNPC = NpcIndex Then
                        Npclist(UserList(Userindex).MascotasIndex(i)).TargetNPC = 0
                        Npclist(UserList(Userindex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
                    End If
                End If
            Next i
            Call MuereNpc(NpcIndex, Userindex)
        End If
    End With
End Sub

Public Sub NpcDano(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
    Dim dano        As Integer
    Dim Lugar       As Integer
    Dim obj         As ObjData
    Dim BoatDefense As Integer
    Dim HeadDefense As Integer
    Dim BodyDefense As Integer
    Dim BoatIndex   As Integer
    Dim HelmetIndex As Integer
    Dim ArmourIndex As Integer
    Dim ShieldIndex As Integer
    dano = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
    With UserList(Userindex)
        If .flags.Navegando = 1 Then
            BoatIndex = .Invent.BarcoObjIndex
            If BoatIndex > 0 Then
                obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(obj.MinDef, obj.MaxDef)
            End If
        End If
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                HelmetIndex = .Invent.CascoEqpObjIndex
                If HelmetIndex > 0 Then
                    obj = ObjData(HelmetIndex)
                    HeadDefense = RandomNumber(obj.MinDef, obj.MaxDef)
                End If
                
            Case Else
                Dim MinDef As Integer
                Dim MaxDef As Integer
                ArmourIndex = .Invent.ArmourEqpObjIndex
                If ArmourIndex > 0 Then
                    obj = ObjData(ArmourIndex)
                    MinDef = obj.MinDef
                    MaxDef = obj.MaxDef
                End If
                ShieldIndex = .Invent.EscudoEqpObjIndex
                If ShieldIndex > 0 Then
                    obj = ObjData(ShieldIndex)
                    MinDef = MinDef + obj.MinDef
                    MaxDef = MaxDef + obj.MaxDef
                End If
                BodyDefense = RandomNumber(MinDef, MaxDef)
        End Select
        dano = dano - HeadDefense - BodyDefense - BoatDefense
        If dano < 1 Then dano = 1
        Call WriteMultiMessage(Userindex, eMessages.NPCHitUser, Lugar, dano)
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHp = .Stats.MinHp - dano
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
        If .flags.Meditando Then
            If dano > Fix(.Stats.MinHp / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(Userindex)
                Call WriteConsoleMsg(Userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                .Char.FX = 0
                .Char.loops = 0
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If
        End If
        If .Stats.MinHp <= 0 Then
            Call WriteMultiMessage(Userindex, eMessages.NPCKillUser)
            If criminal(Userindex) Then
                If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    Call RestarCriminalidad(Userindex)
                End If
            End If
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
            Else
                With Npclist(NpcIndex)
                    If .Stats.Alineacion = 0 Then
                        .Movement = .flags.OldMovement
                        .Hostile = .flags.OldHostil
                        .flags.AttackedBy = vbNullString
                    End If
                End With
            End If
            Call UserDie(Userindex)
        End If
    End With
End Sub

Public Sub RestarCriminalidad(ByVal Userindex As Integer)
    Dim EraCriminal As Boolean
    EraCriminal = criminal(Userindex)
    With UserList(Userindex).Reputacion
        If .BandidoRep > 0 Then
            .BandidoRep = .BandidoRep - vlASALTO
            If .BandidoRep < 0 Then .BandidoRep = 0
        ElseIf .LadronesRep > 0 Then
            .LadronesRep = .LadronesRep - (vlCAZADOR * 10)
            If .LadronesRep < 0 Then .LadronesRep = 0
        End If
        If EraCriminal And Not criminal(Userindex) Then
            If esCaos(Userindex) Then Call ExpulsarFaccionCaos(Userindex)
            Call RefreshCharStatus(Userindex)
        End If
    End With
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal Userindex As Integer, Optional ByVal CheckElementales As Boolean = True)
    Dim j As Integer
    If UserList(Userindex).NroMascotas = 0 Then Exit Sub
    If Not PuedeAtacarNPC(Userindex, NpcIndex, , True) Then Exit Sub
    With UserList(Userindex)
        For j = 1 To MAXMASCOTAS
            If .MascotasIndex(j) > 0 Then
                If .MascotasIndex(j) <> NpcIndex Then
                    If CheckElementales Or (Npclist(.MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(.MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                        If Npclist(.MascotasIndex(j)).TargetNPC = 0 Then Npclist(.MascotasIndex(j)).TargetNPC = NpcIndex
                        Npclist(.MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
                    End If
                End If
            End If
        Next j
    End With
End Sub

Public Sub AllFollowAmo(ByVal Userindex As Integer)
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(Userindex).MascotasIndex(j))
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Boolean
    With UserList(Userindex)
        If .flags.AdminInvisible = 1 Then Exit Function
        If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)
        End If
    End With
    With Npclist(NpcIndex)
        If IntervaloPermiteAtacarNpc(NpcIndex) Then
            NpcAtacaUser = True
            Call CheckPets(NpcIndex, Userindex, False)
            If .Target = 0 Then .Target = Userindex
            If UserList(Userindex).flags.AtacadoPorNpc = 0 And UserList(Userindex).flags.AtacadoPorUser = 0 Then
                UserList(Userindex).flags.AtacadoPorNpc = NpcIndex
            End If
        Else
            NpcAtacaUser = False
            Exit Function
        End If
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
    End With
    If NpcImpacto(NpcIndex, Userindex) Then
        With UserList(Userindex)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            Call NpcDano(NpcIndex, Userindex)
            Call WriteUpdateHP(Userindex)
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(Userindex)
        End With
        Call SubirSkill(Userindex, eSkill.Tacticas, False)
    Else
        Call WriteMultiMessage(Userindex, eMessages.NPCSwing)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1, DAMAGE_FALLO))
        Call SubirSkill(Userindex, eSkill.Tacticas, True)
    End If
    Call CheckUserLevel(Userindex)
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
    Dim PoderAtt  As Long
    Dim PoderEva  As Long
    Dim ProbExito As Long
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDanoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    Dim dano        As Integer
    With Npclist(Atacante)
        dano = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - dano
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessageCreateDamage(Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, dano, DAMAGE_NORMAL))
        If Npclist(Victima).Stats.MinHp < 1 Then
            .Movement = .flags.OldMovement
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil
            End If
            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)
            End If
            Call MuereNpc(Victima, .MaestroUser)
        End If
    End With
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
    Dim MasterIndex As Integer
    With Npclist(Atacante)
        If Npclist(Victima).NPCtype = eNPCType.Pretoriano Then
            If Not ClanPretoriano(Npclist(Victima).ClanIndex).CanAtackMember(Victima) Then
                Call WriteConsoleMsg(.MaestroUser, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
                .TargetNPC = 0
                Exit Sub
            End If
        End If
        If IntervaloPermiteAtacarNpc(Atacante) Then
            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
            End If
        Else
            Exit Sub
        End If
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
        MasterIndex = .MaestroUser
        If MasterIndex > 0 Then
            If Npclist(Victima).Owner = MasterIndex Then
                Call IntervaloPerdioNpc(MasterIndex, True)
            End If
        End If
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
            Call NpcDanoNpc(Atacante, Victima)
        Else
            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessageCreateDamage(Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, 1, DAMAGE_FALLO))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessageCreateDamage(Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, 1, DAMAGE_FALLO))
            End If
        End If
    End With
End Sub

Public Function UsuarioAtacaNpc(ByVal Userindex As Integer, ByVal NpcIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    If Not PuedeAtacarNPC(Userindex, NpcIndex) Then Exit Function
    With UserList(Userindex)
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)
        End If
    End With
    Call NPCAtacado(NpcIndex, Userindex)
    If UserImpactoNpc(Userindex, NpcIndex) Then
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        End If
        Call UserDanoNpc(Userindex, NpcIndex)
    Else
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SWING, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, 1, DAMAGE_FALLO))
        Call WriteMultiMessage(Userindex, eMessages.UserSwing)
    End If
    Call QuitarSta(Userindex, RandomNumber(1, 10))
    UserList(Userindex).flags.Ignorado = False
    UsuarioAtacaNpc = True
    Exit Function
ErrorHandler:
    Dim UserName As String
    If Userindex > 0 Then UserName = UserList(Userindex).Name
    Call LogError("Error en UsuarioAtacaNpc. Error " & Err.Number & " : " & Err.description & ". User: " & Userindex & "-> " & UserName & ". NpcIndex: " & NpcIndex & ".")
End Function

Public Sub UsuarioAtaca(ByVal Userindex As Integer)
    Dim index     As Integer
    Dim AttackPos As WorldPos
    If Not IntervaloPermiteUsarArcos(Userindex, False) Then Exit Sub
    If Not IntervaloPermiteMagiaGolpe(Userindex) Then
        If Not IntervaloPermiteAtacar(Userindex) Then
            Exit Sub
        End If
    End If
    With UserList(Userindex)
        #If ProteccionGM = 1 Then
            If (.flags.Privilegios And PlayerType.User) = 0 Then
                Call WriteConsoleMsg(Userindex, "Los GMs no pueden atacar.", FONTTYPE_SERVER)
                Exit Sub
            End If
        #End If
        If .Stats.MinSta < 10 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(Userindex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Exit Sub
        End If
        index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).Userindex
        If index > 0 Then
            Call UsuarioAtacaUsuario(Userindex, index)
            Call WriteUpdateUserStats(Userindex)
            Call WriteUpdateUserStats(index)
            Exit Sub
        End If
        index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.Map).Pk = False Then
                    Call WriteConsoleMsg(Userindex, "No puedes atacar mascotas en zona segura.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                End If
                Call UsuarioAtacaNpc(Userindex, index)
            Else
                Call WriteConsoleMsg(Userindex, "No puedes atacar a este NPC.", FontTypeNames.FONTTYPE_WARNING)
            End If
            Call WriteUpdateUserStats(Userindex)
            Exit Sub
        End If
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        Call WriteUpdateUserStats(Userindex)
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim ProbRechazo            As Long
    Dim Rechazo                As Boolean
    Dim ProbExito              As Long
    Dim PoderAtaque            As Long
    Dim UserPoderEvasion       As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma                   As Integer
    Dim SkillTacticas          As Long
    Dim SkillDefensa           As Long
    Dim ProbEvadir             As Long
    Dim Skill                  As eSkill
    With UserList(VictimaIndex)
        SkillTacticas = .Stats.UserSkills(eSkill.Tacticas)
        SkillDefensa = .Stats.UserSkills(eSkill.Defensa)
        Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
        UserPoderEvasion = PoderEvasion(VictimaIndex)
        If .Invent.EscudoEqpObjIndex > 0 Then
            UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
            UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
            If UserList(AtacanteIndex).Invent.MunicionEqpObjIndex > 0 Then
                If Not (UserList(AtacanteIndex).Clase = eClass.Hunter And UserList(AtacanteIndex).flags.Oculto = 1) Then
                    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessageProyectil(AtacanteIndex, UserList(AtacanteIndex).Char.CharIndex, .Char.CharIndex, ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).GrhIndex))
                End If
            End If
        Else
            UserPoderEvasionEscudo = 0
        End If
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(Arma).proyectil = 1 Then
                PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
                Skill = eSkill.Proyectiles
            Else
                PoderAtaque = PoderAtaqueArma(AtacanteIndex)
                Skill = eSkill.Armas
            End If
        Else
            PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
            Skill = eSkill.Wrestling
        End If
        ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
        If .flags.Meditando Then
            ProbEvadir = (100 - ProbExito) * 0.75
            ProbExito = MinimoInt(90, 100 - ProbEvadir)
        End If
        UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
        If .Invent.EscudoEqpObjIndex > 0 Then
            If Not UsuarioImpacto Then
                Dim SumaSkills As Integer
                SumaSkills = MaximoInt(1, SkillDefensa + SkillTacticas)
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / SumaSkills))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                If Rechazo Then
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, .Pos.X, .Pos.Y))
                    Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
                    Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, True)
                Else
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, False)
                End If
            End If
        End If
        If Not UsuarioImpacto Then
            Call SubirSkill(AtacanteIndex, Skill, False)
        End If
    End With
    Exit Function
ErrorHandler:
    Dim AtacanteNick As String
    Dim VictimaNick  As String
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    Call LogError("Error en UsuarioImpacto. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    With UserList(AtacanteIndex)
        If .flags.Equitando = 1 Then
            Call UnmountMontura(AtacanteIndex)
            Call WriteEquitandoToggle(AtacanteIndex)
        End If
        If Abs(.Pos.X - UserList(VictimaIndex).Pos.X) > RANGO_VISION_X Or Abs(.Pos.Y - UserList(VictimaIndex).Pos.Y) > RANGO_VISION_Y Then
            Call WriteConsoleMsg(AtacanteIndex, "Estas muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If
            If .Clase = eClass.Bandit Then
                Call DoDesequipar(AtacanteIndex, VictimaIndex)
            ElseIf .Clase = eClass.Thief Then
                Call DoHandInmo(AtacanteIndex, VictimaIndex)
            End If
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, False)
            Call UserDanoUser(AtacanteIndex, VictimaIndex)
        Else
            If .flags.AdminInvisible = 1 Then
                Call UserList(AtacanteIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            End If
            Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessageCreateDamage(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, 1, DAMAGE_FALLO))
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, True)
        End If
        If .Clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)
    End With
    UsuarioAtacaUsuario = True
    Exit Function
ErrorHandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.description)
End Function

Public Sub UserDanoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    On Error GoTo ErrorHandler
    Dim dano          As Long
    Dim Lugar         As Byte
    Dim obj           As ObjData
    Dim BoatDefense   As Integer
    Dim BodyDefense   As Integer
    Dim HeadDefense   As Integer
    Dim WeaponBoost   As Integer
    Dim BoatIndex     As Integer
    Dim WeaponIndex   As Integer
    Dim HelmetIndex   As Integer
    Dim ArmourIndex   As Integer
    Dim ShieldIndex   As Integer
    Dim BarcaIndex    As Integer
    Dim ArmaIndex     As Integer
    Dim CascoIndex    As Integer
    Dim ArmaduraIndex As Integer
    dano = CalcularDano(AtacanteIndex)
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 Then
            BoatIndex = .Invent.BarcoObjIndex
            If BoatIndex > 0 Then
                obj = ObjData(BoatIndex)
                dano = dano + RandomNumber(obj.MinHIT, obj.MaxHIT)
            End If
        End If
        If UserList(VictimaIndex).flags.Navegando = 1 Then
            BoatIndex = UserList(VictimaIndex).Invent.BarcoObjIndex
            If BoatIndex > 0 Then
                obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(obj.MinDef, obj.MaxDef)
            End If
        End If
        WeaponIndex = .Invent.WeaponEqpObjIndex
        If WeaponIndex > 0 Then
            WeaponBoost = ObjData(WeaponIndex).Refuerzo
        End If
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                HelmetIndex = UserList(VictimaIndex).Invent.CascoEqpObjIndex
                If HelmetIndex > 0 Then
                    obj = ObjData(HelmetIndex)
                    HeadDefense = RandomNumber(obj.MinDef, obj.MaxDef)
                End If
                
            Case Else
                Dim MinDef As Integer
                Dim MaxDef As Integer
                ArmourIndex = UserList(VictimaIndex).Invent.ArmourEqpObjIndex
                If ArmourIndex > 0 Then
                    obj = ObjData(ArmourIndex)
                    MinDef = obj.MinDef
                    MaxDef = obj.MaxDef
                End If
                ShieldIndex = UserList(VictimaIndex).Invent.EscudoEqpObjIndex
                If ShieldIndex > 0 Then
                    obj = ObjData(ShieldIndex)
                    MinDef = MinDef + obj.MinDef
                    MaxDef = MaxDef + obj.MaxDef
                End If
                BodyDefense = RandomNumber(MinDef, MaxDef)
        End Select
        dano = dano + WeaponBoost - HeadDefense - BodyDefense - BoatDefense
        If dano < 0 Then dano = 1
        Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, Lugar, dano)
        Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, Lugar, dano)
        UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - dano
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            If WeaponIndex > 0 Then
                If ObjData(WeaponIndex).proyectil Then
                    Call SubirSkill(AtacanteIndex, eSkill.Proyectiles, True)
                    If PuedeAcuchillar(AtacanteIndex) Then
                        Call DoAcuchillar(AtacanteIndex, 0, VictimaIndex, dano)
                    End If
                Else
                    Call SubirSkill(AtacanteIndex, eSkill.Armas, True)
                End If
            Else
                Call SubirSkill(AtacanteIndex, eSkill.Wrestling, True)
            End If
            If PuedeApunalar(AtacanteIndex) Then
                Call DoApunalar(AtacanteIndex, 0, VictimaIndex, dano)
            End If
            Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, dano)
        End If
        If Not PuedeApunalar(AtacanteIndex) Then
            Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateDamage(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, dano, DAMAGE_NORMAL))
        End If
        If UserList(VictimaIndex).Stats.MinHp <= 0 Then
            If UserList(VictimaIndex).flags.AtacablePor <> AtacanteIndex Then
                Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
                Call ContarMuerte(VictimaIndex, AtacanteIndex)
            End If
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If .MascotasIndex(j) > 0 Then
                    If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
                        Npclist(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))
                    End If
                End If
            Next j
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex, AtacanteIndex)
        Else
            Call WriteUpdateHP(VictimaIndex)
        End If
    End With
    Call CheckUserLevel(AtacanteIndex)
    Exit Sub
ErrorHandler:
    Dim AtacanteNick As String
    Dim VictimaNick  As String
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    Call LogError("Error en UserDanoUser. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    Dim EraCriminal       As Boolean
    Dim VictimaEsAtacable As Boolean
    If Not criminal(AttackerIndex) Then
        If Not criminal(VictimIndex) Then
            VictimaEsAtacable = UserList(VictimIndex).flags.AtacablePor = AttackerIndex
            If Not VictimaEsAtacable Then Call VolverCriminal(AttackerIndex)
        End If
    End If
    With UserList(VictimIndex)
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(VictimIndex)
            Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
    EraCriminal = criminal(AttackerIndex)
    If Not VictimaEsAtacable Then
        With UserList(AttackerIndex).Reputacion
            If Not criminal(VictimIndex) Then
                .BandidoRep = .BandidoRep + vlASALTO
                If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
                .NobleRep = .NobleRep * 0.5
                If .NobleRep < 0 Then .NobleRep = 0
            Else
                .NobleRep = .NobleRep + vlNoble
                If .NobleRep > MAXREP Then .NobleRep = MAXREP
            End If
        End With
    End If
    If criminal(AttackerIndex) Then
        If UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(AttackerIndex)
        If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
    ElseIf EraCriminal Then
        Call RefreshCharStatus(AttackerIndex)
    End If
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    Call CancelExit(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    Dim iCount As Integer
    For iCount = 1 To MAXMASCOTAS
        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
        End If
    Next iCount
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un espiritu.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    If UserList(AttackerIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(VictimIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = (UserList(VictimIndex).flags.AdminInvisible = 0)
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
    End Select
    If Not criminal(VictimIndex) Then
        If Not criminal(AttackerIndex) Then
            If esArmada(AttackerIndex) Then
                If esArmada(VictimIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "Los soldados del ejercito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Function
                End If
            End If
            If UserList(VictimIndex).flags.AtacablePor = AttackerIndex Then
                If ToogleToAtackable(AttackerIndex, VictimIndex, False) Then
                    PuedeAtacar = True
                    Exit Function
                End If
            End If
        End If
    Else
        If esCaos(VictimIndex) Then
            If esCaos(AttackerIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legion oscura tienen prohibido atacarse entre si.", FontTypeNames.FONTTYPE_WARNING)
                Exit Function
            End If
        End If
    End If
    If UserList(AttackerIndex).flags.Seguro Then
        If Not criminal(VictimIndex) Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
    Else
        If Not criminal(VictimIndex) Then
            If esArmada(AttackerIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "Los soldados del ejercito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
        End If
    End If
    If MapInfo(UserList(VictimIndex).Pos.Map).Pk = False Then
        If esArmada(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasReal > 11 Then
                If UserList(VictimIndex).Pos.Map = 58 Or UserList(VictimIndex).Pos.Map = 59 Or UserList(VictimIndex).Pos.Map = 60 Then
                    Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! Estas siendo atacado y no podras defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True
                    Exit Function
                End If
            End If
        End If
        If esCaos(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasCaos > 11 Then
                If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
                    Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! Estas siendo atacado y no podras defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True
                    Exit Function
                End If
            End If
        End If
        Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aqui no puedes atacar a otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or MapData(UserList(AttackerIndex).Pos.Map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    PuedeAtacar = True
    Exit Function
ErrorHandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.description)
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal Paraliza As Boolean = False, Optional ByVal IsPet As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        If UserList(AttackerIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(AttackerIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If UserList(AttackerIndex).flags.Privilegios And PlayerType.Consejero Then
            Exit Function
        End If
        If UserList(AttackerIndex).flags.EnConsulta Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If .Attackable = 0 Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If Abs(.Pos.X - UserList(AttackerIndex).Pos.X) > RANGO_VISION_X Or Abs(.Pos.Y - UserList(AttackerIndex).Pos.Y) > RANGO_VISION_Y Then
            Call WriteConsoleMsg(AttackerIndex, "Estas muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
        If .Hostile = 0 Then
            If .NPCtype = eNPCType.Guardiascaos Then
                If esCaos(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias del Caos siendo de la legion oscura.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            ElseIf .NPCtype = eNPCType.GuardiaReal Then
                If esArmada(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias Reales siendo del ejercito real.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Para poder atacar Guardias Reales debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else
                    Call WriteConsoleMsg(AttackerIndex, "Atacaste un Guardia Real! Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(AttackerIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                End If
            ElseIf .MaestroUser = 0 Then
                If Not criminal(AttackerIndex) Then
                    If esArmada(AttackerIndex) Then
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejercito real no pueden atacar npcs no hostiles.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar a este NPC debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    Else
                        Call WriteConsoleMsg(AttackerIndex, "Atacaste un NPC no-hostil. Continua haciendolo y te podras convertir en criminal.", FontTypeNames.FONTTYPE_INFO)
                        Call DisNobAuBan(AttackerIndex, 0, 1000)
                        PuedeAtacarNPC = True
                        Exit Function
                    End If
                End If
            End If
        End If
        Dim MasterIndex As Integer
        MasterIndex = .MaestroUser
        If MasterIndex > 0 Then
            If Not criminal(MasterIndex) Then
                If Not criminal(AttackerIndex) Then
                    If UserList(MasterIndex).flags.AtacablePor = AttackerIndex Then
                        Call ToogleToAtackable(AttackerIndex, MasterIndex)
                        PuedeAtacarNPC = True
                        Exit Function
                    End If
                    If esArmada(AttackerIndex) Then
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejercito real no pueden atacar mascotas de ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    Else
                        Call WriteConsoleMsg(AttackerIndex, "Has atacado la Mascota de un ciudadano. Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                        Call VolverCriminal(AttackerIndex)
                        PuedeAtacarNPC = True
                        Exit Function
                    End If
                Else
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
            ElseIf esCaos(MasterIndex) Then
                If esCaos(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legion oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        ElseIf .Owner > 0 Then
            Dim OwnerUserIndex As Integer
            OwnerUserIndex = .Owner
            If OwnerUserIndex = AttackerIndex Then
                PuedeAtacarNPC = True
                Call IntervaloPerdioNpc(OwnerUserIndex, True)
                Exit Function
            End If
            If UserList(OwnerUserIndex).flags.ShareNpcWith = AttackerIndex Then
                PuedeAtacarNPC = True
                Exit Function
            End If
            If Not SameClan(OwnerUserIndex, AttackerIndex) And Not SameParty(OwnerUserIndex, AttackerIndex) Then
                If IntervaloPerdioNpc(OwnerUserIndex) Then
                    Call PerdioNpc(OwnerUserIndex)
                    Call ApropioNpc(AttackerIndex, NpcIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                ElseIf Paraliza Then
                    If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
                        If Not criminal(AttackerIndex) And Not criminal(OwnerUserIndex) Then
                            If esArmada(AttackerIndex) Then
                                If esArmada(OwnerUserIndex) Then
                                    Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ejercito Real no pueden paralizar criaturas ya paralizadas pertenecientes a otros miembros del Ejercito Real", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                Else
                                    If UserList(AttackerIndex).flags.Seguro Then
                                        Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Function
                                    Else
                                        If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                            Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                            PuedeAtacarNPC = True
                                        End If
                                        Exit Function
                                    End If
                                End If
                            Else
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                Else
                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True
                                    End If
                                    Exit Function
                                End If
                            End If
                        Else
                            If esCaos(AttackerIndex) And esCaos(OwnerUserIndex) Then
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legion oscura no pueden paralizar criaturas ya paralizadas por otros legionarios.", FontTypeNames.FONTTYPE_INFO)
                                Exit Function
                            End If
                        End If
                    Else
                        If OwnerUserIndex = 0 Then
                            If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
                                Call ApropioNpc(AttackerIndex, NpcIndex)
                            End If
                        ElseIf OwnerUserIndex = AttackerIndex Then
                            Call IntervaloPerdioNpc(OwnerUserIndex, True)
                        End If
                        PuedeAtacarNPC = True
                        Exit Function
                    End If
                Else
                    If Not criminal(OwnerUserIndex) Then
                        If esArmada(AttackerIndex) Then
                            If esArmada(OwnerUserIndex) Then
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ejercito Real no pueden atacar criaturas pertenecientes a otros miembros del Ejercito Real", FontTypeNames.FONTTYPE_INFO)
                                Exit Function
                            Else
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas ya pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                Else
                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has atacado a la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True
                                    End If
                                    Exit Function
                                End If
                            End If
                        Else
                            If Not criminal(AttackerIndex) Then
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                Else
                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has atacado a la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True
                                    End If
                                    Exit Function
                                End If
                            Else
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                End If
                                PuedeAtacarNPC = True
                            End If
                        End If
                    Else
                        If esCaos(OwnerUserIndex) Then
                            If esCaos(AttackerIndex) Then
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la Legion Oscura no pueden atacar criaturas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If Not criminal(AttackerIndex) Or esCaos(AttackerIndex) Then
                If Npclist(NpcIndex).NPCtype <> eNPCType.Pretoriano Then
                    If Npclist(NpcIndex).NPCtype <> DRAGON Then
                        If Not IsPet Then
                            If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
                                Call ApropioNpc(AttackerIndex, NpcIndex)
                            Else
                                If Not Paraliza Then Call ApropioNpc(AttackerIndex, NpcIndex)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
        If Not ClanPretoriano(Npclist(NpcIndex).ClanIndex).CanAtackMember(NpcIndex) Then
            Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ejercito antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
    End If
    PuedeAtacarNPC = True
    Exit Function
ErrorHandler:
    Dim AtckName  As String
    Dim OwnerName As String
    If AttackerIndex > 0 Then AtckName = UserList(AttackerIndex).Name
    If OwnerUserIndex > 0 Then OwnerName = UserList(OwnerUserIndex).Name
    Call LogError("Error en PuedeAtacarNpc. Erorr: " & Err.Number & " - " & Err.description & " Atacante: " & AttackerIndex & "-> " & AtckName & ". Owner: " & OwnerUserIndex & "-> " & OwnerName & ". NpcIndex: " & NpcIndex & ".")
End Function

Private Function SameClan(ByVal Userindex As Integer, ByVal OtherUserIndex As Integer) As Boolean
    SameClan = (UserList(Userindex).GuildIndex = UserList(OtherUserIndex).GuildIndex) And UserList(Userindex).GuildIndex <> 0
End Function

Private Function SameParty(ByVal Userindex As Integer, ByVal OtherUserIndex As Integer) As Boolean
    SameParty = UserList(Userindex).PartyIndex = UserList(OtherUserIndex).PartyIndex And UserList(Userindex).PartyIndex <> 0
End Function

Sub CalcularDarExp(ByVal Userindex As Integer, ByVal NpcIndex As Integer, ByVal ElDano As Long)
    Dim ExpaDar As Long
    If ElDano <= 0 Then ElDano = 0
    If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
    If ElDano > Npclist(NpcIndex).Stats.MinHp Then ElDano = Npclist(NpcIndex).Stats.MinHp
    ExpaDar = CLng(ElDano * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))
    If ExpaDar <= 0 Then Exit Sub
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
    End If
    If ExpaDar > 0 Then
        If UserList(Userindex).PartyIndex > 0 Then
            Call mdParty.ObtenerExito(Userindex, ExpaDar, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
        Else
            UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + ExpaDar
            If UserList(Userindex).Stats.Exp > MAXEXP Then
                UserList(Userindex).Stats.Exp = MAXEXP
            End If
            Call WriteMultiMessage(Userindex, eMessages.EarnExp, ExpaDar)
        End If
        Call CheckUserLevel(Userindex)
    End If
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
    On Error GoTo ErrorHandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If
    Exit Function
ErrorHandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim ObjInd As Integer
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
        End If
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, "" & UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                End If
            End If
        End If
    End If
End Sub

Public Sub LanzarProyectil(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte)
    On Error GoTo ErrorHandler
    Dim MunicionSlot    As Byte
    Dim MunicionIndex   As Integer
    Dim WeaponSlot      As Byte
    Dim WeaponIndex     As Integer
    Dim TargetUserIndex As Integer
    Dim TargetNpcIndex  As Integer
    Dim DummyInt        As Integer
    Dim Threw           As Boolean
    Threw = True
    With UserList(Userindex)
        With .Invent
            MunicionSlot = .MunicionEqpSlot
            MunicionIndex = .MunicionEqpObjIndex
            WeaponSlot = .WeaponEqpSlot
            WeaponIndex = .WeaponEqpObjIndex
        End With
        If WeaponIndex = 0 Then
            DummyInt = 1
            Call WriteConsoleMsg(Userindex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
        ElseIf WeaponSlot < 1 Or WeaponSlot > .CurrentInventorySlots Then
            DummyInt = 1
            Call WriteConsoleMsg(Userindex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
        ElseIf ObjData(WeaponIndex).Municion = 1 Then
            If MunicionSlot < 1 Or MunicionSlot > .CurrentInventorySlots Then
                DummyInt = 1
                Call WriteConsoleMsg(Userindex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
            ElseIf MunicionIndex = 0 Then
                DummyInt = 1
                Call WriteConsoleMsg(Userindex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
            ElseIf ObjData(MunicionIndex).OBJType <> eOBJType.otFlechas Then
                DummyInt = 1
                Call WriteConsoleMsg(Userindex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
            ElseIf .Invent.Object(MunicionSlot).Amount < 1 Then
                DummyInt = 1
                Call WriteConsoleMsg(Userindex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
            End If
        ElseIf ObjData(WeaponIndex).proyectil <> 1 Then
            DummyInt = 2
        End If
        If DummyInt <> 0 Then
            If DummyInt = 1 Then
                Call Desequipar(Userindex, WeaponSlot)
            End If
            Call Desequipar(Userindex, MunicionSlot)
            Exit Sub
        End If
        If .Stats.MinSta >= 10 Then
            Call QuitarSta(Userindex, RandomNumber(1, 10))
        Else
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(Userindex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        Call LookatTile(Userindex, .Pos.Map, X, Y)
        TargetUserIndex = .flags.TargetUser
        TargetNpcIndex = .flags.TargetNPC
        If TargetUserIndex > 0 Then
            If Abs(UserList(TargetUserIndex).Pos.X - .Pos.X) > RANGO_VISION_X Or Abs(UserList(TargetUserIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                Call WriteConsoleMsg(Userindex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            If TargetUserIndex = Userindex Then
                Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Threw = UsuarioAtacaUsuario(Userindex, TargetUserIndex)
        ElseIf TargetNpcIndex > 0 Then
            If Abs(Npclist(TargetNpcIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(TargetNpcIndex).Pos.X - .Pos.X) > RANGO_VISION_X Then
                Call WriteConsoleMsg(Userindex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            If Npclist(TargetNpcIndex).Attackable <> 0 Then
                Threw = UsuarioAtacaNpc(Userindex, TargetNpcIndex)
            End If
        End If
        If Threw Then
            Dim Slot As Byte
            If ObjData(WeaponIndex).Municion = 1 Then
                Slot = MunicionSlot
            Else
                Slot = WeaponSlot
            End If
            Call QuitarUserInvItem(Userindex, Slot, 1)
            Call UpdateUserInv(False, Userindex, Slot)
        End If
    End With
    Exit Sub
ErrorHandler:
    Dim UserName As String
    If Userindex > 0 Then UserName = UserList(Userindex).Name
    Call LogError("Error en LanzarProyectil " & Err.Number & ": " & Err.description & ". User: " & UserName & "(" & Userindex & ")")
End Sub
