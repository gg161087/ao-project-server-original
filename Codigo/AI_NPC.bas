Attribute VB_Name = "AI"
Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    NpcObjeto = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    SacerdotePretorianoAi = 20
    GuerreroPretorianoAi = 21
    MagoPretorianoAi = 22
    CazadorPretorianoAi = 23
    ReyPretoriano = 24
End Enum

Public Const ELEMENTALFUEGO  As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA   As Integer = 92
Private Const VISION_EXTRA         As Byte = 2
Public Const RANGO_VISION_NPC_X    As Byte = RANGO_VISION_X + VISION_EXTRA
Public Const RANGO_VISION_NPC_Y    As Byte = RANGO_VISION_Y + VISION_EXTRA

Private Sub GuardiasAI(ByVal NpcIndex As Integer, ByVal DelCaos As Boolean)
    Dim nPos          As WorldPos
    Dim headingloop   As Byte
    Dim UI            As Integer
    Dim UserProtected As Boolean
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).Userindex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                            If Not DelCaos Then
                                If criminal(UI) Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            Else
                                If Not criminal(UI) Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
    Dim nPos          As WorldPos
    Dim headingloop   As Byte
    Dim UI            As Integer
    Dim NPCI          As Integer
    Dim atacoPJ       As Boolean
    Dim UserProtected As Boolean
    atacoPJ = False
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).Userindex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 And Not atacoPJ Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            atacoPJ = True
                            If .Movement = NpcObjeto Then
                                If RandomNumber(1, 3) = 3 Then atacoPJ = False
                            End If
                            If atacoPJ Then
                                If .flags.LanzaSpells Then
                                    If .flags.AtacaDoble Then
                                        If (RandomNumber(0, 1)) Then
                                            If NpcAtacaUser(NpcIndex, UI) Then
                                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                            End If
                            If NpcAtacaUser(NpcIndex, UI) Then
                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                            End If
                            Exit Sub
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
    Dim nPos          As WorldPos
    Dim headingloop   As eHeading
    Dim UI            As Integer
    Dim UserProtected As Boolean
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).Userindex
                    If UI > 0 Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
    Dim tHeading      As Byte
    Dim Userindex     As Integer
    Dim SignoNS       As Integer
    Dim SignoEO       As Integer
    Dim i             As Long
    Dim UserProtected As Boolean
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                Userindex = Areas.ConnGroups(.Pos.Map).Item(i)
                If Abs(UserList(Userindex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X And Sgn(UserList(Userindex).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(Userindex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y And Sgn(UserList(Userindex).Pos.Y - .Pos.Y) = SignoNS Then
                        UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And UserList(Userindex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado Or UserList(Userindex).flags.EnConsulta
                        If UserList(Userindex).flags.Muerto = 0 Then
                            If Not UserProtected Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, Userindex)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next i
        Else
            Dim OwnerIndex As Integer
            OwnerIndex = .Owner
            If OwnerIndex > 0 Then
                If UserList(OwnerIndex).Pos.Map = .Pos.Map Then
                    If Abs(UserList(OwnerIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                        If Abs(UserList(OwnerIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                            If UserList(OwnerIndex).flags.invisible = 0 And UserList(OwnerIndex).flags.Oculto = 0 And Not UserList(OwnerIndex).flags.EnConsulta And Not UserList(OwnerIndex).flags.Ignorado Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, OwnerIndex)
                                tHeading = FindDirection(.Pos, UserList(OwnerIndex).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    Call LogError("El npc: " & .Name & "(" & NpcIndex & "), intenta atacar a " & UserList(OwnerIndex).Name & "(Index: " & OwnerIndex & ", Mapa: " & UserList(OwnerIndex).Pos.Map & ") desde el mapa " & .Pos.Map)
                    .Owner = 0
                End If
            End If
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                Userindex = Areas.ConnGroups(.Pos.Map).Item(i)
                If Abs(UserList(Userindex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(Userindex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                        With UserList(Userindex)
                            UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And .flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or .flags.Ignorado Or .flags.EnConsulta
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, Userindex)
                                tHeading = FindDirection(Npclist(NpcIndex).Pos, .Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                        End With
                    End If
                End If
            Next i
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
        End If
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI       As Integer
    Dim i        As Long
    Dim SignoNS  As Integer
    Dim SignoEO  As Integer
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UI = Areas.ConnGroups(.Pos.Map).Item(i)
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacara a ciudadanos si eres miembro del ejercito real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.AttackedBy = vbNullString
                                    Exit Sub
                                End If
                            End If
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                Else
                                    If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                        If Npclist(NpcIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NpcIndex, UI)
                                        End If
                                    End If
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next i
        Else
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UI = Areas.ConnGroups(.Pos.Map).Item(i)
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacara a ciudadanos si eres miembro del ejercito real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.AttackedBy = vbNullString
                                    Call FollowAmo(NpcIndex)
                                    Exit Sub
                                End If
                            End If
                            If (UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0) Or (.flags.SiguiendoGm = True) Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                Else
                                    If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                        If Npclist(NpcIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NpcIndex, UI)
                                        End If
                                    End If
                                End If
                                tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
            .flags.SiguiendoGm = False
        End If
    End With
End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
    Dim Userindex     As Integer
    Dim tHeading      As Byte
    Dim i             As Long
    Dim UserProtected As Boolean
    With Npclist(NpcIndex)
        For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
            Userindex = Areas.ConnGroups(.Pos.Map).Item(i)
            If Abs(UserList(Userindex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                If Abs(UserList(Userindex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                    If Not criminal(Userindex) Then
                        UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And UserList(Userindex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado Or UserList(Userindex).flags.EnConsulta
                        If UserList(Userindex).flags.Muerto = 0 And UserList(Userindex).flags.invisible = 0 And UserList(Userindex).flags.Oculto = 0 And UserList(Userindex).flags.AdminPerseguible And Not UserProtected Then
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, Userindex)
                            End If
                            tHeading = FindDirection(.Pos, UserList(Userindex).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next i
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
    Dim Userindex     As Integer
    Dim tHeading      As Byte
    Dim i             As Long
    Dim SignoNS       As Integer
    Dim SignoEO       As Integer
    Dim UserProtected As Boolean
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                Userindex = Areas.ConnGroups(.Pos.Map).Item(i)
                If Abs(UserList(Userindex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X And Sgn(UserList(Userindex).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(Userindex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y And Sgn(UserList(Userindex).Pos.Y - .Pos.Y) = SignoNS Then
                        If criminal(Userindex) Then
                            With UserList(Userindex)
                                UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And .flags.NoPuedeSerAtacado
                                UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado Or UserList(Userindex).flags.EnConsulta
                                If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                    If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NpcIndex, Userindex)
                                    End If
                                    Exit Sub
                                End If
                            End With
                        End If
                    End If
                End If
            Next i
        Else
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                Userindex = Areas.ConnGroups(.Pos.Map).Item(i)
                If Abs(UserList(Userindex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(Userindex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                        If criminal(Userindex) Then
                            UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And UserList(Userindex).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado
                            If UserList(Userindex).flags.Muerto = 0 And UserList(Userindex).flags.invisible = 0 And UserList(Userindex).flags.Oculto = 0 And UserList(Userindex).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, Userindex)
                                End If
                                If .flags.Inmovilizado = 1 Then Exit Sub
                                tHeading = FindDirection(.Pos, UserList(Userindex).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI       As Integer
    With Npclist(NpcIndex)
        If .Target = 0 And .TargetNPC = 0 Then
            UI = .MaestroUser
            If UI > 0 Then
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End With
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim X        As Long
    Dim Y        As Long
    Dim NI       As Integer
    Dim bNoEsta  As Boolean
    Dim SignoNS  As Integer
    Dim SignoEO  As Integer
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_NPC_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_NPC_X Step IIf(SignoEO = 0, 1, SignoEO)
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).NpcIndex
                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True
                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                    End If
                                Else
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                    End If
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                Next X
            Next Y
        Else
            For Y = .Pos.Y - RANGO_VISION_NPC_Y To .Pos.Y + RANGO_VISION_NPC_Y
                For X = .Pos.X - RANGO_VISION_NPC_Y To .Pos.X + RANGO_VISION_NPC_Y
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).NpcIndex
                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True
                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                    End If
                                Else
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                    End If
                                End If
                                If .flags.Inmovilizado = 1 Then Exit Sub
                                If .TargetNPC = 0 Then Exit Sub
                                tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.Map, X, Y).NpcIndex).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                        End If
                    End If
                Next X
            Next Y
        End If
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
End Sub

Public Sub AiNpcObjeto(ByVal NpcIndex As Integer)
    Dim Userindex     As Integer
    Dim i             As Long
    Dim UserProtected As Boolean
    With Npclist(NpcIndex)
        For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
            Userindex = Areas.ConnGroups(.Pos.Map).Item(i)
            If Abs(UserList(Userindex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                If Abs(UserList(Userindex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                    With UserList(Userindex)
                        UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And .flags.NoPuedeSerAtacado
                        If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                            If RandomNumber(1, 3) < 3 Then
                                If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, Userindex)
                                End If
                                Exit Sub
                            End If
                        End If
                    End With
                End If
            End If
        Next i
    End With
End Sub

Sub NPCAI(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        If .MaestroUser = 0 Then
            If .NPCtype = eNPCType.GuardiaReal Then
                Call GuardiasAI(NpcIndex, False)
            ElseIf .NPCtype = eNPCType.Guardiascaos Then
                Call GuardiasAI(NpcIndex, True)
            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
                Call HostilMalvadoAI(NpcIndex)
            ElseIf .Hostile And .Stats.Alineacion = 0 Then
                Call HostilBuenoAI(NpcIndex)
            End If
        Else
            'Call HostilBuenoAI(NpcIndex)
        End If
        Select Case .Movement
            Case TipoAI.MueveAlAzar
                If .flags.Inmovilizado = 1 Then Exit Sub
                If .NPCtype = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCriminal(NpcIndex)
                ElseIf .NPCtype = eNPCType.Guardiascaos Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCiudadano(NpcIndex)
                Else
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                End If
                
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NpcIndex)

            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            
            Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NpcIndex)
            
            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Then Exit Sub
                Call SeguirAmo(NpcIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(NpcIndex)
                
            Case TipoAI.NpcObjeto
                Call AiNpcObjeto(NpcIndex)
                
            Case TipoAI.NpcPathfinding
                If .flags.Inmovilizado = 1 Then Exit Sub
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    If .PFINFO.NoPath Then
                        Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                    End If
                Else
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        .PFINFO.PathLenght = 0
                    End If
                End If
        End Select
    End With
    Exit Sub
ErrorHandler:
    With Npclist(NpcIndex)
        Call LogError("Error en NPCAI. Error: " & Err.Number & " - " & Err.description & ". " & "Npc: " & .Name & ", Index: " & NpcIndex & ", MaestroUser: " & .MaestroUser & ", MaestroNpc: " & .MaestroNpc & ", Mapa: " & .Pos.Map & " x:" & .Pos.X & " y:" & .Pos.Y & " Mov:" & .Movement & " TargU:" & .Target & " TargN:" & .TargetNPC)
    End With
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
End Sub

Function UserNear(ByVal NpcIndex As Integer) As Boolean
    With Npclist(NpcIndex)
        UserNear = Not Int(Distance(.Pos.X, .Pos.Y, UserList(.PFINFO.TargetUser).Pos.X, UserList(.PFINFO.TargetUser).Pos.Y)) > 1
    End With
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True
    End If
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
    PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
    Dim tmpPos   As WorldPos
    Dim tHeading As Byte
    With Npclist(NpcIndex)
        tmpPos.Map = .Pos.Map
        tmpPos.X = .PFINFO.Path(.PFINFO.CurPos).Y
        tmpPos.Y = .PFINFO.Path(.PFINFO.CurPos).X
        tHeading = FindDirection(.Pos, tmpPos)
        MoveNPCChar NpcIndex, tHeading
        .PFINFO.CurPos = .PFINFO.CurPos + 1
    End With
End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
    Dim Y As Long
    Dim X As Long
    With Npclist(NpcIndex)
        For Y = .Pos.Y - 10 To .Pos.Y + 10
            For X = .Pos.X - 10 To .Pos.X + 10
                If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                    If MapData(.Pos.Map, X, Y).Userindex > 0 Then
                        Dim tmpUserIndex As Integer
                        tmpUserIndex = MapData(.Pos.Map, X, Y).Userindex
                        With UserList(tmpUserIndex)
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                Npclist(NpcIndex).PFINFO.Target.X = .Pos.Y
                                Npclist(NpcIndex).PFINFO.Target.Y = .Pos.X 'ops!
                                Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                                Call SeekPath(NpcIndex)
                                Exit Function
                            End If
                        End With
                    End If
                End If
            Next X
        Next Y
    End With
End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
    With UserList(Userindex)
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
    End With
    Dim K As Integer
    K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NpcIndex, Userindex, Npclist(NpcIndex).Spells(K))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
    Dim K As Integer
    K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(K))
End Sub

Public Sub SacerdoteHealUser(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_CURAR_SACERDOTE, .Pos.X, .Pos.Y))
        .Stats.MinHp = .Stats.MaxHp
        Call WriteUpdateHP(Userindex)
        Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado!!", FontTypeNames.FONTTYPE_INFO)
        If EsNewbie(Userindex) Then
            Call SacerdoteHealEffectsAndRestoreMana(Userindex)
        End If
        Call WriteUpdateUserStats(Userindex)
    End With
End Sub

Public Sub SacerdoteResucitateUser(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_RESUCITAR_SACERDOTE, .Pos.X, .Pos.Y))
        Call RevivirUsuario(Userindex)
        Call WriteConsoleMsg(Userindex, "Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
        If EsNewbie(Userindex) Then
            Call SacerdoteHealEffectsAndRestoreMana(Userindex)
        End If
    End With
End Sub

Private Sub SacerdoteHealEffectsAndRestoreMana(ByVal Userindex As Integer)
    Dim MensajeAyuda As String
    MensajeAyuda = "Cuando dejes de ser newbie no lo hara mas el sacerdote y deberas comprar pociones o curarte con hechizos"
    With UserList(Userindex)
        If .flags.Maldicion = 1 Then
            .flags.Maldicion = 0
            Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado de la maldicion.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
        End If
        If .flags.Ceguera = 1 Then
            .flags.Ceguera = 0
            Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado de la ceguera.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
        End If
        If .flags.Envenenado = 1 Then
            .flags.Envenenado = 0
            Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
        End If
        .Stats.MinMAN = .Stats.MaxMAN
        Call WriteUpdateMana(Userindex)
        Call WriteConsoleMsg(Userindex, "El sacerdote te ha restaurado el mana completamente.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub
