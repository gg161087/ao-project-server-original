Attribute VB_Name = "Acciones"
Option Explicit

Sub Accion(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim tempIndex As Integer
    On Error Resume Next
    If (Abs(UserList(Userindex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(Userindex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    If InMapBounds(Map, X, Y) Then
        With UserList(Userindex)
            If MapData(Map, X, Y).NpcIndex > 0 Then
                tempIndex = MapData(Map, X, Y).NpcIndex
                .flags.TargetNPC = tempIndex
                If Npclist(tempIndex).Comercia = 1 Then
                    If .flags.Muerto = 1 Then
                        Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                        Exit Sub
                    End If
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call IniciarComercioNPC(Userindex)
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    If .flags.Muerto = 1 Then
                        Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                        Exit Sub
                    End If
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call IniciarDeposito(Userindex)
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex)) Then
                        Call SacerdoteResucitateUser(Userindex)
                    End If
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex) Then
                        Call SacerdoteHealUser(Userindex)
                    End If
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Artesano Then
                    If .flags.Muerto = 1 Then
                        Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                        Exit Sub
                    End If
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del artesano.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call WriteInitCraftsman(Userindex)
                End If
            ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas
                        Call AccionParaPuerta(Map, X, Y, Userindex)
                        
                    Case eOBJType.otCarteles
                        Call AccionParaCartel(Map, X, Y, Userindex)

                    Case eOBJType.otForos
                        Call AccionParaForo(Map, X, Y, Userindex)

                    Case eOBJType.otLena
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(Map, X, Y, Userindex)
                        End If
                End Select
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas
                        Call AccionParaPuerta(Map, X + 1, Y, Userindex)
                End Select
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas
                        Call AccionParaPuerta(Map, X + 1, Y + 1, Userindex)
                End Select
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas
                        Call AccionParaPuerta(Map, X, Y + 1, Userindex)
                End Select
            End If
        End With
    End If
End Sub

Public Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
    On Error Resume Next
    Dim Pos As WorldPos
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y
    If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If SendPosts(Userindex, ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).ForoID) Then
        Call WriteShowForumForm(Userindex)
    End If
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
    On Error Resume Next
    If Not (Distance(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, X, Y) > 2) Then
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
            If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    Call Bloquear(True, Map, X, Y, 0)
                    Call Bloquear(True, Map, X - 1, Y, 0)
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                Else
                    Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                Call Bloquear(True, Map, X - 1, Y, 1)
                Call Bloquear(True, Map, X, Y, 1)
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
            End If
            UserList(Userindex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        Else
            Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
    On Error Resume Next
    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
        If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
            Call WriteShowSignal(Userindex, MapData(Map, X, Y).ObjInfo.ObjIndex)
        End If
    End If
End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
    On Error Resume Next
    Dim Suerte             As Byte
    Dim exito              As Byte
    Dim obj                As obj
    Dim SkillSupervivencia As Byte
    Dim Pos                As WorldPos
    With Pos
        .Map = Map
        .X = X
        .Y = Y
    End With
    With UserList(Userindex)
        If Distancia(Pos, .Pos) > 2 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
            Call WriteConsoleMsg(Userindex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        SkillSupervivencia = .Stats.UserSkills(eSkill.Supervivencia)
        If SkillSupervivencia < 6 Then
            Suerte = 3
        ElseIf SkillSupervivencia <= 10 Then
            Suerte = 2
        Else
            Suerte = 1
        End If
        exito = RandomNumber(1, Suerte)
        If exito = 1 Then
            If MapInfo(.Pos.Map).Zona <> Ciudad Then
                With obj
                    .ObjIndex = FOGATA
                    .Amount = 1
                End With
                Call WriteConsoleMsg(Userindex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
                Call MakeObj(obj, Map, X, Y)
                Call mLimpieza.AgregarObjetoLimpieza(Pos)
                Call SubirSkill(Userindex, eSkill.Supervivencia, True)
            Else
                Call WriteConsoleMsg(Userindex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(Userindex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
            Call SubirSkill(Userindex, eSkill.Supervivencia, False)
        End If
    End With
End Sub

Public Sub AccionParaSacerdote(ByVal Userindex As Integer)
    With UserList(Userindex)
        If .flags.Muerto = 1 Then
            Call SacerdoteResucitateUser(Userindex)
        End If
        If .Stats.MinHp < .Stats.MaxHp Then
            Call SacerdoteHealUser(Userindex)
        End If
    End With
End Sub
