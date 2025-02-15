Attribute VB_Name = "Extra"
Option Explicit

Public Function EsNewbie(ByVal Userindex As Integer) As Boolean
    EsNewbie = UserList(Userindex).Stats.ELV <= LimiteNewbie
End Function

Public Function esArmada(ByVal Userindex As Integer) As Boolean
    esArmada = (UserList(Userindex).Faccion.ArmadaReal = 1)
End Function

Public Function esCaos(ByVal Userindex As Integer) As Boolean
    esCaos = (UserList(Userindex).Faccion.FuerzasCaos = 1)
End Function

Public Function EsGm(ByVal Userindex As Integer) As Boolean
    EsGm = (UserList(Userindex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function

Public Sub DoTileEvents(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
     On Error GoTo ErrorHandler
    Dim nPos       As WorldPos
    Dim FxFlag     As Boolean
    Dim TelepRadio As Integer
    Dim DestPos    As WorldPos
    If InMapBounds(Map, X, Y) Then
        With MapData(Map, X, Y)
            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
                TelepRadio = ObjData(.ObjInfo.ObjIndex).Radio
            End If
            If .TileExit.Map > 0 And .TileExit.Map <= NumMaps Then
                If FxFlag And TelepRadio > 0 Then
                    Dim attemps As Long
                    Dim exitMap As Boolean
                    Do
                        DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)
                        attemps = attemps + 1
                        exitMap = MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map > 0 And MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map <= NumMaps
                    Loop Until (attemps >= 5 Or exitMap = False)
                    If attemps >= 5 Then
                        DestPos.X = .TileExit.X
                        DestPos.Y = .TileExit.Y
                    End If
                Else
                    DestPos.X = .TileExit.X
                    DestPos.Y = .TileExit.Y
                End If
                DestPos.Map = .TileExit.Map
                If EsGm(Userindex) Then
                    Call LogGM(UserList(Userindex).Name, "Utilizo un teleport hacia el mapa " & DestPos.Map & " (" & DestPos.X & "," & DestPos.Y & ")")
                End If
                If MapInfo(DestPos.Map).OnDeathGoTo.Map <> 0 Then
                    If UserList(Userindex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "Solo se permite entrar al mapa a los personajes vivos.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                        Exit Sub
                    End If
                End If
                If MapInfo(DestPos.Map).Restringir = eRestrict.restrict_newbie Then
                    If EsNewbie(Userindex) Or EsGm(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, False)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_armada Then
                    If esArmada(Userindex) Or EsGm(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros del ejercito real.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_caos Then
                    If esCaos(Userindex) Or EsGm(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "Mapa exclusivo para miembros de la legion oscura.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_faccion Then
                    If esArmada(Userindex) Or esCaos(Userindex) Or EsGm(Userindex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                            Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "Solo se permite entrar al mapa si eres miembro de alguna faccion.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(Userindex).Pos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                Else
                    If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(Userindex)) Then
                        Call WarpUserChar(Userindex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(DestPos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(Userindex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                End If
                Dim aN As Integer
                aN = UserList(Userindex).flags.AtacadoPorNpc
                If aN > 0 Then
                    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                    Npclist(aN).flags.AttackedBy = vbNullString
                End If
                aN = UserList(Userindex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(Userindex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                UserList(Userindex).flags.AtacadoPorNpc = 0
                UserList(Userindex).flags.NPCAtacado = 0
            End If
        End With
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)
End Sub

Function InRangoVision(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X > UserList(Userindex).Pos.X - MinXBorder And X < UserList(Userindex).Pos.X + MinXBorder Then
        If Y > UserList(Userindex).Pos.Y - MinYBorder And Y < UserList(Userindex).Pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function
        End If
    End If
    InRangoVision = False
End Function

Public Function InVisionRangeAndMap(ByVal Userindex As Integer, ByRef OtherUserPos As WorldPos) As Boolean
    With UserList(Userindex)
        If .Pos.Map <> OtherUserPos.Map Then Exit Function
        If OtherUserPos.X < .Pos.X - MinXBorder Or OtherUserPos.X > .Pos.X + MinXBorder Then Exit Function
        If OtherUserPos.Y < .Pos.Y - MinYBorder And OtherUserPos.Y > .Pos.Y + MinYBorder Then Exit Function
    End With
    InVisionRangeAndMap = True
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean
    If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
        If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function
        End If
    End If
    InRangoVisionNPC = False
End Function

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True
    End If
End Function

Private Function RhombLegalPos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, ByVal Distance As Long, Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False) As Boolean
    Dim i As Long
    vX = Pos.X - Distance
    vY = Pos.Y
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    vX = Pos.X
    vY = Pos.Y - Distance
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    vX = Pos.X + Distance
    vY = Pos.Y
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    vX = Pos.X
    vY = Pos.Y + Distance
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    RhombLegalPos = False
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, ByVal Distance As Long, ByVal ObjIndex As Integer, ByVal ObjAmount As Long, ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim HayObj      As Boolean
    Dim X           As Integer
    Dim Y           As Integer
    Dim MapObjIndex As Integer
    vX = Pos.X - Distance
    vY = Pos.Y
    For i = 0 To Distance - 1
        X = vX + i
        Y = vY - i
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    vX = Pos.X
    vY = Pos.Y - Distance
    For i = 0 To Distance - 1
        X = vX + i
        Y = vY + i
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    vX = Pos.X + Distance
    vY = Pos.Y
    For i = 0 To Distance - 1
        X = vX - i
        Y = vY + i
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    vX = Pos.X
    vY = Pos.Y + Distance
    For i = 0 To Distance - 1
        X = vX - i
        Y = vY - i
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    RhombLegalTilePos = False
    Exit Function
ErrorHandler:
    Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.description)
End Function

Public Function HayObjeto(ByVal Mapa As Integer, ByVal X As Long, ByVal Y As Long, ByVal ObjIndex As Integer, ByVal ObjAmount As Long) As Boolean
    Dim MapObjIndex As Integer
    MapObjIndex = MapData(Mapa, X, Y).ObjInfo.ObjIndex
    If MapObjIndex <> 0 Then
        If MapObjIndex = ObjIndex Then
            HayObjeto = (MapData(Mapa, X, Y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
        Else
            HayObjeto = True
        End If
    Else
        HayObjeto = False
    End If
End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False)
    Dim Found As Boolean
    Dim LoopC As Integer
    Dim tX    As Long
    Dim tY    As Long
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    LoopC = 1
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, CheckExitTile) Then
        Found = True
    Else
        While (Not Found) And LoopC <= 12
            If RhombLegalPos(Pos, tX, tY, LoopC, PuedeAgua, PuedeTierra, CheckExitTile) Then
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
End Sub

Public Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
    Call ClosestLegalPos(Pos, nPos, , , True)
End Sub

Function NameIndex(ByVal Name As String) As Integer
    Dim Userindex As Long
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    Name = UCase$(Name)
    If InStrB(Name, "+") <> 0 Then
        Name = Replace(Name, "+", " ")
    End If
    Userindex = 1
     Do Until StrComp(UCase$(UserList(Userindex).Name), Name) = 0
        Userindex = Userindex + 1
        If Userindex > MaxUsers Then
            NameIndex = 0
            Exit Function
        End If
    Loop
    NameIndex = Userindex
End Function

Function CheckForSameIP(ByVal Userindex As Integer, ByVal UserIP As String) As Boolean
    Dim LoopC As Long
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).IP = UserIP And Userindex <> LoopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC
    CheckForSameIP = False
End Function

Function CheckForSameName(ByVal Name As String) As Boolean
    Dim LoopC As Long
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged Then
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                Exit Function
            End If
        End If
    Next LoopC
    CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
    Select Case Head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1
    End Select
End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False) As Boolean
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPos = False
    Else
        With MapData(Map, X, Y)
            If PuedeAgua And PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.Userindex = 0) And (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = (.Blocked <> 1) And (.Userindex = 0) And (.NpcIndex = 0) And (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.Userindex = 0) And (.NpcIndex = 0) And (HayAgua(Map, X, Y))
            Else
                LegalPos = False
            End If
        End With
        If CheckExitTile Then
            LegalPos = LegalPos And (MapData(Map, X, Y).TileExit.Map = 0)
        End If
    End If
End Function

Function MoveToLegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
    Dim Userindex        As Integer
    Dim IsDeadChar       As Boolean
    Dim IsAdminInvisible As Boolean
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        MoveToLegalPos = False
    Else
        With MapData(Map, X, Y)
            Userindex = .Userindex
            If Userindex > 0 Then
                IsDeadChar = (UserList(Userindex).flags.Muerto = 1)
                IsAdminInvisible = (UserList(Userindex).flags.AdminInvisible = 1)
            Else
                IsDeadChar = False
                IsAdminInvisible = False
            End If
            If PuedeAgua And PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                MoveToLegalPos = (.Blocked <> 1) And (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (HayAgua(Map, X, Y))
            Else
                MoveToLegalPos = False
            End If
        End With
    End If
End Function

Public Sub FindLegalPos(ByVal Userindex As Integer, ByVal Map As Integer, ByRef X As Integer, ByRef Y As Integer)
    If MapData(Map, X, Y).Userindex <> 0 Or MapData(Map, X, Y).NpcIndex <> 0 Then
        If MapData(Map, X, Y).Userindex = Userindex Then Exit Sub
        Dim FoundPlace     As Boolean
        Dim tX             As Long
        Dim tY             As Long
        Dim Rango          As Long
        Dim OtherUserIndex As Integer
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    If MapData(Map, tX, tY).Userindex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        Exit For
                    End If
                Next tX
                If FoundPlace Then Exit For
            Next tY
            If FoundPlace Then Exit For
        Next Rango
        If FoundPlace Then
            X = tX
            Y = tY
        Else
            OtherUserIndex = MapData(Map, X, Y).Userindex
            If OtherUserIndex <> 0 Then
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                    End If
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")
                    End If
                End If
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If
End Sub

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
    Dim IsDeadChar       As Boolean
    Dim Userindex        As Integer
    Dim IsAdminInvisible As Boolean
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If
    With MapData(Map, X, Y)
        Userindex = .Userindex
        If Userindex > 0 Then
            IsDeadChar = UserList(Userindex).flags.Muerto = 1
            IsAdminInvisible = (UserList(Userindex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
        If AguaValida = 0 Then
            LegalPosNPC = (.Blocked <> 1) And (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.POSINVALIDA Or IsPet) And Not HayAgua(Map, X, Y)
        Else
            LegalPosNPC = (.Blocked <> 1) And (.Userindex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.POSINVALIDA Or IsPet)
        End If
    End With
End Function

Sub SendHelp(ByVal index As Integer)
    Dim NumHelpLines As Integer
    Dim LoopC        As Integer
    NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))
    For LoopC = 1 To NumHelpLines
        Call WriteConsoleMsg(index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
    Next LoopC
End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo ErrorHandler
    Dim FoundChar      As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex  As Integer
    Dim Stat           As String
    Dim ft             As FontTypeNames
    With UserList(Userindex)
        If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then
            Exit Sub
        End If
        If InMapBounds(Map, X, Y) Then
            With .flags
                .TargetMap = Map
                .TargetX = X
                .TargetY = Y
                If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y
                    FoundSomething = 1
                ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                    If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y
                        FoundSomething = 1
                    End If
                ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                    If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y + 1
                        FoundSomething = 1
                    End If
                ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                    If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        .TargetObjMap = Map
                        .TargetObjX = X
                        .TargetObjY = Y + 1
                        FoundSomething = 1
                    End If
                End If
                If FoundSomething = 1 Then
                    .TargetObj = MapData(Map, .TargetObjX, .TargetObjY).ObjInfo.ObjIndex
                    If MostrarCantidad(.TargetObj) Then
                        Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name & " - " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(Userindex, ObjData(.TargetObj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                If Y + 1 <= YMaxMapSize Then
                    If MapData(Map, X, Y + 1).Userindex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).Userindex
                        FoundChar = 1
                    End If
                    If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
                        FoundChar = 2
                    End If
                End If
                If FoundChar = 0 Then
                    If MapData(Map, X, Y).Userindex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).Userindex
                        FoundChar = 1
                    End If
                    If MapData(Map, X, Y).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).NpcIndex
                        FoundChar = 2
                    End If
                End If
            End With
            If FoundChar = 1 Then
                If UserList(TempCharIndex).flags.AdminInvisible = 0 Or .flags.Privilegios And PlayerType.Dios Then
                    With UserList(TempCharIndex)
                        If LenB(.DescRM) = 0 And .showName Then
                            If EsNewbie(TempCharIndex) Then
                                Stat = Stat & " <NEWBIE>"
                            End If
                            If .Faccion.ArmadaReal = 1 Then
                                Stat = Stat & " <Ejercito Real> " & "<" & TituloReal(TempCharIndex) & ">"
                            ElseIf .Faccion.FuerzasCaos = 1 Then
                                Stat = Stat & " <Legion Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                            End If
                            If .GuildIndex > 0 Then
                                Stat = Stat & " Clan: '" & modGuilds.GuildName(.GuildIndex) & "'"
                            End If
                            Stat = Stat & " Nivel: " & UserList(TempCharIndex).Stats.ELV
                            If Len(UserList(TempCharIndex).Desc) > 1 Then
                                Stat = UserList(TempCharIndex).Name & " - " & UserList(TempCharIndex).Desc & " (" & ListaClases(UserList(TempCharIndex).Clase) & " " & ListaRazas(UserList(TempCharIndex).raza) & Stat & "  " & " | "
                            Else
                                Stat = UserList(TempCharIndex).Name & " (" & ListaClases(UserList(TempCharIndex).Clase) & " " & ListaRazas(UserList(TempCharIndex).raza) & Stat & " " & " | "
                            End If
                            If UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.05) Then
                                Stat = Stat & " Muerto)"
                            ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.1) Then
                                Stat = Stat & " Casi muerto)"
                            ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.25) Then
                                Stat = Stat & " Muy Malherido)"
                            ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.5) Then
                                Stat = Stat & " Malherido)"
                            ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.75) Then
                                Stat = Stat & " Herido)"
                            ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp) Then
                                Stat = Stat & " Levemente Herido)"
                            Else
                                Stat = Stat & " Intacto)"
                            End If
                            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                                Stat = Stat & " [CONSEJO DE BANDERBILL]"
                                ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                                Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                                ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                            Else
                                If Not .flags.Privilegios And PlayerType.User Then
                                    Stat = Stat & " <GAME MASTER>"
                                    If .flags.Privilegios = PlayerType.Dios Then
                                        ft = FontTypeNames.FONTTYPE_DIOS
                                    ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                                        ft = FontTypeNames.FONTTYPE_GM
                                    ElseIf .flags.Privilegios = PlayerType.Consejero Then
                                        ft = FontTypeNames.FONTTYPE_CONSEJO
                                    ElseIf .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                                        ft = FontTypeNames.FONTTYPE_EJECUCION
                                    End If
                                ElseIf criminal(TempCharIndex) Then
                                    Stat = Stat & " <CRIMINAL>"
                                    ft = FontTypeNames.FONTTYPE_CRIMINAL
                                Else
                                    Stat = Stat & " <CIUDADANO>"
                                    ft = FontTypeNames.FONTTYPE_CITIZEN
                                End If
                            End If
                        End If
                    End With
                    If LenB(Stat) > 0 Then
                        Call WriteConsoleMsg(Userindex, Stat, ft)
                    End If
                    FoundSomething = 1
                    .flags.TargetUser = TempCharIndex
                    .flags.TargetNPC = 0
                    .flags.TargetNpcTipo = eNPCType.Comun
                Else
                    Stat = .DescRM
                    ft = FontTypeNames.FONTTYPE_INFOBOLD
                End If
            End If
            With .flags
                If FoundChar = 2 Then

                    Dim estatus            As String
                    Dim MinHp              As Long
                    Dim MaxHp              As Long
                    Dim SupervivenciaSkill As Byte
                    Dim sDesc              As String
                    Dim TimeParalizado     As String
                    MinHp = Npclist(TempCharIndex).Stats.MinHp
                    MaxHp = Npclist(TempCharIndex).Stats.MaxHp
                    SupervivenciaSkill = UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia)
                    If .Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                        estatus = "(" & MinHp & "/" & MaxHp & ") "
                    Else
                        If .Muerto = 0 Then
                            If SupervivenciaSkill <= 10 Then
                                estatus = "(Dudoso) "
                            ElseIf SupervivenciaSkill <= 20 Then
                                If MinHp < (MaxHp / 2) Then
                                    estatus = "(Herido) "
                                Else
                                    estatus = "(Sano) "
                                End If
                            ElseIf SupervivenciaSkill <= 30 Then
                                If MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Malherido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Herido) "
                                Else
                                    estatus = "(Sano) "
                                End If
                            ElseIf SupervivenciaSkill <= 40 Then
                                If MinHp < (MaxHp * 0.25) Then
                                    estatus = "(Muy malherido) "
                                ElseIf MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Herido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Levemente herido) "
                                Else
                                    estatus = "(Sano) "
                                End If
                            ElseIf SupervivenciaSkill < 60 Then
                                If MinHp < (MaxHp * 0.05) Then
                                    estatus = "(Agonizando) "
                                ElseIf MinHp < (MaxHp * 0.1) Then
                                    estatus = "(Casi muerto) "
                                ElseIf MinHp < (MaxHp * 0.25) Then
                                    estatus = "(Muy Malherido) "
                                ElseIf MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Herido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Levemente herido) "
                                ElseIf MinHp < (MaxHp) Then
                                    estatus = "(Sano) "
                                Else
                                    estatus = "(Intacto) "
                                End If
                            Else
                                estatus = "(" & MinHp & "/" & MaxHp & ") "
                            End If
                        End If
                    End If
                    If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) = 100 Then
                        If Npclist(TempCharIndex).flags.Paralizado = 1 Or Npclist(TempCharIndex).flags.Inmovilizado = 1 Then
                            TimeParalizado = " - Tiempo de paralisis: " & Npclist(TempCharIndex).Contadores.Paralisis & " segundos."
                        End If
                    End If
                    If Len(Npclist(TempCharIndex).Desc) > 1 Then
                        Stat = Npclist(TempCharIndex).Desc
                        If Npclist(TempCharIndex).NPCtype = eNPCType.Noble Then
                            If Npclist(TempCharIndex).flags.Faccion = 0 Then
                                If UserList(Userindex).Faccion.FuerzasCaos = 1 Then
                                    Stat = MENSAJE_REY_CAOS
                                    If .Privilegios And PlayerType.User Then
                                        If .Muerto = 0 Then Call UserDie(Userindex)
                                    End If
                                ElseIf criminal(Userindex) Then
                                    If UserList(Userindex).Faccion.CiudadanosMatados > 0 Or UserList(Userindex).Faccion.Reenlistadas > 4 Then 'Es criminal no enlistable.
                                        Stat = MENSAJE_REY_CRIMINAL_NOENLISTABLE
                                    Else
                                        Stat = MENSAJE_REY_CRIMINAL_ENLISTABLE
                                    End If
                                End If
                            Else
                                If UserList(Userindex).Faccion.ArmadaReal = 1 Then
                                    Stat = MENSAJE_DEMONIO_REAL
                                    If .Privilegios And PlayerType.User Then
                                        If .Muerto = 0 Then Call UserDie(Userindex)
                                    End If
                                ElseIf Not criminal(Userindex) Then
                                    If UserList(Userindex).Faccion.RecibioExpInicialReal = 1 Or UserList(Userindex).Faccion.Reenlistadas > 4 Then 'Es ciudadano no enlistable.
                                        Stat = MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE
                                    Else
                                        Stat = MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE
                                    End If
                                End If
                            End If
                        End If
                        Call WriteChatOverHead(Userindex, Stat, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                    Else
                        If Npclist(TempCharIndex).MaestroUser > 0 Then
                            Call WriteConsoleMsg(Userindex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & TimeParalizado, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(Userindex, estatus & Npclist(TempCharIndex).Name & TimeParalizado, FontTypeNames.FONTTYPE_INFO)
                            If Len(Npclist(TempCharIndex).flags.AttackedFirstBy) > 0 And (UserList(Userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                                Call WriteConsoleMsg(Userindex, "Le pego primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                    FoundSomething = 1
                    .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                    .TargetNPC = TempCharIndex
                    .TargetUser = 0
                    .TargetObj = 0
                End If
                If FoundChar = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                End If
                If FoundSomething = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0
                End If
            End With
        Else
            If FoundSomething = 0 Then
                With .flags
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0
                End With
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.description)
End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
    Dim X As Integer
    Dim Y As Integer
    X = Pos.X - Target.X
    Y = Pos.Y - Target.Y
    If Sgn(X) = -1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
        Exit Function
    End If
    If Sgn(X) = 1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
        Exit Function
    End If
    If Sgn(X) = 1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
        Exit Function
    End If
    If Sgn(X) = -1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
        Exit Function
    End If
    If Sgn(X) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
        Exit Function
    End If
    If Sgn(X) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
        Exit Function
    End If
    If Sgn(X) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST
        Exit Function
    End If
    If Sgn(X) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST
        Exit Function
    End If
    If Sgn(X) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0
        Exit Function
    End If
End Function

Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean
    With ObjData(index)
        ItemNoEsDeMapa = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And .OBJType <> eOBJType.otTeleport
    End With
End Function

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
    With ObjData(index)
        MostrarCantidad = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And .OBJType <> eOBJType.otTeleport
    End With
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean
    EsObjetoFijo = OBJType = eOBJType.otForos Or OBJType = eOBJType.otCarteles Or OBJType = eOBJType.otArboles Or OBJType = eOBJType.otYacimiento
End Function

Public Function RestrictStringToByte(ByRef restrict As String) As Byte
    restrict = UCase$(restrict)
    Select Case restrict
        Case "NEWBIE"
            RestrictStringToByte = 1
        
        Case "ARMADA"
            RestrictStringToByte = 2
        
        Case "CAOS"
            RestrictStringToByte = 3
        
        Case "FACCION"
            RestrictStringToByte = 4
        
        Case Else
            RestrictStringToByte = 0
    End Select
End Function

Public Function RestrictByteToString(ByVal restrict As Byte) As String
    Select Case restrict
        Case 1
            RestrictByteToString = "NEWBIE"
        
        Case 2
            RestrictByteToString = "ARMADA"
        
        Case 3
            RestrictByteToString = "CAOS"
        
        Case 4
            RestrictByteToString = "FACCION"
        
        Case 0
            RestrictByteToString = "NO"
    End Select
End Function

Public Function TerrainStringToByte(ByRef restrict As String) As Byte
    restrict = UCase$(restrict)
    Select Case restrict
        Case "NIEVE"
            TerrainStringToByte = 1
        
        Case "DESIERTO"
            TerrainStringToByte = 2
        
        Case "CIUDAD"
            TerrainStringToByte = 3
        
        Case "CAMPO"
            TerrainStringToByte = 4
        
        Case "DUNGEON"
            TerrainStringToByte = 5
        
        Case Else
            TerrainStringToByte = 0
    End Select
End Function

Public Function TerrainByteToString(ByVal restrict As Byte) As String
    Select Case restrict
        Case 1
            TerrainByteToString = "NIEVE"
        
        Case 2
            TerrainByteToString = "DESIERTO"
        
        Case 3
            TerrainByteToString = "CIUDAD"
        
        Case 4
            TerrainByteToString = "CAMPO"
        
        Case 5
            TerrainByteToString = "DUNGEON"
        
        Case 0
            TerrainByteToString = "BOSQUE"
    End Select
End Function
