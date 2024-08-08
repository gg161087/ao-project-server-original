Attribute VB_Name = "Areas"
Option Explicit

Public Const XMaxMapSize        As Byte = 100
Public Const XMinMapSize        As Byte = 1
Public Const YMaxMapSize        As Byte = 100
Public Const YMinMapSize        As Byte = 1
Public Const XWindow            As Byte = 23
Public Const YWindow            As Byte = 19
Private Const TileBufferSize    As Byte = 5
Public Const RANGO_VISION_X     As Byte = XWindow \ 2
Public Const RANGO_VISION_Y     As Byte = YWindow \ 2
Public Const AREAS_X            As Byte = RANGO_VISION_X + TileBufferSize
Public Const AREAS_Y            As Byte = RANGO_VISION_Y + TileBufferSize
Private AreasIO As clsIniManager
Private FILE_AREAS As String
Private Const USER_NUEVO        As Byte = 255
Public ConnGroups()             As Collection

Public Type AreaInfo
    AreaPerteneceX              As Integer
    AreaPerteneceY              As Integer
End Type

Public Sub InitializeAreas()
    Dim i As Long
    ReDim ConnGroups(1 To NumMaps) As Collection
    For i = 1 To NumMaps
        Set ConnGroups(i) = New Collection
    Next i
End Sub

Public Sub AgregarUser(ByVal Userindex As Integer, ByVal Map As Integer, Optional ByVal ButIndex As Boolean = False)
    If Not MapaValido(Map) Then Exit Sub
    Dim EsNuevo As Boolean
        EsNuevo = True
    Dim i As Integer
    For i = 1 To ConnGroups(Map).Count()
        If ConnGroups(Map).Item(i) = Userindex Then
            EsNuevo = False
            Exit For
        End If
    Next i
    If EsNuevo Then
        Call ConnGroups(Map).Add(Userindex)
    End If
    With UserList(Userindex)
        .AreasInfo.AreaPerteneceX = -1
        .AreasInfo.AreaPerteneceY = -1
    End With
    Call CheckUpdateNeededUser(Userindex, USER_NUEVO, ButIndex)
End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        .AreasInfo.AreaPerteneceX = -1
        .AreasInfo.AreaPerteneceY = -1
    End With
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub

Public Sub QuitarUser(ByVal Userindex As Integer, ByVal Map As Integer)
    Dim LoopA As Long
    For LoopA = 1 To ConnGroups(Map).Count()
        If ConnGroups(Map).Item(LoopA) = Userindex Then
            Call ConnGroups(Map).Remove(LoopA)
            Exit For
        End If
    Next LoopA
End Sub

Public Sub CheckUpdateNeededUser(ByVal Userindex As Integer, ByVal heading As Byte, Optional ByVal ButIndex As Boolean = False, Optional verInvis As Byte = 0)
    With UserList(Userindex)
        If .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X And _
           .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y Then _
                Exit Sub
        Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, CurUser As Long, Map As Long
        Call CalcularNuevaArea(.Pos.X, .Pos.Y, heading, MinX, MaxX, MinY, MaxY)
        Call WriteAreaChanged(Userindex)
        Map = .Pos.Map
        For X = MinX To MaxX
            For Y = MinY To MaxY
                If MapData(Map, X, Y).Userindex Then
                    CurUser = MapData(Map, X, Y).Userindex
                    If Userindex <> CurUser Then
                        If Not (UserList(CurUser).flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, Userindex, CurUser, Map, X, Y)
                            If UserList(CurUser).flags.Navegando = 0 Then
                                If UserList(CurUser).flags.invisible Or UserList(CurUser).flags.Oculto Then
                                    If UserList(Userindex).flags.Privilegios And PlayerType.User Then
                                        Call WriteSetInvisible(Userindex, UserList(CurUser).Char.CharIndex, True)
                                    End If
                                End If
                            End If
                        End If
                        If Not (.flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, CurUser, Userindex, .Pos.Map, .Pos.X, .Pos.Y)
                            If .flags.Navegando = 0 Then
                                If .flags.invisible Or .flags.Oculto Then
                                    If UserList(CurUser).flags.Privilegios And PlayerType.User Then
                                        Call WriteSetInvisible(CurUser, .Char.CharIndex, True)
                                    End If
                                End If
                            Else
                                Call WriteConsoleMsg(CurUser, "No podes hacerte invisible navegando.", FONTTYPE_INFO)
                            End If
                        End If
                    ElseIf heading = USER_NUEVO And Not ButIndex Then
                        Call MakeUserChar(False, Userindex, Userindex, Map, X, Y)
                        If .flags.AdminInvisible = 1 Or .flags.Navegando = 0 And (.flags.invisible Or .flags.Oculto) Then
                            Call WriteSetInvisible(Userindex, .Char.CharIndex, True)
                        End If
                    End If
                End If
                If MapData(Map, X, Y).NpcIndex Then
                    Call MakeNPCChar(False, Userindex, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                End If
                If MapData(Map, X, Y).ObjInfo.ObjIndex Then
                    CurUser = MapData(Map, X, Y).ObjInfo.ObjIndex
                    If Not EsObjetoFijo(ObjData(CurUser).OBJType) Then
                        Call WriteObjectCreate(Userindex, ObjData(CurUser).GrhIndex, X, Y)

                        If ObjData(CurUser).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, Userindex, X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(False, Userindex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                        End If
                    End If
                End If
            Next Y
        Next X
        .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X
        .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal heading As Byte)
    With Npclist(NpcIndex)
        If .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X And _
           .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y Then _
                Exit Sub
        Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer, Userindex As Long
        Call CalcularNuevaArea(.Pos.X, .Pos.Y, heading, MinX, MaxX, MinY, MaxY)
        If MapInfo(.Pos.Map).NumUsers <> 0 Then
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    If MapData(.Pos.Map, X, Y).Userindex Then _
                        Call MakeNPCChar(False, MapData(.Pos.Map, X, Y).Userindex, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)
                Next Y
            Next X
        End If
        .AreasInfo.AreaPerteneceX = .Pos.X \ AREAS_X
        .AreasInfo.AreaPerteneceY = .Pos.Y \ AREAS_Y
    End With
End Sub

Private Sub CalcularNuevaArea(ByVal X As Integer, ByVal Y As Integer, ByVal heading As Byte, ByRef MinX As Integer, ByRef MaxX As Integer, ByRef MinY As Integer, ByRef MaxY As Integer)
    Dim AreaX As Integer, AreaY As Integer
    Dim MinAreaX As Integer, MaxAreaX As Integer, MinAreaY As Integer, MaxAreaY As Integer
    AreaX = X \ AREAS_X
    AreaY = Y \ AREAS_Y
    Select Case heading
        Case eHeading.NORTH
            MinAreaX = AreaX - 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY - 1

        Case eHeading.EAST
            MinAreaX = AreaX + 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY + 1

        Case eHeading.SOUTH
            MinAreaX = AreaX - 1
            MinAreaY = AreaY + 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY + 1

        Case eHeading.WEST
            MinAreaX = AreaX - 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX - 1
            MaxAreaY = AreaY + 1

        Case Else
            MinAreaX = AreaX - 1
            MinAreaY = AreaY - 1
            MaxAreaX = AreaX + 1
            MaxAreaY = AreaY + 1
    End Select
    MinX = MinAreaX * AREAS_X
    MinY = MinAreaY * AREAS_Y
    MaxX = (MaxAreaX + 1) * AREAS_X - 1
    MaxY = (MaxAreaY + 1) * AREAS_Y - 1
    If MinX < XMinMapSize Then MinX = XMinMapSize
    If MinY < YMinMapSize Then MinY = YMinMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
End Sub

Public Function EstanMismoArea(ByVal UserA As Integer, ByVal UserB As Integer) As Boolean
    EstanMismoArea = Abs(UserList(UserA).AreasInfo.AreaPerteneceX - UserList(UserB).AreasInfo.AreaPerteneceX) <= 1 And _
                     Abs(UserList(UserA).AreasInfo.AreaPerteneceY - UserList(UserB).AreasInfo.AreaPerteneceY) <= 1
End Function

Public Function EstanMismoAreaNPC(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Boolean
    EstanMismoAreaNPC = Abs(UserList(Userindex).AreasInfo.AreaPerteneceX - Npclist(NpcIndex).AreasInfo.AreaPerteneceX) <= 1 And _
                        Abs(UserList(Userindex).AreasInfo.AreaPerteneceY - Npclist(NpcIndex).AreasInfo.AreaPerteneceY) <= 1
End Function

Public Function EstanMismoAreaPos(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    EstanMismoAreaPos = Abs(UserList(Userindex).AreasInfo.AreaPerteneceX - X \ AREAS_X) <= 1 And _
                        Abs(UserList(Userindex).AreasInfo.AreaPerteneceY - Y \ AREAS_Y) <= 1
End Function
