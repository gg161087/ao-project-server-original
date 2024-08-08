Attribute VB_Name = "Retos"
Option Explicit

Private Const MAX_RETOS_SIMULTANEOS As Byte = 4
Public Arenas(1 To MAX_RETOS_SIMULTANEOS) As tMapEvent
Public Retos(1 To MAX_RETOS_SIMULTANEOS) As tRetos

Public Enum eTipoReto
    None = 0
    FightOne = 1
    FightTwo = 2
    FightThree = 3
End Enum

Public Type tRetoUser
    Userindex As Integer
    Team As Byte
    Rounds As Byte
End Type

Private Type tMapEvent
    Map As Integer
    X As Byte
    Y As Byte
    X2 As Byte
    Y2 As Byte
End Type

Private Type tRetos
    Run As Boolean
    Users() As tRetoUser
    RequiredGld As Long
End Type

Public Sub LoadArenas()
    Dim i       As Long
    Dim RetosIO As clsIniManager
    Set RetosIO = New clsIniManager
    Call RetosIO.Initialize(DatPath & "Retos.dat")
    For i = LBound(Arenas) To UBound(Arenas)
        Arenas(i).Map = RetosIO.GetValue("ARENA" & CStr(i), "Mapa")
        Arenas(i).X = RetosIO.GetValue("ARENA" & CStr(i), "X")
        Arenas(i).X2 = RetosIO.GetValue("ARENA" & CStr(i), "X2")
        Arenas(i).Y = RetosIO.GetValue("ARENA" & CStr(i), "Y")
        Arenas(i).Y2 = RetosIO.GetValue("ARENA" & CStr(i), "Y2")
    Next
    Set RetosIO = Nothing
End Sub

Private Sub ResetDueloUser(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
        With UserList(Userindex)
            If .Counters.TimeFight > 0 Then
                .Counters.TimeFight = 0
                Call WriteUserInEvent(Userindex)
            End If
            With Retos(.flags.SlotReto)
                .Users(UserList(Userindex).flags.SlotRetoUser).Userindex = 0
                .Users(UserList(Userindex).flags.SlotRetoUser).Team = 0
                .Users(UserList(Userindex).flags.SlotRetoUser).Rounds = 0
            End With
            .flags.SlotReto = 0
            .flags.SlotRetoUser = 255
            Call StatsDuelos(Userindex)
            Call WarpPosAnt(Userindex)
        End With
    Exit Sub
ErrorHandler:

End Sub

Private Sub ResetDuelo(ByVal SlotReto As Byte)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Userindex > 0 Then
                ResetDueloUser .Users(LoopC).Userindex
            End If
            .Users(LoopC).Userindex = 0
            .Users(LoopC).Rounds = 0
            .Users(LoopC).Team = 0
        Next LoopC
        .RequiredGld = 0
        .Run = False
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ResetDuelo()"
End Sub

Private Function FreeSlotArena() As Byte
    Dim LoopC As Integer
    For LoopC = 1 To MAX_RETOS_SIMULTANEOS
        If Retos(LoopC).Run = False Then
            FreeSlotArena = LoopC
            Exit Function
        End If
    Next LoopC
End Function

Private Function FreeSlot() As Byte
    Dim LoopC As Integer
    FreeSlot = 0
    For LoopC = 1 To MAX_RETOS_SIMULTANEOS
        With Retos(LoopC)
            If .Run = False Then
            FreeSlot = LoopC
            Exit For
            End If
        End With
    Next LoopC
End Function

Private Sub PasateInteger(ByVal SlotArena As Byte, ByRef Users() As String)
    On Error GoTo ErrorHandler
    With Retos(SlotArena)
        Dim LoopC As Integer
        ReDim .Users(LBound(Users()) To UBound(Users())) As tRetoUser
        For LoopC = LBound(.Users()) To UBound(.Users())
            .Users(LoopC).Userindex = NameIndex(Users(LoopC))
            If .Users(LoopC).Userindex > 0 Then
                UserList(.Users(LoopC).Userindex).Stats.Gld = UserList(.Users(LoopC).Userindex).Stats.Gld - .RequiredGld
                Call WriteUpdateGold(.Users(LoopC).Userindex)
            End If
        Next LoopC
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : PasateInteger()"
End Sub

Private Sub RewardUsers(ByVal SlotReto As Byte, ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim obj As obj
    With UserList(Userindex)
        .Stats.Gld = .Stats.Gld + (Retos(SlotReto).RequiredGld * 2)
        Call WriteUpdateGold(Userindex)
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : RewardUsers()"
End Sub

Private Function SetSubTipo(ByRef Users() As String) As eTipoReto
    On Error GoTo ErrorHandler
    If UBound(Users()) = 1 Then
        SetSubTipo = FightOne
        Exit Function
    End If
    If UBound(Users()) = 3 Then
        SetSubTipo = FightTwo
        Exit Function
    End If
    If UBound(Users()) = 5 Then
        SetSubTipo = FightThree
        Exit Function
    End If
    SetSubTipo = 0
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SetSubTipo()"
End Function

Private Function CanSetUsers(ByRef Users() As String) As Boolean
    On Error GoTo ErrorHandler
    Dim tUser As Integer
    Dim tmpUsers() As String
    Dim LoopC As Integer, loopX As Integer
    Dim Tmp As String
    If SetSubTipo(Users()) = 0 Then
        CanSetUsers = False
        Exit Function
    End If
    ReDim tmpUsers(LBound(Users()) To UBound(Users())) As String
    For LoopC = LBound(Users()) To UBound(Users())
        tmpUsers(LoopC) = Users(LoopC)
    Next LoopC
    For LoopC = LBound(Users()) To UBound(Users())
        For loopX = LBound(Users()) To UBound(Users()) - LoopC
            If Not loopX = UBound(Users()) Then
                If StrComp(UCase$(tmpUsers(loopX)), UCase$(tmpUsers(loopX + 1))) = 0 Then
                    CanSetUsers = False
                    Exit Function
                Else
                    Tmp = tmpUsers(loopX)
                    tmpUsers(loopX) = tmpUsers(loopX + 1)
                    tmpUsers(loopX + 1) = Tmp
                End If
            End If
        Next loopX
    Next LoopC
    CanSetUsers = True
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanSetUsers()"
End Function

Private Function CanContinueFight(ByVal Userindex As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim SlotReto As Byte
    Dim SlotRetoUser As Byte
    SlotReto = UserList(Userindex).flags.SlotReto
    SlotRetoUser = UserList(Userindex).flags.SlotRetoUser
    CanContinueFight = False
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Userindex > 0 And .Users(LoopC).Userindex <> Userindex Then
                If .Users(SlotRetoUser).Team = .Users(LoopC).Team Then
                    With UserList(.Users(LoopC).Userindex)
                        If .flags.Muerto = 0 Then
                            CanContinueFight = True
                            Exit Function
                        End If
                    End With
                End If
            End If
        Next LoopC
    End With
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanContinueFight()"
End Function

Private Function AttackerFight(ByVal SlotReto As Byte, ByVal TeamUser As Byte) As Integer
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Userindex > 0 Then
                If .Users(LoopC).Team > 0 And .Users(LoopC).Team <> TeamUser Then
                    AttackerFight = .Users(LoopC).Userindex
                    Exit For
                End If
            End If
        Next LoopC
    End With
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AttackerFight()"
End Function

Private Function CanAcceptFight(ByVal Userindex As Integer, ByVal UserName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim SlotTemp As Byte
    Dim tUser As Integer
    Dim ArrayNulo As Long
    tUser = NameIndex(UserName)
    If tUser <= 0 Then
        CanAcceptFight = False
        Exit Function
    End If
    With UserList(tUser)
        SlotTemp = SearchFight(UCase$(UserList(Userindex).Name), .RetoTemp.Users, .RetoTemp.Accepts)
        If SlotTemp = 255 Then
            CanAcceptFight = False
            Exit Function
        End If
        If .RetoTemp.Accepts(SlotTemp) = 1 Then
            CanAcceptFight = False
            Exit Function
        End If
        .RetoTemp.Accepts(SlotTemp) = 1
        CanAcceptFight = True
        If CheckAccepts(.RetoTemp.Accepts) Then
            GoFight tUser
        End If
    End With
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanAcceptFight()"
End Function

Private Function ValidateFight_Users(ByVal Userindex As Integer, ByVal GldRequired As Long, ByRef Users() As String) As Boolean
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim tUser As Integer
    For LoopC = LBound(Users()) To UBound(Users())
        If Users(LoopC) <> vbNullString Then
            tUser = NameIndex(Users(LoopC))
            If tUser > 0 Then
                If EsGm(tUser) Then
                End If
            End If
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "El personaje " & Users(LoopC) & " esta offline", FontTypeNames.FONTTYPE_INFO)
                ValidateFight_Users = False
                Exit Function
            End If
            With UserList(tUser)
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(Userindex, "El personaje " & Users(LoopC) & " esta muerto.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
                If MapInfo(.Pos.Map).Pk = True Then
                    ValidateFight_Users = False
                    Exit Function
                End If
                If (.flags.SlotReto > 0) Then
                    Call WriteConsoleMsg(Userindex, "El personaje " & Users(LoopC) & " esta participando en otro evento.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
                If .flags.Comerciando Then
                    Call WriteConsoleMsg(Userindex, "El personaje " & Users(LoopC) & " no esta disponible en este momento.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
                If .Stats.Gld < GldRequired Then
                    Call WriteConsoleMsg(Userindex, "El personaje " & .Name & " no tiene las monedas en su billetera.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
            End With
        End If
    Next LoopC
    ValidateFight_Users = True
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ValidateFight_Users()"
End Function

Private Function ValidateFight(ByVal Userindex As Integer, ByVal GldRequired As Long, ByRef Users() As String) As Boolean
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim tUser As Integer
    If GldRequired < 0 Or GldRequired > 100000000 Then
        Call WriteConsoleMsg(Userindex, "Oro Minimo: 0 . Oro Maximo 100.000.000", FontTypeNames.FONTTYPE_INFO)
        ValidateFight = False
        Exit Function
    End If
    If Not CanSetUsers(Users) Then
        Call LogRetos("POSIBLE HACKEO: " & UserList(Userindex).Name & " hackeo el sistema de retos.")
        ValidateFight = False
        Exit Function
    End If
    If Not ValidateFight_Users(Userindex, GldRequired, Users()) Then
        ValidateFight = False
        Exit Function
    End If
    ValidateFight = True
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ValidateFight()"
End Function

Private Function StrTeam(ByRef Users() As tRetoUser) As String
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim strtemp(1) As String
    If UBound(Users()) = 1 Then
        If Users(0).Userindex > 0 Then
            strtemp(0) = UserList(Users(0).Userindex).Name
        Else
            strtemp(0) = "Usuario descalificado"
        End If
        If Users(1).Userindex > 0 Then
            strtemp(1) = UserList(Users(1).Userindex).Name
        Else
            strtemp(1) = "Usuario descalificado"
        End If
        StrTeam = strtemp(0) & " vs " & strtemp(1)
        Exit Function
    End If
    For LoopC = LBound(Users()) To UBound(Users())
        If Users(LoopC).Userindex > 0 Then
            If LoopC < ((1 + UBound(Users)) / 2) Then
                strtemp(0) = strtemp(0) & UserList(Users(LoopC).Userindex).Name & ", "
            Else
                strtemp(1) = strtemp(1) & UserList(Users(LoopC).Userindex).Name & ", "
            End If
        End If
    Next LoopC
    If Not strtemp(0) = vbNullString Then
        strtemp(0) = Left$(strtemp(0), Len(strtemp(0)) - 2)
    Else
        strtemp(0) = "Equipo descalificado"
    End If
    If Not strtemp(1) = vbNullString Then
        strtemp(1) = Left$(strtemp(1), Len(strtemp(1)) - 2)
    Else
        strtemp(1) = "Equipo descalificado"
    End If
    StrTeam = strtemp(0) & " vs " & strtemp(1)
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : StrTeam()"
End Function

Private Function CheckAccepts(ByRef Accepts() As Byte) As Boolean
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    CheckAccepts = True
    For LoopC = LBound(Accepts()) To UBound(Accepts())
        If Accepts(LoopC) = 0 Then
            CheckAccepts = False
            Exit Function
        End If
    Next LoopC
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CheckAccepts()"
End Function

Private Function SearchFight(ByVal UserName As String, ByRef Users() As String, ByRef Accepts() As Byte) As Byte
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    SearchFight = 255
    For LoopC = LBound(Users()) To UBound(Users())
        If StrComp(UCase$(Users(LoopC)), UCase$(UserName)) = 0 And Accepts(LoopC) = 0 Then
            SearchFight = LoopC
            Exit Function
        End If
    Next LoopC
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SearchFight()"
End Function

Public Function CanAttackReto(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    CanAttackReto = True
    With UserList(AttackerIndex)
        If .flags.SlotReto > 0 Then
            CanAttackReto = True
            Exit Function
        End If
    End With
    Exit Function
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanAttackReto()"
End Function

Private Sub SendInvitation(ByVal Userindex As Integer, ByVal GldRequired As Long, ByRef Users() As String)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim strtemp As String
    Dim tUser As Integer
    Dim Str() As tRetoUser
    With UserList(Userindex)
        With .RetoTemp
            ReDim .Accepts(LBound(Users()) To UBound(Users())) As Byte
            ReDim .Users(LBound(Users()) To UBound(Users())) As String
            .RequiredGld = GldRequired
            .Users = Users
            .Accepts(UBound(Users())) = 1
        End With
    End With
    ReDim Str(LBound(Users()) To UBound(Users())) As tRetoUser
    For LoopC = LBound(Users()) To UBound(Users())
        Str(LoopC).Userindex = NameIndex(Users(LoopC))
    Next LoopC
    strtemp = StrTeam(Str) & "."
    strtemp = strtemp & IIf(GldRequired > 0, " Oro requerido: " & GldRequired & ".", vbNullString)
    strtemp = strtemp & " Para aceptar tipea /ACEPTAR " & UserList(Userindex).Name
    For LoopC = LBound(Users()) To UBound(Users())
        tUser = NameIndex(Users(LoopC))
        If tUser <> Userindex Then
            Call WriteConsoleMsg(tUser, strtemp, FontTypeNames.FONTTYPE_INFO)
        End If
    Next LoopC
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendInvitation()"
End Sub

Private Sub GoFight(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim GldRequired As Long
    Dim SlotArena As Byte
    SlotArena = FreeSlotArena
    If SlotArena = 0 Then
        Exit Sub
    End If
    With UserList(Userindex)
        If ValidateFight(Userindex, .RetoTemp.RequiredGld, .RetoTemp.Users) Then
            Retos(SlotArena).RequiredGld = .RetoTemp.RequiredGld
            Retos(SlotArena).Run = True
            Call PasateInteger(SlotArena, .RetoTemp.Users)
            Call SetUserEvent(SlotArena, Retos(SlotArena).Users)
            Call WarpFight(Retos(SlotArena).Users)
        End If
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : GoFight()"
End Sub

Private Sub SetUserEvent(ByVal SlotReto As Byte, ByRef Users() As tRetoUser)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim SlotRetoUser As Byte
    For LoopC = LBound(Users()) To UBound(Users())
        If Users(LoopC).Userindex > 0 Then
            With Users(LoopC)
                If .Userindex > 0 Then
                    UserList(.Userindex).flags.SlotReto = SlotReto
                    UserList(.Userindex).flags.SlotRetoUser = LoopC
                End If
            End With
            With Retos(SlotReto)
                If LoopC < ((1 + UBound(Users())) / 2) Then
                    .Users(LoopC).Team = 2
                Else
                    .Users(LoopC).Team = 1
                End If
            End With
            With UserList(Users(LoopC).Userindex)
                .PosAnt.Map = .Pos.Map
                .PosAnt.X = .Pos.X
                .PosAnt.Y = .Pos.Y
            End With
        End If
    Next LoopC
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SetUserEvent()"
End Sub

Private Sub WarpFight(ByRef Users() As tRetoUser)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim tUser As Integer
    Dim Pos As WorldPos
    Const Tile_Extra As Byte = 5
    For LoopC = LBound(Users()) To UBound(Users())
        tUser = Users(LoopC).Userindex
        If tUser > 0 Then
            Pos.Map = Arenas(UserList(tUser).flags.SlotReto).Map
            If Users(LoopC).Team = 1 Then
                Pos.X = Arenas(UserList(tUser).flags.SlotReto).X
                Pos.Y = Arenas(UserList(tUser).flags.SlotReto).Y
            Else
                Pos.X = Arenas(UserList(tUser).flags.SlotReto).X2
                Pos.Y = Arenas(UserList(tUser).flags.SlotReto).Y2
            End If
            With UserList(tUser)
                .Counters.TimeFight = 10
                Call WriteUserInEvent(tUser)
                Call ClosestStablePos(Pos, Pos)
                Call WarpUserChar(tUser, Pos.Map, Pos.X, Pos.Y, False)
            End With
        End If
    Next LoopC
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : WarpFight()"
End Sub

Private Sub AddRound(ByVal SlotReto As Byte, ByVal Team As Byte)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Team = Team And .Users(LoopC).Userindex > 0 Then
                .Users(LoopC).Rounds = .Users(LoopC).Rounds + 1
            End If
        Next LoopC
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AddRound()"
End Sub

Private Sub SendMsjUsers(ByVal strMsj As String, ByRef Users() As String)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim tUser As Integer
    For LoopC = LBound(Users()) To UBound(Users())
        tUser = NameIndex(Users(LoopC))
        If tUser > 0 Then
            Call WriteConsoleMsg(tUser, strMsj, FontTypeNames.FONTTYPE_VENENO)
        End If
    Next LoopC
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendMsjUsers()"
End Sub

Private Function ExistCompanero(ByVal Userindex As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim SlotReto As Byte
    Dim SlotRetoUser As Byte
    SlotReto = UserList(Userindex).flags.SlotReto
    SlotRetoUser = UserList(Userindex).flags.SlotRetoUser
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Userindex > 0 And LoopC <> SlotRetoUser Then
                If .Users(LoopC).Team = .Users(SlotRetoUser).Team Then
                    ExistCompanero = True
                    Exit Function
                End If
            End If
        Next LoopC
    End With
    ExistCompanero = False
    Exit Function
ErrorHandler:
    LogRetos "Error " & Err.Number & " (" & Err.description & ") in procedure ExistCompanero"
End Function

Public Sub UserDieFight(ByVal Userindex As Integer, ByVal AttackerIndex As Integer, ByVal Forzado As Boolean)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim strtemp As String
    Dim SlotReto As Byte
    Dim TeamUser As Byte
    Dim Rounds As Byte
    Dim Deslogged As Boolean
    Dim ExistTeam As Boolean
    SlotReto = UserList(Userindex).flags.SlotReto
    Deslogged = False
    If AttackerIndex = 0 Then
        AttackerIndex = AttackerFight(SlotReto, Retos(SlotReto).Users(UserList(Userindex).flags.SlotRetoUser).Team)
        Deslogged = True
    End If
    TeamUser = Retos(SlotReto).Users(UserList(AttackerIndex).flags.SlotRetoUser).Team
    ExistTeam = ExistCompanero(Userindex)
    If Forzado Then
        If Not ExistTeam Then
            Call FinishFight(SlotReto, TeamUser)
            Call ResetDuelo(SlotReto)
            Exit Sub
        End If
    End If
    With UserList(Userindex)
        If Not CanContinueFight(Userindex) Then
            With Retos(SlotReto)
                For LoopC = LBound(.Users()) To UBound(.Users())
                    With .Users(LoopC)
                        If .Userindex > 0 And .Team = TeamUser Then
                            If Rounds = 0 Then
                                Call AddRound(SlotReto, .Team)
                                Rounds = .Rounds
                            End If
                            Call WriteConsoleMsg(.Userindex, "Has ganado el round. Rounds ganados: " & .Rounds & ".", FontTypeNames.FONTTYPE_VENENO)
                        End If
                    End With
                    If .Users(LoopC).Userindex > 0 Then StatsDuelos .Users(LoopC).Userindex
                Next LoopC
                If Rounds >= (3 / 2) + 0.5 Or Forzado Then
                    Call FinishFight(SlotReto, TeamUser)
                    Call ResetDuelo(SlotReto)
                    Exit Sub
                Else
                    Call FinishFight(SlotReto, TeamUser, True)
                End If
            End With
        End If
        If Deslogged Then
            Call ResetDueloUser(Userindex)
        End If
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : UserDieFight()"
End Sub

Private Sub StatsDuelos(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        If .flags.Muerto Then
            Call RevivirUsuario(Userindex)
            .Stats.MinHp = .Stats.MaxHp
            .Stats.MinMAN = .Stats.MaxMAN
            .Stats.MinSta = .Stats.MaxSta
            Call WriteUpdateUserStats(Userindex)
            Exit Sub
        End If
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMAN = .Stats.MaxMAN
        .Stats.MinSta = .Stats.MaxSta
        Call WriteUpdateUserStats(Userindex)
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : StatsDuelos()"
End Sub

Private Sub FinishFight(ByVal SlotReto As Byte, ByVal Team As Byte, Optional ByVal ChangeTeam As Boolean)
    On Error GoTo ErrorHandler
    Dim LoopC As Integer
    Dim strtemp As String
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).Userindex > 0 Then
                If UserList(.Users(LoopC).Userindex).Counters.TimeFight > 0 Then
                    UserList(.Users(LoopC).Userindex).Counters.TimeFight = 0
                    WriteUserInEvent .Users(LoopC).Userindex
                End If
                If Team = .Users(LoopC).Team Then
                    If ChangeTeam Then
                        Call StatsDuelos(.Users(LoopC).Userindex)
                    Else
                        .Run = False
                        Call StatsDuelos(.Users(LoopC).Userindex)
                        Call RewardUsers(SlotReto, .Users(LoopC).Userindex)
                        If .Users(LoopC).Rounds > 0 Then
                            Call WriteConsoleMsg(.Users(LoopC).Userindex, "Has ganado el reto con " & .Users(LoopC).Rounds & " rounds a tu favor.", FontTypeNames.FONTTYPE_VENENO)
                        Else
                            Call WriteConsoleMsg(.Users(LoopC).Userindex, "Has ganado el reto.", FontTypeNames.FONTTYPE_VENENO)
                        End If
                        strtemp = strtemp & UserList(.Users(LoopC).Userindex).Name & ", "
                    End If
                End If
            End If
        Next LoopC
        If ChangeTeam Then
            Call WarpFight(.Users())
        Else
            strtemp = Left$(strtemp, Len(strtemp) - 2)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos: " & StrTeam(.Users()) & ". Ganador " & strtemp & ". Apuesta por " & .RequiredGld & " Monedas de Oro", FontTypeNames.FONTTYPE_INFO))
            Call LogRetos("Retos: " & StrTeam(.Users()) & ". Ganador el team de " & strtemp & ". Apuesta por " & .RequiredGld & " Monedas de Oro")
        End If
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : FinishFight() en linea " & Erl
End Sub

Public Sub SendFight(ByVal Userindex As Integer, ByVal GldRequired As Long, ByRef Users() As String)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        If ValidateFight(Userindex, GldRequired, Users) Then
            Call SendInvitation(Userindex, GldRequired, Users)
            Call WriteConsoleMsg(Userindex, "Espera noticias para concretar el reto que has enviado. Recuerda que si vuelves a mandar, la anterior solicitud se cancela.", FontTypeNames.FONTTYPE_WARNING)
        End If
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendFight()"
End Sub

Public Sub AcceptFight(ByVal Userindex As Integer, ByVal UserName As String)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        If CanAcceptFight(Userindex, UserName) Then
            Call WriteConsoleMsg(Userindex, "Has aceptado la invitación.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrorHandler:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AcceptFight()"
End Sub

Public Sub WarpPosAnt(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim Pos As WorldPos
    With UserList(Userindex)
        Pos.Map = .PosAnt.Map
        Pos.X = .PosAnt.X
        Pos.Y = .PosAnt.Y
        Call FindLegalPos(Userindex, Pos.Map, Pos.X, Pos.Y)
        Call WarpUserChar(Userindex, Pos.Map, Pos.X, Pos.Y, False)
        .PosAnt.Map = 0
        .PosAnt.X = 0
        .PosAnt.Y = 0
    End With
   On Error GoTo 0
   Exit Sub
ErrorHandler:
    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure WarpPosAnt of Modulo General in line " & Erl
End Sub

