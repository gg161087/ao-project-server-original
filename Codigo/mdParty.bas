Attribute VB_Name = "mdParty"
Option Explicit

Public Const MAX_PARTIES               As Integer = 300
Public Const MINPARTYLEVEL             As Byte = 15
Public Const PARTY_MAXMEMBERS          As Byte = 5
Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = False
Public Const MAXPARTYDELTALEVEL        As Byte = 7
Public Const MAXDISTANCIAINGRESOPARTY  As Byte = 2
Public Const PARTY_MAXDISTANCIA        As Byte = 18
Public Const CASTIGOS                  As Boolean = False
Public ExponenteNivelParty             As Single

Public Type tPartyMember
    Userindex As Integer
    Experiencia As Double
End Type

Public Function NextParty() As Integer
    Dim i As Integer
    NextParty = -1
    For i = 1 To MAX_PARTIES
        If Parties(i) Is Nothing Then
            NextParty = i
            Exit Function
        End If
    Next i
End Function

Public Function PuedeCrearParty(ByVal Userindex As Integer) As Boolean
    PuedeCrearParty = True
    If (UserList(Userindex).flags.Privilegios And PlayerType.User) = 0 Then
        Call WriteConsoleMsg(Userindex, "Los miembros del staff no pueden crear partys!", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf CInt(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma)) * UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) < 100 Then
        Call WriteConsoleMsg(Userindex, "Tu carisma y liderazgo no son suficientes para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf UserList(Userindex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(Userindex, "Estas muerto!!", FontTypeNames.FONTTYPE_PARTY)
        PuedeCrearParty = False
    End If
End Function

Public Sub CrearParty(ByVal Userindex As Integer)
    Dim tInt As Integer
    With UserList(Userindex)
        If .PartyIndex = 0 Then
            If .flags.Muerto = 0 Then
                If .Stats.UserSkills(eSkill.Liderazgo) >= 5 Then
                    tInt = mdParty.NextParty
                    If tInt = -1 Then
                        Call WriteConsoleMsg(Userindex, "Por el momento no se pueden crear mas parties.", FontTypeNames.FONTTYPE_PARTY)
                        Exit Sub
                    Else
                        Set Parties(tInt) = New clsParty
                        If Not Parties(tInt).NuevoMiembro(Userindex) Then
                            Call WriteConsoleMsg(Userindex, "La party esta llena, no puedes entrar.", FontTypeNames.FONTTYPE_PARTY)
                            Set Parties(tInt) = Nothing
                            Exit Sub
                        Else
                            Call WriteConsoleMsg(Userindex, "Has formado una party!", FontTypeNames.FONTTYPE_PARTY)
                            .PartyIndex = tInt
                            .PartySolicitud = 0
                            If Not Parties(tInt).HacerLeader(Userindex) Then
                                Call WriteConsoleMsg(Userindex, "No puedes hacerte lider.", FontTypeNames.FONTTYPE_PARTY)
                            Else
                                Call WriteConsoleMsg(Userindex, "Te has convertido en lider de la party!", FontTypeNames.FONTTYPE_PARTY)
                            End If
                        End If
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No tienes suficientes puntos de liderazgo para liderar una party.", FontTypeNames.FONTTYPE_PARTY)
                End If
            Else
                Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            End If
        Else
            Call WriteConsoleMsg(Userindex, "Ya perteneces a una party.", FontTypeNames.FONTTYPE_PARTY)
        End If
    End With
End Sub

Public Sub SolicitarIngresoAParty(ByVal Userindex As Integer)
    Dim TargetUserIndex As Integer
    Dim PartyIndex      As Integer
    With UserList(Userindex)
        If (.flags.Privilegios And PlayerType.User) = 0 Then
            Call WriteConsoleMsg(Userindex, "Los miembros del staff no pueden unirse a partys!", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        If .PartyIndex > 0 Then
            Call WriteConsoleMsg(Userindex, "Ya perteneces a una party, escribe /SALIRPARTY para abandonarla", FontTypeNames.FONTTYPE_PARTY)
            .PartySolicitud = 0
            Exit Sub
        End If
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
            .PartySolicitud = 0
            Exit Sub
        End If
        TargetUserIndex = .flags.TargetUser
        If TargetUserIndex > 0 Then
            PartyIndex = UserList(TargetUserIndex).PartyIndex
            If PartyIndex > 0 Then
                If Parties(PartyIndex).EsPartyLeader(TargetUserIndex) Then
                    .PartySolicitud = PartyIndex
                    Call WriteConsoleMsg(Userindex, "El lider decidira si te acepta en la party.", FontTypeNames.FONTTYPE_PARTY)
                    Call WriteConsoleMsg(TargetUserIndex, .Name & " solicita ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(Userindex, UserList(TargetUserIndex).Name & " no es lider de la party.", FontTypeNames.FONTTYPE_PARTY)
                End If
            Else
                Call WriteConsoleMsg(Userindex, UserList(TargetUserIndex).Name & " no pertenece a ninguna party.", FontTypeNames.FONTTYPE_PARTY)
                .PartySolicitud = 0
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(Userindex, "Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY", FontTypeNames.FONTTYPE_PARTY)
            .PartySolicitud = 0
        End If
    End With
End Sub

Public Sub SalirDeParty(ByVal Userindex As Integer)
    Dim PI As Integer
    PI = UserList(Userindex).PartyIndex
    If PI > 0 Then
        If Parties(PI).SaleMiembro(Userindex) Then
            Set Parties(PI) = Nothing
        Else
            UserList(Userindex).PartyIndex = 0
        End If
    Else
        Call WriteConsoleMsg(Userindex, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
    Dim PI As Integer
    PI = UserList(leader).PartyIndex
    If PI = UserList(OldMember).PartyIndex Then
        If Parties(PI).SaleMiembro(OldMember) Then
            Set Parties(PI) = Nothing
        Else
            UserList(OldMember).PartyIndex = 0
        End If
    Else
        Call WriteConsoleMsg(leader, LCase(UserList(OldMember).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Function UserPuedeEjecutarComandos(ByVal User As Integer) As Boolean
    Dim PI As Integer
    PI = UserList(User).PartyIndex
    If PI > 0 Then
        If Parties(PI).EsPartyLeader(User) Then
            UserPuedeEjecutarComandos = True
        Else
            Call WriteConsoleMsg(User, "No eres el lider de tu party!", FontTypeNames.FONTTYPE_PARTY)
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(User, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
End Function

Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
    Dim PI    As Integer
    Dim razon As String
    PI = UserList(leader).PartyIndex
    With UserList(NewMember)
        If .PartySolicitud = PI Then
            If Not .flags.Muerto = 1 Then
                If .PartyIndex = 0 Then
                    If Parties(PI).PuedeEntrar(NewMember, razon) Then
                        If Parties(PI).NuevoMiembro(NewMember) Then
                            Call Parties(PI).MandarMensajeAConsola(UserList(leader).Name & " ha aceptado a " & .Name & " en la party.", "Servidor")
                            .PartyIndex = PI
                            .PartySolicitud = 0
                        Else
                            Call SendData(SendTarget.ToAdmins, leader, PrepareMessageConsoleMsg(" Servidor> CATASTROFE EN PARTIES, NUEVO MIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_PARTY))
                        End If
                    Else
                        Call WriteConsoleMsg(leader, razon, FontTypeNames.FONTTYPE_PARTY)
                    End If
                Else
                    If .PartyIndex = PI Then
                        Call WriteConsoleMsg(leader, LCase(.Name) & " ya es miembro de la party.", FontTypeNames.FONTTYPE_PARTY)
                    Else
                        Call WriteConsoleMsg(leader, .Name & " ya es miembro de otra party.", FontTypeNames.FONTTYPE_PARTY)
                    End If
                    Exit Sub
                End If
            Else
                Call WriteConsoleMsg(leader, "Esta muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_PARTY)
                Exit Sub
            End If
        Else
            If .PartyIndex = PI Then
                Call WriteConsoleMsg(leader, LCase(.Name) & " ya es miembro de la party.", FontTypeNames.FONTTYPE_PARTY)
            Else
                Call WriteConsoleMsg(leader, LCase(.Name) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
            End If
            Exit Sub
        End If
    End With
End Sub

Private Function IsPartyMember(ByVal Userindex As Integer, ByVal PartyIndex As Integer)
    Dim MemberIndex As Integer
    For MemberIndex = 1 To PARTY_MAXMEMBERS
    Next MemberIndex
End Function

Public Sub BroadCastParty(ByVal Userindex As Integer, ByRef texto As String)
    Dim PI As Integer
    PI = UserList(Userindex).PartyIndex
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(Userindex).Name)
    End If
End Sub

Public Sub OnlineParty(ByVal Userindex As Integer)
    Dim i                                    As Integer
    Dim PI                                   As Integer
    Dim Text                                 As String
    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    PI = UserList(Userindex).PartyIndex
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline())
        Text = "Nombre(Exp): "
        For i = 1 To PARTY_MAXMEMBERS
            If MembersOnline(i) > 0 Then
                Text = Text & " - " & UserList(MembersOnline(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(i))) & ")"
            End If
        Next i
        Text = Text & ". Experiencia total: " & Parties(PI).ObtenerExperienciaTotal
        Call WriteConsoleMsg(Userindex, Text, FontTypeNames.FONTTYPE_PARTY)
    End If
End Sub

Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
    Dim PI As Integer
    If OldLeader = NewLeader Then Exit Sub
    PI = UserList(OldLeader).PartyIndex
    If PI = UserList(NewLeader).PartyIndex Then
        If UserList(NewLeader).flags.Muerto = 0 Then
            If Parties(PI).HacerLeader(NewLeader) Then
                Call Parties(PI).MandarMensajeAConsola("El nuevo lider de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
            Else
                Call WriteConsoleMsg(OldLeader, "No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_PARTY)
            End If
        Else
            Call WriteConsoleMsg(OldLeader, "Esta muerto!", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(OldLeader, LCase(UserList(NewLeader).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Sub ActualizaExperiencias()
    Dim i As Integer
    If Not PARTY_EXPERIENCIAPORGOLPE Then
        haciendoBK = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Distribuyendo experiencia en parties.", FontTypeNames.FONTTYPE_PARTY))
        For i = 1 To MAX_PARTIES
            If Not Parties(i) Is Nothing Then
                Call Parties(i).FlushExperiencia
            End If
        Next i
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Experiencia distribuida.", FontTypeNames.FONTTYPE_PARTY))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        haciendoBK = False
    End If
End Sub

Public Sub ObtenerExito(ByVal Userindex As Integer, ByVal Exp As Long, Mapa As Integer, X As Integer, Y As Integer)
    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
    Call Parties(UserList(Userindex).PartyIndex).ObtenerExito(Exp, Mapa, X, Y)
End Sub

Public Function CantMiembros(ByVal Userindex As Integer) As Integer
    CantMiembros = 0
    If UserList(Userindex).PartyIndex > 0 Then
        CantMiembros = Parties(UserList(Userindex).PartyIndex).CantMiembros
    End If
End Function

Public Sub ActualizarSumaNivelesElevados(ByVal Userindex As Integer)
    If UserList(Userindex).PartyIndex > 0 Then
        Call Parties(UserList(Userindex).PartyIndex).UpdateSumaNivelesElevados(UserList(Userindex).Stats.ELV)
    End If
End Sub
