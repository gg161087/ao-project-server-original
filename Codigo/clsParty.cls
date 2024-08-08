VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_members(1 To PARTY_MAXMEMBERS) As tPartyMember
Private p_expTotal                       As Long
Private p_Fundador                       As Integer
Private p_CantMiembros                   As Integer
Private p_SumaNivelesElevados            As Single

Public Sub Class_Initialize()
    p_expTotal = 0
    p_CantMiembros = 0
    p_SumaNivelesElevados = 0
End Sub

Public Sub Class_Terminate()

End Sub

Public Sub UpdateSumaNivelesElevados(ByVal Lvl As Integer)
    p_SumaNivelesElevados = p_SumaNivelesElevados - ((Lvl - 1) ^ ExponenteNivelParty) + Lvl ^ ExponenteNivelParty
End Sub

Public Function MiExperiencia(ByVal Userindex As Integer) As Long
    Dim i As Integer
    i = 1
    While i <= PARTY_MAXMEMBERS And p_members(i).Userindex <> Userindex
        i = i + 1
    Wend
    If i <= PARTY_MAXMEMBERS Then
        MiExperiencia = Fix(p_members(i).Experiencia)
    Else
        MiExperiencia = -1
    End If
End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal Mapa As Integer, X As Integer, Y As Integer)
    Dim i           As Integer
    Dim UI          As Integer
    Dim expThisUser As Double
    p_expTotal = p_expTotal + ExpGanada
    For i = 1 To PARTY_MAXMEMBERS
        UI = p_members(i).Userindex
        If UI > 0 Then
            expThisUser = CDbl(ExpGanada * (UserList(p_members(i).Userindex).Stats.ELV ^ ExponenteNivelParty) / p_SumaNivelesElevados)
            If Mapa = UserList(UI).Pos.Map And UserList(UI).flags.Muerto = 0 Then
                If Distance(UserList(UI).Pos.X, UserList(UI).Pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
                    p_members(i).Experiencia = p_members(i).Experiencia + expThisUser
                    If p_members(i).Experiencia < 0 Then
                        p_members(i).Experiencia = 0
                    End If
                    If PARTY_EXPERIENCIAPORGOLPE Then
                        UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + Fix(expThisUser)
                        If UserList(UI).Stats.Exp > MAXEXP Then UserList(UI).Stats.Exp = MAXEXP
                        Call CheckUserLevel(UI)
                        Call WriteUpdateUserStats(UI)
                    End If
                End If
            End If
        End If
    Next i
End Sub

Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
    Dim i As Integer
    For i = 1 To PARTY_MAXMEMBERS
        If p_members(i).Userindex > 0 Then
            Call WriteConsoleMsg(p_members(i).Userindex, " [" & Sender & "] " & texto, FontTypeNames.FONTTYPE_PARTY)
        End If
    Next i
End Sub

Public Function EsPartyLeader(ByVal Userindex As Integer) As Boolean
    EsPartyLeader = (Userindex = p_Fundador)
End Function

Public Function NuevoMiembro(ByVal Userindex As Integer) As Boolean
    Dim i As Integer
    i = 1
    While i <= PARTY_MAXMEMBERS And p_members(i).Userindex > 0
        i = i + 1
    Wend
    If i <= PARTY_MAXMEMBERS Then
        p_members(i).Experiencia = 0
        p_members(i).Userindex = Userindex
        NuevoMiembro = True
        p_CantMiembros = p_CantMiembros + 1
        p_SumaNivelesElevados = p_SumaNivelesElevados + (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
    Else
        NuevoMiembro = False
    End If
End Function

Public Function SaleMiembro(ByVal Userindex As Integer) As Boolean
    Dim i           As Integer
    Dim j           As Integer
    Dim MemberIndex As Integer
    i = 1
    SaleMiembro = False
    While i <= PARTY_MAXMEMBERS And p_members(i).Userindex <> Userindex
        i = i + 1
    Wend
    If i = 1 Then
        SaleMiembro = True
        Call MandarMensajeAConsola("El lider disuelve la party.", "Servidor")
        For j = PARTY_MAXMEMBERS To 1 Step -1
            With p_members(j)
                If .Userindex > 0 Then
                    If j <> 1 Then
                        Call WriteConsoleMsg(.Userindex, "Abandonas la party liderada por " & UserList(p_members(1).Userindex).Name & ".", FontTypeNames.FONTTYPE_PARTY)
                    End If
                    Call WriteConsoleMsg(.Userindex, "Durante la misma has conseguido " & CStr(Fix(.Experiencia)) & " puntos de experiencia.", FontTypeNames.FONTTYPE_PARTY)
                    If Not PARTY_EXPERIENCIAPORGOLPE Then
                        UserList(.Userindex).Stats.Exp = UserList(.Userindex).Stats.Exp + Fix(.Experiencia)
                        If UserList(.Userindex).Stats.Exp > MAXEXP Then UserList(.Userindex).Stats.Exp = MAXEXP
                        Call CheckUserLevel(.Userindex)
                        Call WriteUpdateUserStats(.Userindex)
                    End If
                    Call MandarMensajeAConsola(UserList(.Userindex).Name & " abandona la party.", "Servidor")
                    UserList(.Userindex).PartyIndex = 0
                    p_CantMiembros = p_CantMiembros - 1
                    p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
                    .Userindex = 0
                    .Experiencia = 0
                End If
            End With
        Next j
    Else
        If i <= PARTY_MAXMEMBERS Then
            MemberIndex = p_members(i).Userindex
            With UserList(MemberIndex)
                If Not PARTY_EXPERIENCIAPORGOLPE Then
                    .Stats.Exp = .Stats.Exp + Fix(p_members(i).Experiencia)
                    If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                    Call CheckUserLevel(MemberIndex)
                    Call WriteUpdateUserStats(MemberIndex)
                End If
                Call MandarMensajeAConsola(.Name & " abandona la party.", "Servidor")
                Call WriteConsoleMsg(MemberIndex, "Durante la misma has conseguido " & CStr(Fix(p_members(i).Experiencia)) & " puntos de experiencia.", FontTypeNames.FONTTYPE_PARTY)
                p_CantMiembros = p_CantMiembros - 1
                p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
                MemberIndex = 0
                p_members(i).Experiencia = 0
                p_members(i).Userindex = 0
                CompactMemberList
            End With
        End If
    End If
End Function

Public Function HacerLeader(ByVal Userindex As Integer) As Boolean
    Dim i              As Integer
    Dim OldLeader      As Integer
    Dim oldExp         As Double
    Dim UserIndexIndex As Integer
    UserIndexIndex = 0
    HacerLeader = True
    For i = 1 To PARTY_MAXMEMBERS
        If p_members(i).Userindex > 0 Then
            If p_members(i).Userindex = Userindex Then
                UserIndexIndex = i
            End If
        End If
    Next i
    If Not HacerLeader Then Exit Function
    If UserIndexIndex = 0 Then
        Call LogError("INCONSISTENCIA DE PARTIES")
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!", FontTypeNames.FONTTYPE_PARTY))
        HacerLeader = False
        Exit Function
    End If
    OldLeader = p_members(1).Userindex
    oldExp = p_members(1).Experiencia
    p_members(1).Userindex = p_members(UserIndexIndex).Userindex
    p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia
    p_members(UserIndexIndex).Userindex = OldLeader
    p_members(UserIndexIndex).Experiencia = oldExp
    p_Fundador = p_members(1).Userindex
End Function

Public Sub ObtenerMiembrosOnline(ByRef MemberList() As Integer)
    Dim i As Integer
    For i = 1 To PARTY_MAXMEMBERS
        If p_members(i).Userindex > 0 Then
            MemberList(i) = p_members(i).Userindex
        End If
    Next i
End Sub

Public Function ObtenerExperienciaTotal() As Long
    ObtenerExperienciaTotal = p_expTotal
End Function

Public Function PuedeEntrar(ByVal Userindex As Integer, ByRef razon As String) As Boolean
    Dim esArmada As Boolean
    Dim esCaos   As Boolean
    Dim MyLevel  As Integer
    Dim i        As Integer
    Dim rv       As Boolean
    Dim UI       As Integer
    rv = True
    esArmada = (UserList(Userindex).Faccion.ArmadaReal = 1)
    esCaos = (UserList(Userindex).Faccion.FuerzasCaos = 1)
    MyLevel = UserList(Userindex).Stats.ELV
    rv = Distancia(UserList(p_members(1).Userindex).Pos, UserList(Userindex).Pos) <= MAXDISTANCIAINGRESOPARTY
    If rv Then
        rv = (p_members(PARTY_MAXMEMBERS).Userindex = 0)
        If rv Then
            For i = 1 To PARTY_MAXMEMBERS
                UI = p_members(i).Userindex
                If UI > 0 Then
                    If esArmada And criminal(UI) Then
                        razon = "Los miembros del ejercito real no entran a una party con criminales."
                        rv = False
                    End If
                    If esCaos And Not criminal(UI) Then
                        razon = "Los miembros de la legion oscura no entran a una party con ciudadanos."
                        rv = False
                    End If
                    If UserList(UI).Faccion.ArmadaReal = 1 And criminal(Userindex) Then
                        razon = "Los criminales no entran a parties con miembros del ejercito real."
                        rv = False
                    End If
                    If UserList(UI).Faccion.FuerzasCaos = 1 And Not criminal(Userindex) Then
                        razon = "Los ciudadanos no entran a parties con miembros de la legion oscura."
                        rv = False
                    End If
                    If Not rv Then Exit For
                End If
            Next i
        Else
            razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS
        End If
    Else
        razon = "El usuario " & UserList(Userindex).Name & " se encuentra muy lejos."
    End If
    PuedeEntrar = rv
End Function

Public Sub FlushExperiencia()
    Dim i As Integer
    If Not PARTY_EXPERIENCIAPORGOLPE Then
        For i = 1 To PARTY_MAXMEMBERS
            If p_members(i).Userindex > 0 Then
                If p_members(i).Experiencia > 0 Then
                    UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
                    If UserList(p_members(i).Userindex).Stats.Exp > MAXEXP Then UserList(p_members(i).Userindex).Stats.Exp = MAXEXP
                    Call CheckUserLevel(p_members(i).Userindex)
                Else
                    If Abs(UserList(p_members(i).Userindex).Stats.Exp) > Abs(Fix(p_members(i).Experiencia)) Then
                        UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
                    Else
                        UserList(p_members(i).Userindex).Stats.Exp = 0
                    End If
                End If
                p_members(i).Experiencia = 0
                Call WriteUpdateUserStats(p_members(i).Userindex)
            End If
        Next i
    End If
End Sub

Private Sub CompactMemberList()
    Dim i         As Integer
    Dim freeIndex As Integer
    i = 1
    While i <= PARTY_MAXMEMBERS
        If p_members(i).Userindex = 0 And freeIndex = 0 Then
            freeIndex = i
        ElseIf p_members(i).Userindex > 0 And freeIndex > 0 Then
            p_members(freeIndex).Experiencia = p_members(i).Experiencia
            p_members(freeIndex).Userindex = p_members(i).Userindex
            p_members(i).Userindex = 0
            p_members(i).Experiencia = 0
            i = freeIndex
            freeIndex = 0
        End If
        i = i + 1
    Wend
End Sub

Public Function CantMiembros() As Integer
    CantMiembros = p_CantMiembros
End Function
