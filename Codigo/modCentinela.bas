Attribute VB_Name = "modCentinela"
Option Explicit
 
Public isCentinelaActivated As Boolean
Const NUM_CENTINELAS   As Byte = 5
Const NUM_NPC          As Integer = 16
Const MAPA_EXPLOTAR    As Integer = 15
Const X_EXPLOTAR       As Byte = 50
Const Y_EXPLOTAR       As Byte = 50
Const LIMITE_TIEMPO    As Long = 120000
Const CARCEL_TIEMPO    As Byte = 5
Const REVISION_TIEMPO  As Long = 1800000

Type Centinelas
    MiNpcIndex         As Integer
    Invocado           As Boolean
    RevisandoSlot      As Integer
    TiempoInicio       As Long
    CodigoCheck        As String
End Type
 
Public Centinelas(1 To NUM_CENTINELAS) As Centinelas
 
Sub CambiarEstado(ByVal gmIndex As Integer)
    isCentinelaActivated = Not isCentinelaActivated
    Call WriteVar(IniPath & "Server.ini", "INIT", "CentinelaAuditoriaTrabajoActivo", IIf(isCentinelaActivated, 1, 0))
    Dim Message As String
    Message = UserList(gmIndex).Name & " cambio el estado del Centinela a " & IIf(isCentinelaActivated, " ACTIVADO.", " DESACTIVADO.")
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_CENTINELA))
    Call LogGM(UserList(gmIndex).Name, Message)
End Sub
 
Sub EnviarAUsuario(ByVal Userindex As Integer, ByVal CIndex As Byte)
    With Centinelas(CIndex)
        .CodigoCheck = GenerarClave
        .MiNpcIndex = SpawnNpc(NUM_NPC, DarPosicion(Userindex), True, False)
        .Invocado = (.MiNpcIndex <> 0)
        If Not .Invocado Then
            .CodigoCheck = vbNullString
            Exit Sub
        End If
        Call AvisarUsuario(Userindex, CIndex)
        .TiempoInicio = GetTickCount()
        .RevisandoSlot = Userindex
    End With
    With UserList(Userindex).CentinelaUsuario
        .CentinelaCheck = False
        .centinelaIndex = CIndex
        .Codigo = Centinelas(CIndex).CodigoCheck
        .Revisando = True
    End With
End Sub
 
Sub AvisarUsuarios()
    Dim i As Long
    For i = 1 To NUM_CENTINELAS
        With Centinelas(i)
            If .Invocado Then
                Call AvisarUsuario(.RevisandoSlot, CByte(i))
            End If
        End With
    Next i
End Sub
 
Sub AvisarUsuario(ByVal userSlot As Integer, ByVal centinelaIndex As Byte, Optional ByVal IngresoFallido As Boolean = False)
    With Centinelas(centinelaIndex)
        Dim DataSend As String
        If Not IngresoFallido Then
            If (GetTickCount() - .TiempoInicio) > (LIMITE_TIEMPO / 2) Then
                DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, Debes escribir /CENTINELA " & .CodigoCheck & " En menos de 2 minutos.", Npclist(.MiNpcIndex).Char.CharIndex, vbYellow)
            Else
                DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, Tienes menos de un minuto para escribir /CENTINELA " & .CodigoCheck & ".", Npclist(.MiNpcIndex).Char.CharIndex, vbYellow)
            End If
        Else
            DataSend = PrepareMessageChatOverHead("CONTROL DE MACRO INASISTIDO, El codigo ingresado NO es correcto, debes escribir : /CENTINELA " & .CodigoCheck & ".", Npclist(.MiNpcIndex).Char.CharIndex, vbYellow)
        End If
        Call UserList(userSlot).outgoingData.WriteASCIIStringFixed(DataSend)
    End With
End Sub
 
Sub ChekearUsuarios()
    Dim LoopC  As Long
    Dim CIndex As Byte
    For LoopC = 1 To LastUser
        With UserList(LoopC)
            If .CentinelaUsuario.Revisando Then
                Call TiempoUsuario(CInt(LoopC))
            Else
                If .Counters.Trabajando <> 0 Then
                    If Not .CentinelaUsuario.CentinelaCheck Or ((GetTickCount() - .CentinelaUsuario.UltimaRevision) > REVISION_TIEMPO) Then
                        CIndex = ProximoCentinela
                        If CIndex <> 0 Then
                            Call EnviarAUsuario(CInt(LoopC), CIndex)
                        End If
                    End If
                End If
            End If
        End With
    Next LoopC
End Sub
 
Sub IngresaClave(ByVal Userindex As Integer, ByRef Clave As String)
    Clave = UCase$(Clave)
    Dim centinelaIndex As Byte
    centinelaIndex = UserList(Userindex).CentinelaUsuario.centinelaIndex
    If Not centinelaIndex <> 0 Then Exit Sub
    If Not UserList(Userindex).CentinelaUsuario.Revisando Then Exit Sub
    If CheckCodigo(Clave, centinelaIndex) Then
        Call AprobarUsuario(Userindex, centinelaIndex)
    Else
        Call AvisarUsuario(Userindex, centinelaIndex, True)
    End If
End Sub
 
Sub AprobarUsuario(ByVal Userindex As Integer, ByVal CIndex As Byte)
    With UserList(Userindex)
        Call LimpiarIndice(.CentinelaUsuario.centinelaIndex)
        With .CentinelaUsuario
            .CentinelaCheck = True
            .centinelaIndex = 0
            .Codigo = vbNullString
            .Revisando = False
            .UltimaRevision = GetTickCount()
        End With
        Call Protocol.WriteConsoleMsg(Userindex, "El control ha finalizado.", FontTypeNames.FONTTYPE_DIOS)
    End With
End Sub
 
Sub LimpiarIndice(ByVal centinelaIndex As Byte)
    With Centinelas(centinelaIndex)
        .Invocado = False
        .CodigoCheck = vbNullString
        .RevisandoSlot = 0
        .TiempoInicio = 0
        If .MiNpcIndex <> 0 Then
            Call QuitarNPC(.MiNpcIndex)
        End If
    End With
End Sub
 
Sub TiempoUsuario(ByVal Userindex As Integer)
    Dim centinelaIndex As Byte
    With UserList(Userindex).CentinelaUsuario
        centinelaIndex = .centinelaIndex
        If Not centinelaIndex <> 0 Then Exit Sub
        If (GetTickCount - Centinelas(centinelaIndex).TiempoInicio) > LIMITE_TIEMPO Then
            Call UsuarioInActivo(Userindex)
        End If
    End With
End Sub
 
Sub UsuarioInActivo(ByVal Userindex As Integer)
    Call WarpUserChar(Userindex, MAPA_EXPLOTAR, X_EXPLOTAR, Y_EXPLOTAR, True)
    Call Encarcelar(Userindex, CARCEL_TIEMPO, "El centinela")
    If UserList(Userindex).CentinelaUsuario.centinelaIndex <> 0 Then
        Call LimpiarIndice(UserList(Userindex).CentinelaUsuario.centinelaIndex)
    End If
    Call Protocol.WriteConsoleMsg(Userindex, "El centinela te ha ejecutado y encarcelado por Macro Inasistido.", FontTypeNames.FONTTYPE_DIOS)
    Dim ClearType As CentinelaUser
    UserList(Userindex).CentinelaUsuario = ClearType
    UserList(Userindex).CentinelaUsuario.CentinelaCheck = True
End Sub
 
Function GenerarClave() As String
    Dim NumCharacters As Byte
    Dim LoopC         As Long
    NumCharacters = 4
    For LoopC = 1 To NumCharacters
        If (LoopC Mod 2) <> 0 Then
            GenerarClave = GenerarClave & Chr$(RandomNumber(97, 122))
        Else
            GenerarClave = GenerarClave & RandomNumber(1, 9)
        End If
    Next LoopC
    GenerarClave = UCase$(GenerarClave)
End Function
 
Function DarPosicion(ByVal Userindex As Integer) As WorldPos
    With UserList(Userindex)
        DarPosicion = .Pos
        DarPosicion.X = .Pos.X + 1
    End With
End Function
 
Function ProximoCentinela() As Byte
    Dim i As Long
    For i = 1 To NUM_CENTINELAS
        If Not Centinelas(i).Invocado Then
            ProximoCentinela = CByte(i)
            Exit Function
        End If
    Next i
    ProximoCentinela = 0
End Function
 
Function CheckCodigo(ByRef Ingresada As String, ByVal CIndex As Byte) As Boolean
    CheckCodigo = (Not Ingresada <> Centinelas(CIndex).CodigoCheck)
End Function
