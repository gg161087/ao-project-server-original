Attribute VB_Name = "modInvisibles"
Option Explicit
#Const MODO_INVISIBILIDAD = 0

Public Sub PonerInvisible(ByVal Userindex As Integer, ByVal estado As Boolean)
    #If MODO_INVISIBILIDAD = 0 Then
        UserList(Userindex).flags.invisible = IIf(estado, 1, 0)
        UserList(Userindex).flags.Oculto = IIf(estado, 1, 0)
        UserList(Userindex).Counters.Invisibilidad = 0
        Call SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, Not estado)
    #Else
        Dim EstadoActual As Boolean
        EstadoActual = (UserList(Userindex).flags.invisible = 1)
        If Modo = True Then
            UserList(Userindex).flags.invisible = 1
            Call SendData(SendTarget.toMap, UserList(Userindex).Pos.Map, PrepareMessageCharacterRemove(UserList(Userindex).Char.CharIndex))
        Else
        End If
    #End If
End Sub

