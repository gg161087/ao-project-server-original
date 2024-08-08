Attribute VB_Name = "modNuevoTimer"
Option Explicit

Public Function IntervaloPermiteAtacarNpc(ByVal NpcIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    With Npclist(NpcIndex)
        If TActual - .Contadores.Ataque >= 3000 Then
            If Actualizar Then
                .Contadores.Ataque = TActual
            End If
            IntervaloPermiteAtacarNpc = True
        Else
            IntervaloPermiteAtacarNpc = False
        End If
    End With
End Function

Public Function IntervaloPermiteLanzarSpell(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(Userindex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
        If Actualizar Then
            UserList(Userindex).Counters.TimerLanzarSpell = TActual
        End If
        Call modAntiCheat.RestaCount(Userindex, 0, 0, 1, 0)
        IntervaloPermiteLanzarSpell = True
    Else
        IntervaloPermiteLanzarSpell = False
        Call modAntiCheat.AddCount(Userindex, 0, 0, 1, 0)
    End If
End Function

Public Function IntervaloPermiteAtacar(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(Userindex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
        If Actualizar Then
            UserList(Userindex).Counters.TimerPuedeAtacar = TActual
            UserList(Userindex).Counters.TimerGolpeUsar = TActual
        End If
        Call modAntiCheat.RestaCount(Userindex, 0, 1, 0, 0)
        IntervaloPermiteAtacar = True
    Else
        IntervaloPermiteAtacar = False
        Call modAntiCheat.AddCount(Userindex, 0, 1, 0, 0)
    End If
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(Userindex).Counters.TimerGolpeUsar >= IntervaloGolpeUsar Then
        If Actualizar Then
            UserList(Userindex).Counters.TimerGolpeUsar = TActual
        End If
        IntervaloPermiteGolpeUsar = True
    Else
        IntervaloPermiteGolpeUsar = False
    End If
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    With UserList(Userindex)
        If .Counters.TimerMagiaGolpe > .Counters.TimerLanzarSpell Then
            Exit Function
        End If
        TActual = GetTickCount()
        If TActual - .Counters.TimerLanzarSpell >= IntervaloMagiaGolpe Then
            If Actualizar Then
                .Counters.TimerMagiaGolpe = TActual
                .Counters.TimerPuedeAtacar = TActual
                .Counters.TimerGolpeUsar = TActual
            End If
            IntervaloPermiteMagiaGolpe = True
        Else
            IntervaloPermiteMagiaGolpe = False
        End If
    End With
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    If UserList(Userindex).Counters.TimerGolpeMagia > UserList(Userindex).Counters.TimerPuedeAtacar Then
        Exit Function
    End If
    TActual = GetTickCount()
    If TActual - UserList(Userindex).Counters.TimerPuedeAtacar >= IntervaloGolpeMagia Then
        If Actualizar Then
            UserList(Userindex).Counters.TimerGolpeMagia = TActual
            UserList(Userindex).Counters.TimerLanzarSpell = TActual
        End If
        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False
    End If
End Function

Public Function IntervaloPermiteTrabajar(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(Userindex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then
        If Actualizar Then UserList(Userindex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajar = True
    Else
        IntervaloPermiteTrabajar = False
    End If
End Function

Public Function IntervaloPermiteUsar(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(Userindex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
        If Actualizar Then
            UserList(Userindex).Counters.TimerUsar = TActual
        End If
        Call modAntiCheat.RestaCount(Userindex, 0, 0, 0, 1)
        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False
        Call modAntiCheat.AddCount(Userindex, 0, 0, 0, 1)
    End If
End Function

Public Function IntervaloPermiteUsarArcos(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(Userindex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
        If Actualizar Then UserList(Userindex).Counters.TimerPuedeUsarArco = TActual
        Call modAntiCheat.RestaCount(Userindex, 1, 0, 0, 0)
        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False
        Call modAntiCheat.AddCount(Userindex, 1, 0, 0, 0)
    End If
End Function

Public Function IntervaloPermiteSerAtacado(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    With UserList(Userindex)
        If Actualizar Then
            .Counters.TimerPuedeSerAtacado = TActual
            .flags.NoPuedeSerAtacado = True
            IntervaloPermiteSerAtacado = False
        Else
            If TActual - .Counters.TimerPuedeSerAtacado >= IntervaloPuedeSerAtacado Then
                .flags.NoPuedeSerAtacado = False
                IntervaloPermiteSerAtacado = True
            Else
                IntervaloPermiteSerAtacado = False
            End If
        End If
    End With
End Function

Public Function IntervaloPerdioNpc(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    With UserList(Userindex)
        If Actualizar Then
            .Counters.TimerPerteneceNpc = TActual
            IntervaloPerdioNpc = False
        Else
            If TActual - .Counters.TimerPerteneceNpc >= IntervaloOwnedNpc Then
                IntervaloPerdioNpc = True
            Else
                IntervaloPerdioNpc = False
            End If
        End If
    End With
End Function

Public Function IntervaloEstadoAtacable(ByVal Userindex As Integer, Optional ByVal Actualizar As Boolean = False) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    With UserList(Userindex)
        If Actualizar Then
            .Counters.TimerEstadoAtacable = TActual
            IntervaloEstadoAtacable = True
        Else
            If TActual - .Counters.TimerEstadoAtacable >= IntervaloAtacable Then
                IntervaloEstadoAtacable = False
            Else
                IntervaloEstadoAtacable = True
            End If
        End If
    End With
End Function

Public Function IntervaloGoHome(ByVal Userindex As Integer, Optional ByVal TimeInterval As Long, Optional ByVal Actualizar As Boolean = False) As Boolean
    Dim TActual As Long
    TActual = GetTickCount()
    With UserList(Userindex)
        If Actualizar Then
            .flags.Traveling = 1
            .Counters.goHome = TActual + TimeInterval
        Else
            If TActual >= .Counters.goHome Then
                IntervaloGoHome = True
            End If
        End If
    End With
End Function

Public Function checkInterval(ByRef startTime As Long, ByVal timeNow As Long, ByVal interval As Long) As Boolean
    Dim lInterval As Long
    If timeNow < startTime Then
        lInterval = &H7FFFFFFF - startTime + timeNow + 1
    Else
        lInterval = timeNow - startTime
    End If
    If lInterval >= interval Then
        startTime = timeNow
        checkInterval = True
    Else
        checkInterval = False
    End If
End Function

