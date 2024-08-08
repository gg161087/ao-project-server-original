Attribute VB_Name = "SecurityIp"
Option Explicit

Private IpTables()                     As Long
Private EntrysCounter                  As Long
Private MaxValue                       As Long
Private Multiplicado                   As Long
Private Const IntervaloEntreConexiones As Long = 200
Private MaxConTables()                 As Long
Private MaxConTablesEntry              As Long

Private Enum e_SecurityIpTabla
    IP_INTERVALOS = 1
    IP_LIMITECONEXIONES = 2
End Enum

Public Sub InitIpTables(ByVal OptCountersValue As Long)
    EntrysCounter = OptCountersValue
    Multiplicado = 1
    ReDim IpTables(EntrysCounter * 2 - 1) As Long
    MaxValue = 0
    ReDim MaxConTables(Declaraciones.MaxUsers * 2 - 1) As Long
    MaxConTablesEntry = 0
End Sub

Public Sub IpSecurityMantenimientoLista()
    EntrysCounter = EntrysCounter \ Multiplicado
    Multiplicado = 1
    ReDim IpTables(EntrysCounter * 2 - 1) As Long
    MaxValue = 0
End Sub

Public Function IpSecurityAceptarNuevaConexion(ByVal IP As Long) As Boolean
    Dim IpTableIndex As Long
    IpTableIndex = FindTableIp(IP, IP_INTERVALOS)
    If IpTableIndex >= 0 Then
        If IpTables(IpTableIndex + 1) + IntervaloEntreConexiones <= GetTickCount Then
            IpTables(IpTableIndex + 1) = GetTickCount
            IpSecurityAceptarNuevaConexion = True
            Debug.Print "CONEXION ACEPTADA"
            Exit Function
        Else
            IpSecurityAceptarNuevaConexion = False
            Debug.Print "CONEXION NO ACEPTADA"
            Exit Function
        End If
    Else
        IpTableIndex = Not IpTableIndex
        AddNewIpIntervalo IP, IpTableIndex
        IpTables(IpTableIndex + 1) = GetTickCount
        IpSecurityAceptarNuevaConexion = True
        Exit Function
    End If
End Function

Private Sub AddNewIpIntervalo(ByVal IP As Long, ByVal index As Long)
    If MaxValue + 1 > EntrysCounter Then
        EntrysCounter = EntrysCounter \ Multiplicado
        Multiplicado = Multiplicado + 1
        EntrysCounter = EntrysCounter * Multiplicado
        ReDim Preserve IpTables(EntrysCounter * 2 - 1) As Long
    End If
    Call CopyMemory(IpTables(index + 2), IpTables(index), (MaxValue - index \ 2) * 8)
    IpTables(index) = IP
    MaxValue = MaxValue + 1
End Sub

Public Function IPSecuritySuperaLimiteConexiones(ByVal IP As Long) As Boolean
    Dim IpTableIndex As Long
    IpTableIndex = FindTableIp(IP, IP_LIMITECONEXIONES)
    If IpTableIndex >= 0 Then
        If MaxConTables(IpTableIndex + 1) < LimiteConexionesPorIp Then
            LogIP ("Agregamos conexion a " & IP & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            Debug.Print "suma conexion a " & IP & " total " & MaxConTables(IpTableIndex + 1) + 1
            MaxConTables(IpTableIndex + 1) = MaxConTables(IpTableIndex + 1) + 1
            IPSecuritySuperaLimiteConexiones = False
        Else
            LogIP ("rechazamos conexion de " & IP & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            Debug.Print "rechaza conexion a " & IP
            IPSecuritySuperaLimiteConexiones = True
        End If
    Else
        IPSecuritySuperaLimiteConexiones = False
        If MaxConTablesEntry < Declaraciones.MaxUsers Then
            IpTableIndex = Not IpTableIndex
            AddNewIpLimiteConexiones IP, IpTableIndex
            MaxConTables(IpTableIndex + 1) = 1
        Else
            Call LogCriticEvent("SecurityIP.IPSecuritySuperaLimiteConexiones: Se supero la disponibilidad de slots.")
        End If
    End If
End Function

Private Sub AddNewIpLimiteConexiones(ByVal IP As Long, ByVal index As Long)
    Call CopyMemory(MaxConTables(index + 2), MaxConTables(index), (MaxConTablesEntry - index \ 2) * 8)
    MaxConTables(index) = IP
    MaxConTablesEntry = MaxConTablesEntry + 1
End Sub

Public Sub IpRestarConexion(ByVal IP As Long)
    On Error GoTo ErrorHandler
    Dim key As Long
    key = FindTableIp(IP, IP_LIMITECONEXIONES)
    If key >= 0 Then
        If MaxConTables(key + 1) > 0 Then
            MaxConTables(key + 1) = MaxConTables(key + 1) - 1
        End If
        If MaxConTables(key + 1) <= 0 Then
            MaxConTablesEntry = MaxConTablesEntry - 1
            If key + 2 < UBound(MaxConTables) Then
                Call CopyMemory(MaxConTables(key), MaxConTables(key + 2), (MaxConTablesEntry - (key \ 2)) * 8)
            End If
        End If
    Else
        Call LogIP("restamos conexion a " & IP & " key=" & key & ". NEGATIVO!!")
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en IpRestarConexion. Error: " & Err.Number & " - " & Err.description & ". Ip: " & GetAscIP(IP) & " Key:" & key)
End Sub

Private Function FindTableIp(ByVal IP As Long, ByVal Tabla As e_SecurityIpTabla) As Long
    Dim First  As Long
    Dim Last   As Long
    Dim Middle As Long
    Select Case Tabla
        Case e_SecurityIpTabla.IP_INTERVALOS
            First = 0
            Last = MaxValue - 1
            Do While First <= Last
                Middle = (First + Last) \ 2
                If (IpTables(Middle * 2) < IP) Then
                    First = Middle + 1
                ElseIf (IpTables(Middle * 2) > IP) Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function
                End If
            Loop
            FindTableIp = Not (First * 2)
        
        Case e_SecurityIpTabla.IP_LIMITECONEXIONES
            First = 0
            Last = MaxConTablesEntry - 1
            Do While First <= Last
                Middle = (First + Last) \ 2
                If MaxConTables(Middle * 2) < IP Then
                    First = Middle + 1
                ElseIf MaxConTables(Middle * 2) > IP Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function
                End If
            Loop
            FindTableIp = Not (First * 2)
    End Select
End Function

Public Function DumpTables()
    Dim i As Integer
    For i = 0 To MaxConTablesEntry * 2 - 1 Step 2
        Call LogCriticEvent(GetAscIP(MaxConTables(i)) & " > " & MaxConTables(i + 1))
    Next i
End Function
