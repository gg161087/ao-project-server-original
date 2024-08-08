Attribute VB_Name = "InvNpc"
Option Explicit

Public Function TirarItemAlPiso(Pos As WorldPos, obj As obj, Optional NotPirata As Boolean = True) As WorldPos
    On Error GoTo ErrorHandler
    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    Call Tilelibre(Pos, NuevaPos, obj, NotPirata, True)
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
    End If
    TirarItemAlPiso = NuevaPos
    Exit Function
ErrorHandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, ByVal IsPretoriano As Boolean, ByVal Userindex As Integer)
    On Error Resume Next
    With npc
        Dim i        As Byte
        Dim MiObj    As obj
        Dim NroDrop  As Integer
        Dim Random   As Integer
        Dim ObjIndex As Integer
        If IsPretoriano Then
            For i = 1 To MAX_INVENTORY_SLOTS
                If .Invent.Object(i).ObjIndex > 0 Then
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = .Invent.Object(i).ObjIndex
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
            Next i
            If .GiveGLD > 0 Then Call TirarOroNpc(.GiveGLD, .Pos, Userindex)
            Exit Sub
        End If
        Random = RandomNumber(1, 100)
        If Random <= 90 Then
            NroDrop = 1
            If Random <= 10 Then
                NroDrop = NroDrop + 1
                For i = 1 To 3
                    If RandomNumber(1, 100) <= 10 Then
                        NroDrop = NroDrop + 1
                    Else
                        Exit For
                    End If
                Next i
            End If
            ObjIndex = .Drop(NroDrop).ObjIndex
            If ObjIndex > 0 Then
                If ObjIndex = iORO Then
                    Call TirarOroNpc(.Drop(NroDrop).Amount, .Pos, Userindex)
                Else
                    MiObj.Amount = .Drop(NroDrop).Amount
                    MiObj.ObjIndex = ObjIndex
                    Call TirarItemAlPiso(.Pos, MiObj)
                    If ObjData(ObjIndex).Log = 1 Then
                        Call LogDesarrollo(npc.Name & " dropeo " & MiObj.Amount & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
                    End If
                End If
            End If
        End If
    End With
End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error Resume Next
    Dim i As Integer
    If Npclist(NpcIndex).Invent.NroItems > 0 Then
        For i = 1 To MAX_INVENTORY_SLOTS
            If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
                QuedanItems = True
                Exit Function
            End If
        Next
    End If
    QuedanItems = False
End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
    On Error Resume Next
    Dim ln As String, npcfile As String
    Dim i  As Integer
    npcfile = DatPath & "NPCs.dat"
    For i = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)
        If ObjIndex = val(ReadField(1, ln, 45)) Then
            EncontrarCant = val(ReadField(2, ln, 45))
            Exit Function
        End If
    Next
    EncontrarCant = 0
End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
    On Error Resume Next
    Dim i As Integer
    With Npclist(NpcIndex)
        .Invent.NroItems = 0
        For i = 1 To MAX_INVENTORY_SLOTS
            .Invent.Object(i).ObjIndex = 0
            .Invent.Object(i).Amount = 0
        Next i
        .InvReSpawn = 0
    End With
End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    Dim ObjIndex As Integer
    Dim iCant    As Integer
    With Npclist(NpcIndex)
        ObjIndex = .Invent.Object(Slot).ObjIndex
        If ObjData(.Invent.Object(Slot).ObjIndex).Crucial = 0 Then
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                    Call CargarInvent(NpcIndex)
                End If
            End If
        Else
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                If Not QuedanItems(NpcIndex, ObjIndex) Then
                    iCant = EncontrarCant(NpcIndex, ObjIndex)
                    If iCant Then
                        .Invent.Object(Slot).ObjIndex = ObjIndex
                        .Invent.Object(Slot).Amount = iCant
                        .Invent.NroItems = .Invent.NroItems + 1
                    End If
                End If
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                    Call CargarInvent(NpcIndex)
                End If
            End If
        End If
    End With
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
    Dim LoopC   As Integer
    Dim ln      As String
    Dim npcfile As String
    npcfile = DatPath & "NPCs.dat"
    With Npclist(NpcIndex)
        .Invent.NroItems = val(GetVar(npcfile, "NPC" & .Numero, "NROITEMS"))
        For LoopC = 1 To .Invent.NroItems
            ln = GetVar(npcfile, "NPC" & .Numero, "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
        Next LoopC
    End With
End Sub

Public Sub TirarOroNpc(ByVal Cantidad As Long, ByRef Pos As WorldPos, ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    If Cantidad > 0 Then
        If OroDirectoABille Then
            UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + Cantidad
            Call WriteUpdateGold(Userindex)
            Call WriteConsoleMsg(Userindex, "La criatura te ha dejado " & Format$(Cantidad, "#,###") & " monedas de oro.", FontTypeNames.FONTTYPE_WARNING)
        Else
            Dim MiObj As obj
            Dim RemainingGold As Long
            RemainingGold = Cantidad
            While (RemainingGold > 0)
                If RemainingGold > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    RemainingGold = RemainingGold - MAX_INVENTORY_OBJS
                Else
                    MiObj.Amount = RemainingGold
                    RemainingGold = 0
                End If
                MiObj.ObjIndex = iORO
                Call TirarItemAlPiso(Pos, MiObj)
            Wend
        End If
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en TirarOroNpc en " & Erl & ". Error " & Err.Number & " : " & Err.description)
End Sub
