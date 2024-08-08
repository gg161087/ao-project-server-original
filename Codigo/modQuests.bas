Attribute VB_Name = "Quests"
Option Explicit

Public Const MAXUSERQUESTS As Integer = 15
 
Public Function TieneQuest(ByVal Userindex As Integer, ByVal QuestNumber As Integer) As Byte
    Dim i As Integer
    For i = 1 To MAXUSERQUESTS
        If UserList(Userindex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
            TieneQuest = i
            Exit Function
        End If
    Next i
    TieneQuest = 0
End Function
 
Public Function FreeQuestSlot(ByVal Userindex As Integer) As Byte
    Dim i As Integer
    For i = 1 To MAXUSERQUESTS
        If UserList(Userindex).QuestStats.Quests(i).QuestIndex = 0 Then
            FreeQuestSlot = i
            Exit Function
        End If
    Next i
    FreeQuestSlot = 0
End Function
 
Public Sub HandleQuestAccept(ByVal Userindex As Integer)
    Dim NpcIndex  As Integer
    Dim QuestSlot As Byte
    Call UserList(Userindex).incomingData.ReadByte
    NpcIndex = UserList(Userindex).flags.TargetNPC
    If NpcIndex = 0 Then Exit Sub
    If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    QuestSlot = FreeQuestSlot(Userindex)
    With UserList(Userindex).QuestStats.Quests(QuestSlot)
        .QuestIndex = Npclist(NpcIndex).QuestNumber
        If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
        Call WriteConsoleMsg(Userindex, "Has aceptado la mision " & Chr(34) & QuestList(.QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub
 
Public Sub FinishQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
    Dim i              As Integer
    Dim InvSlotsLibres As Byte
    Dim NpcIndex       As Integer
    NpcIndex = UserList(Userindex).flags.TargetNPC
    With QuestList(QuestIndex)
        If .RequiredOBJs > 0 Then
            For i = 1 To .RequiredOBJs
                If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, Userindex) = False Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub
                End If
            Next i
        End If
        If .RequiredNPCs > 0 Then
            For i = 1 To .RequiredNPCs
                If .RequiredNPC(i).Amount > UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub
                End If
            Next i
        End If
        If .RewardOBJs > 0 Then
            For i = 1 To MAX_INVENTORY_SLOTS
                If UserList(Userindex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
            Next i
            If InvSlotsLibres < .RewardOBJs Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                Exit Sub
            End If
        End If
        Call WriteConsoleMsg(Userindex, "Has completado la mision " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & "!", FontTypeNames.FONTTYPE_INFO)
        If .RequiredOBJs Then
            For i = 1 To .RequiredOBJs
                Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, Userindex)
            Next i
        End If
        If .RewardEXP Then
            UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + .RewardEXP
            Call WriteConsoleMsg(Userindex, "Has ganado " & .RewardEXP & " puntos de experiencia como recompensa.", FontTypeNames.FONTTYPE_INFO)
        End If
        If .RewardGLD Then
            UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + .RewardGLD
            Call WriteConsoleMsg(Userindex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO)
        End If
        If .RewardOBJs > 0 Then
            For i = 1 To .RewardOBJs
                If .RewardOBJ(i).Amount Then
                    Call MeterItemEnInventario(Userindex, .RewardOBJ(i))
                    Call WriteConsoleMsg(Userindex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFO)
                End If
            Next i
        End If
        Call CheckUserLevel(Userindex)
        Call UpdateUserInv(True, Userindex, 0)
        Call CleanQuestSlot(Userindex, QuestSlot)
        Call ArrangeUserQuests(Userindex)
        Call AddDoneQuest(Userindex, QuestIndex)
    End With
End Sub
 
Public Sub AddDoneQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer)
    With UserList(Userindex).QuestStats
        .NumQuestsDone = .NumQuestsDone + 1
        ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
        .QuestsDone(.NumQuestsDone) = QuestIndex
    End With
End Sub
 
Public Function UserDoneQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer) As Boolean
    Dim i As Integer
    With UserList(Userindex).QuestStats
        If .NumQuestsDone Then
            For i = 1 To .NumQuestsDone
                If .QuestsDone(i) = QuestIndex Then
                    UserDoneQuest = True
                    Exit Function
                End If
            Next i
        End If
    End With
    UserDoneQuest = False
End Function
 
Public Sub CleanQuestSlot(ByVal Userindex As Integer, ByVal QuestSlot As Integer)
    Dim i As Integer
    With UserList(Userindex).QuestStats.Quests(QuestSlot)
        If .QuestIndex Then
            If QuestList(.QuestIndex).RequiredNPCs Then
                For i = 1 To QuestList(.QuestIndex).RequiredNPCs
                    .NPCsKilled(i) = 0
                Next i
            End If
        End If
        .QuestIndex = 0
    End With
End Sub
 
Public Sub ResetQuestStats(ByVal Userindex As Integer)
    Dim i As Integer
    For i = 1 To MAXUSERQUESTS
        Call CleanQuestSlot(Userindex, i)
    Next i
    With UserList(Userindex).QuestStats
        .NumQuestsDone = 0
        Erase .QuestsDone
    End With
End Sub
 
Public Sub HandleQuest(ByVal Userindex As Integer)
    Dim NpcIndex As Integer
    Dim tmpByte  As Byte
    Call UserList(Userindex).incomingData.ReadByte
    NpcIndex = UserList(Userindex).flags.TargetNPC
    If NpcIndex = 0 Then Exit Sub
    If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If Npclist(NpcIndex).QuestNumber = 0 Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub
    End If
    If UserDoneQuest(Userindex, Npclist(NpcIndex).QuestNumber) Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Ya has hecho una mision para mi.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub
    End If
    If UserList(Userindex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub
    End If
    tmpByte = TieneQuest(Userindex, Npclist(NpcIndex).QuestNumber)
    If tmpByte Then
        Call FinishQuest(Userindex, Npclist(NpcIndex).QuestNumber, tmpByte)
    Else
        tmpByte = FreeQuestSlot(Userindex)
        If tmpByte = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        Call WriteQuestDetails(Userindex, Npclist(NpcIndex).QuestNumber)
    End If
End Sub
 
Public Sub LoadQuests()
    On Error GoTo ErrorHandler
    Dim Reader    As clsIniManager
    Dim NumQuests As Integer
    Dim tmpStr    As String
    Dim i         As Integer
    Dim j         As Integer
    Set Reader = New clsIniManager
    Call Reader.Initialize(DatPath & "Quests.DAT")
    NumQuests = Reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests)
    For i = 1 To NumQuests
        With QuestList(i)
            .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
            .Desc = Reader.GetValue("QUEST" & i, "Desc")
            .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))
            If .RequiredOBJs > 0 Then
                ReDim .RequiredOBJ(1 To .RequiredOBJs)
                For j = 1 To .RequiredOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
            .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))
            If .RequiredNPCs > 0 Then
                ReDim .RequiredNPC(1 To .RequiredNPCs)
                For j = 1 To .RequiredNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNPC(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
            .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
            .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))
            If .RewardOBJs > 0 Then
                ReDim .RewardOBJ(1 To .RewardOBJs)
                For j = 1 To .RewardOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                    .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
        End With
    Next i
    Set Reader = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical
End Sub
 
Public Sub LoadQuestStats(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)
    Dim i           As Integer
    Dim j           As Integer
    Dim tmpStr      As String
    Dim Fields()    As String
    For i = 1 To MAXUSERQUESTS
        With UserList(Userindex).QuestStats.Quests(i)
            tmpStr = UserFile.GetValue("QUESTS", "Q" & i)
            If tmpStr = vbNullString Then
                .QuestIndex = 0
            Else
                Fields = Split(tmpStr, "-")
                .QuestIndex = val(Fields(0))
                If .QuestIndex Then
                    If QuestList(.QuestIndex).RequiredNPCs Then
                        ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                            .NPCsKilled(j) = val(Fields(j))
                        Next j
                    End If
                End If
            End If
        End With
    Next i
    With UserList(Userindex).QuestStats
        tmpStr = UserFile.GetValue("QUESTS", "QuestsDone")
        If tmpStr = vbNullString Then
            .NumQuestsDone = 0
        Else
            Fields = Split(tmpStr, "-")
            .NumQuestsDone = val(Fields(0))
            If .NumQuestsDone Then
                ReDim .QuestsDone(1 To .NumQuestsDone)
                For i = 1 To .NumQuestsDone
                    .QuestsDone(i) = val(Fields(i))
                Next i
            End If
        End If
    End With
End Sub
 
Public Sub SaveQuestStats(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)
    Dim i      As Integer
    Dim j      As Integer
    Dim tmpStr As String
    For i = 1 To MAXUSERQUESTS
        With UserList(Userindex).QuestStats.Quests(i)
            tmpStr = .QuestIndex
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then
                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        tmpStr = tmpStr & "-" & .NPCsKilled(j)
                    Next j
                End If
            End If
            Call UserFile.ChangeValue("QUESTS", "Q" & i, tmpStr)
        End With
    Next i
    With UserList(Userindex).QuestStats
        tmpStr = .NumQuestsDone
        If .NumQuestsDone Then
            For i = 1 To .NumQuestsDone
                tmpStr = tmpStr & "-" & .QuestsDone(i)
            Next i
        End If
        Call UserFile.ChangeValue("QUESTS", "QuestsDone", tmpStr)
    End With
End Sub
 
Public Sub HandleQuestListRequest(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call WriteQuestListSend(Userindex)
End Sub
 
Public Sub ArrangeUserQuests(ByVal Userindex As Integer)
    Dim i As Integer
    Dim j As Integer
    With UserList(Userindex).QuestStats
        For i = 1 To MAXUSERQUESTS - 1
            If .Quests(i).QuestIndex = 0 Then
                For j = i + 1 To MAXUSERQUESTS
                    If .Quests(j).QuestIndex Then
                        .Quests(i) = .Quests(j)
                        Call CleanQuestSlot(Userindex, j)
                        Exit For
                    End If
                Next j
            End If
        Next i
    End With
End Sub
 
Public Sub HandleQuestDetailsRequest(ByVal Userindex As Integer)
    Dim QuestSlot As Byte
    Call UserList(Userindex).incomingData.ReadByte
    QuestSlot = UserList(Userindex).incomingData.ReadByte
    Call WriteQuestDetails(Userindex, UserList(Userindex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)
End Sub
 
Public Sub HandleQuestAbandon(ByVal Userindex As Integer)
    Call UserList(Userindex).incomingData.ReadByte
    Call CleanQuestSlot(Userindex, UserList(Userindex).incomingData.ReadByte)
    Call ArrangeUserQuests(Userindex)
    Call WriteQuestListSend(Userindex)
End Sub
