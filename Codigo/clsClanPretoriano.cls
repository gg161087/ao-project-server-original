VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClanPretoriano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ePretorianType
    Default = 1
    Custom = 2
End Enum

Private LeftSpawnPos  As WorldPos
Private RightSpawnPos As WorldPos

Private Enum eKingSpells
    LittleHeal = 1
    RemoveParalisis = 2
    BlindAttack = 3
    StupidityAttack = 4
    HealPoison = 5
End Enum

Private Enum eKingAction
    HealAllyParalisis = 1
    HealAllyPoison
    HealAlly
    ChaseTarget
    MassiveAttack
    RetreatAndHeal
End Enum

Private Enum eHealerAction
    HealAllyParalisis = 1
    ParalizePet
    ParalizeUser
    AttackUser
    HealAlly
    RescueFarAlly
End Enum

Private Enum eHealerSpells
    ParalizeUser = 1
    RemoveParalisis = 2
    ParalizeNpc = 3
    Heal = 4
    Tormenta = 5
End Enum

Private Enum eSpellCasterAction
    RemoveInvi = 1
    Attack
End Enum

Private Enum eSpellCasterSpells
    Apocalipsis = 1
    RemoInvi = 2
End Enum

Private Enum eThiefAction
    Attack = 1
    Steal
    None
End Enum

Private Enum eThiefSpells
    Arrow = 1
    Paralisis = 2
End Enum

Private ClanActive                   As Boolean
Private Const THIEF_STEAL_DISTANCE   As Integer = 2
Private WandBreakCounter             As Integer
Private Const MAX_WAND_BREAK_VALUE   As Integer = 6
Private Const START_COUNTDOWN_MIN_HP As Integer = 750
Private KingIdleTime                 As Long
Private Const ALLY_RESPAWN_TIME      As Long = 2000
Private Const MAX_DISTANCE           As Integer = 14
Private Const NRO_PRETORIANOS        As Integer = 9
Private CenterPos                    As WorldPos
Private RespawnClan                  As Boolean
Private ClanMembersAlive             As Integer

Private Type tPretorianData
    NpcIndex As Integer
    NPCAI As ePretorianAI
End Type

Private ClanIndex     As Integer
Private Pretorianos() As tPretorianData

Public Function SpawnClan(ByVal Mapa As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal PretoClanIndex As Integer, Optional ByVal Respawning As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    Dim nPos       As WorldPos
    Dim PretoIndex As Integer
    CenterPos.Map = Mapa
    CenterPos.X = X
    CenterPos.Y = Y
    ReDim Pretorianos(1 To NRO_PRETORIANOS) As tPretorianData
    ClanIndex = PretoClanIndex
    If Not Respawning Then
        If Mapa = MAPA_PRETORIANO Then
            RespawnClan = True
            Call InitializePretoriansVars
            CenterPos = LeftSpawnPos
        Else
            If Not IsValidSpawnArea() Then Exit Function
        End If
    End If
    nPos = CenterPos
    ClanMembersAlive = NRO_PRETORIANOS
    PretoIndex = PretoIndex + 1
    Call SpawnPretorian(nPos, ePretorianAI.King, PretoIndex, True)
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.Healer, PretoIndex)
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X - 6
    Call SpawnPretorian(nPos, ePretorianAI.Healer, PretoIndex)
    PretoIndex = PretoIndex + 1
    nPos.Y = nPos.Y + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    PretoIndex = PretoIndex + 1
    nPos.Y = nPos.Y - 6
    nPos.X = nPos.X - 1
    Call SpawnPretorian(nPos, ePretorianAI.Shooter, PretoIndex)
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X - 4
    Call SpawnPretorian(nPos, ePretorianAI.SpellCaster, PretoIndex)
    PretoIndex = PretoIndex + 1
    Call SpawnPretorian(nPos, ePretorianAI.Thief, PretoIndex)
    WandBreakCounter = MAX_WAND_BREAK_VALUE
    ClanActive = True
    SpawnClan = True
    Exit Function
ErrorHandler:
    Call LogError("Error en CrearClanPretoriano. Error: " & Err.Number & " - " & Err.description)
End Function

Private Function IsValidSpawnArea() As Boolean
    Dim loopX As Long
    Dim LoopY As Long
    For loopX = CenterPos.X - 6 To CenterPos.X + 6
        For LoopY = CenterPos.Y - 6 To CenterPos.Y + 6
            If Not InMapBounds(CenterPos.Map, loopX, LoopY) Then Exit Function
            If HayAgua(CenterPos.Map, loopX, LoopY) Then Exit Function
        Next LoopY
    Next loopX
    IsValidSpawnArea = True
End Function

Private Sub InitializePretoriansVars()
    LeftSpawnPos.Map = MAPA_PRETORIANO
    LeftSpawnPos.X = 35
    LeftSpawnPos.Y = 25
    RightSpawnPos.Map = MAPA_PRETORIANO
    RightSpawnPos.X = 67
    RightSpawnPos.Y = 25
End Sub

Private Sub SpawnPretorian(ByRef SpawnPos As WorldPos, ByVal PretorianAI As Byte, ByVal PretoIndex As Integer, Optional ByVal IsKing As Boolean = False)
    Dim FinalPos       As WorldPos
    Dim NpcIndex       As Integer
    Dim Head           As Integer
    Dim PretoDatNumber As Integer
    PretoDatNumber = RandomNumber(PretorianAIOffset(PretorianAI), PretorianAIOffset(PretorianAI + 1) - 1)
    If (PretoDatNumber And 1) = 1 Then
        Head = RandomTallHead
    Else
        Head = RandomShortHead
    End If
    If IsKing Then
        Dim OtherNpcIndex As Integer
        OtherNpcIndex = MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).NpcIndex
        If OtherNpcIndex <> 0 Then
            Call ClosestLegalPos(SpawnPos, FinalPos)
            If FinalPos.X <> 0 And FinalPos.Y <> 0 Then
                Call SendData(SendTarget.ToNPCArea, OtherNpcIndex, PrepareMessageCharacterMove(Npclist(OtherNpcIndex).Char.CharIndex, FinalPos.X, FinalPos.Y))
                MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).NpcIndex = 0
                Npclist(OtherNpcIndex).Pos = FinalPos
                MapData(FinalPos.Map, FinalPos.X, FinalPos.Y).NpcIndex = OtherNpcIndex
                FinalPos = SpawnPos
            Else
                Call QuitarNPC(OtherNpcIndex)
            End If
        Else
            FinalPos = SpawnPos
        End If
    Else
        Call ClosestLegalPos(SpawnPos, FinalPos, False, True)
    End If
    NpcIndex = CrearNPC(PretorianDatNumbers(PretoDatNumber), FinalPos.Map, FinalPos, Head)
    Pretorianos(PretoIndex).NpcIndex = NpcIndex
    Pretorianos(PretoIndex).NPCAI = PretorianAI
    Npclist(NpcIndex).ClanIndex = ClanIndex
End Sub

Private Function RandomTallHead()
    Dim raza As Integer
    raza = RandomNumber(1, 3)
    If (RandomNumber(0, 1) = 1) Then
        If raza = 1 Then
            RandomTallHead = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)
        ElseIf raza = 2 Then
            RandomTallHead = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)
        Else
            RandomTallHead = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)
        End If
    Else
        If raza = 1 Then
            RandomTallHead = RandomNumber(HUMANO_M_PRIMER_CABEZA, HUMANO_M_ULTIMA_CABEZA)
        ElseIf raza = 2 Then
            RandomTallHead = RandomNumber(ELFO_M_PRIMER_CABEZA, ELFO_M_ULTIMA_CABEZA)
        Else
            RandomTallHead = RandomNumber(DROW_M_PRIMER_CABEZA, DROW_M_ULTIMA_CABEZA)
        End If
    End If
End Function

Private Function RandomShortHead()
    Dim raza As Integer
    raza = RandomNumber(1, 2)
    If (RandomNumber(0, 1) = 1) Then
        If raza = 1 Then
            RandomShortHead = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)
        Else
            RandomShortHead = RandomNumber(GNOMO_H_PRIMER_CABEZA, GNOMO_H_ULTIMA_CABEZA)
        End If
    Else
        If raza = 1 Then
            RandomShortHead = RandomNumber(ENANO_M_PRIMER_CABEZA, ENANO_M_ULTIMA_CABEZA)
        Else
            RandomShortHead = RandomNumber(GNOMO_M_PRIMER_CABEZA, GNOMO_M_ULTIMA_CABEZA)
        End If
    End If
End Function

Public Sub PerformPretorianAI(ByVal NpcIndex As Integer)
    Dim PretorianAI As Byte
    PretorianAI = GetPretorianAI(NpcIndex)
    Select Case PretorianAI
        Case ePretorianAI.King
            Call AI_King(NpcIndex)
            
        Case ePretorianAI.Healer
            Call AI_Healer(NpcIndex)
            
        Case ePretorianAI.SpellCaster
            Call AI_SpellCaster(NpcIndex)
            
        Case ePretorianAI.SwordMaster
            Call AI_SwordMaster(NpcIndex)
            
        Case ePretorianAI.Shooter
            Call AI_Shooter(NpcIndex)
            
        Case ePretorianAI.Thief
            Call AI_Thief(NpcIndex)
    End Select
End Sub

Private Function GetPretorianAI(ByVal NpcIndex As Integer) As Byte
    Dim Counter As Long
    For Counter = 1 To NRO_PRETORIANOS
        If Pretorianos(Counter).NpcIndex = NpcIndex Then
            GetPretorianAI = Pretorianos(Counter).NPCAI
            Exit Function
        End If
    Next Counter
End Function

Public Sub AI_King(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim BestTarget As Integer
    Dim Action     As Byte
    BestTarget = KingBestTarget(NpcIndex, Action)
    Call KingPerformAction(NpcIndex, BestTarget, Action)
    Exit Sub
ErrorHandler:
    LogError ("Error en AI_King. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Function KingBestTarget(ByVal NpcIndex As Integer, ByRef Accion As Byte) As Integer
    On Error GoTo ErrorHandler
    Dim BestTarget As Integer
    If ClanMembersAlive <> 1 Then
        If IntervaloPermiteAtacarNpc(NpcIndex) Then
            BestTarget = AllyParalyzed(NpcIndex, False)
            If BestTarget <> 0 Then
                If RandomNumber(1, 100) < 21 Then
                    Accion = eKingAction.HealAllyParalisis
                End If
            Else
                BestTarget = AllyPoisoned(NpcIndex)
                If BestTarget <> 0 Then
                    Accion = eKingAction.HealAllyPoison
                Else
                    BestTarget = AllyInjured(NpcIndex)
                    If BestTarget <> 0 Then
                        Accion = eKingAction.HealAlly
                    End If
                End If
            End If
        End If
    Else
        Dim Mapa               As Integer
        Dim NPCPosX            As Integer
        Dim NPCPosY            As Integer
        Dim Userindex          As Integer
        Dim Counter            As Long
        Dim BestTargetDistance As Integer
        Dim Distance           As Integer
        
        With Npclist(NpcIndex).Pos
            Mapa = .Map
            NPCPosX = .X
            NPCPosY = .Y
        End With
        
        Dim CounterStart As Long
        Dim CounterEnd   As Long
        Dim CounterStep  As Long
        CounterStep = RandomNumber(0, 1)
        If CounterStep = 1 Then
            CounterStart = 1
            CounterEnd = Areas.ConnGroups(Mapa).Count()
        Else
            CounterStart = Areas.ConnGroups(Mapa).Count()
            CounterEnd = 1
            CounterStep = -1

        End If
        For Counter = CounterStart To CounterEnd Step CounterStep
            Userindex = Areas.ConnGroups(Mapa).Item(Counter)
            If InVisionRange(Userindex, NPCPosX, NPCPosY) Then
                If UserAtacable(Userindex) And UserList(Userindex).flags.Ceguera = 0 Then
                    If BestTarget <> 0 Then
                        Distance = UserDistance(Userindex, NPCPosX, NPCPosY)
                        If Distance < BestTargetDistance Then
                            BestTarget = Userindex
                            BestTargetDistance = Distance
                        End If
                    Else
                        BestTarget = Userindex
                        BestTargetDistance = UserDistance(Userindex, NPCPosX, NPCPosY)
                    End If
                End If
            End If
        Next Counter
        If BestTarget <> 0 Then
            If UserReachable(NpcIndex, BestTarget) Then
                Accion = eKingAction.ChaseTarget
            Else
                Accion = eKingAction.MassiveAttack
            End If
        Else
            Accion = eKingAction.RetreatAndHeal
        End If
    End If
    KingBestTarget = BestTarget
    Exit Function
ErrorHandler:
    LogError ("Error en KingBestTarget. Error: " & Err.Number & " - " & Err.description)
End Function

Private Sub KingPerformAction(ByVal NpcIndex As Integer, ByVal BestTarget As Integer, ByVal Accion As Byte)
    On Error GoTo ErrorHandler
    Dim IdleTime As Long
    With Npclist(NpcIndex)
        Select Case Accion
            Case eKingAction.HealAllyParalisis
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eKingSpells.RemoveParalisis), True)
            
            Case eKingAction.HealAllyPoison
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eKingSpells.HealPoison), True)
            
            Case eKingAction.HealAlly
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eKingSpells.LittleHeal), True)
                If KingIdleTime > ALLY_RESPAWN_TIME Then Call KingReviveAlly(NpcIndex)
                IdleTime = KingIdleTime + 1
                
            Case eKingAction.ChaseTarget, eKingAction.MassiveAttack
                If Accion = eKingAction.ChaseTarget Then
                    Call GreedyWalkTo(NpcIndex, UserList(BestTarget).Pos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
                Else
                    Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eKingSpells.StupidityAttack), True, True)
                    Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eKingSpells.BlindAttack), True, True)
                    Call WriteConsoleMsg(BestTarget, "El rey pretoriano te ha vuelto estupido.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(BestTarget, "El rey pretoriano te ha vuelto ciego ", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(BestTarget, "A la distancia escuchas las siguientes palabras: Cobarde, no eres digno de luchar conmigo si escapas! ", FontTypeNames.FONTTYPE_VENENO)
                End If
                Call KingMeleeAttack(NpcIndex)
            
            Case eKingAction.RetreatAndHeal
                Call ReturnToCenter(NpcIndex)
                If .Stats.MinHp <> .Stats.MaxHp Then
                    Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, .Spells(eKingSpells.LittleHeal), True)
                End If
                If KingIdleTime > ALLY_RESPAWN_TIME Then Call KingReviveAlly(NpcIndex)
                IdleTime = KingIdleTime + 1
        End Select
    End With
    KingIdleTime = IdleTime
    Exit Sub
ErrorHandler:
    LogError ("Error en KingPerformAction. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub KingMeleeAttack(NpcIndex)
    On Error GoTo ErrorHandler
    Dim headingloop As Byte
    Dim nPos        As WorldPos
    Dim Userindex   As Integer
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            Call HeadtoPos(headingloop, nPos)
            If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                Userindex = MapData(nPos.Map, nPos.X, nPos.Y).Userindex
                If Userindex > 0 Then
                    If UserAtacable(Userindex, False) Then
                        If NpcAtacaUser(NpcIndex, Userindex) Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    Exit Sub
ErrorHandler:
    LogError ("Error en KingMeleeAttack. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub KingReviveAlly(ByVal NpcIndex As Integer)
    Dim AllyIndex As Integer
    Dim nPos      As WorldPos
    With Npclist(NpcIndex)
        If .Pos.X = CenterPos.X Then
            If .Pos.Y = CenterPos.Y Then
                AllyIndex = AllyDead()
                If AllyIndex <> 0 Then
                    nPos = CenterPos
                    nPos.X = nPos.X + 1
                    Call SpawnPretorian(nPos, Pretorianos(AllyIndex).NPCAI, AllyIndex)
                    ClanMembersAlive = ClanMembersAlive + 1
                End If
                KingIdleTime = 0
            End If
        End If
    End With
End Sub

Public Sub AI_Healer(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim BestTarget As Integer
    Dim Accion     As Byte
    If IntervaloPermiteAtacarNpc(NpcIndex) Then
        BestTarget = AllyParalyzed(NpcIndex, False)
        If BestTarget <> 0 Then
            Accion = eHealerAction.HealAllyParalisis
        Else
            BestTarget = HealerBestTarget(NpcIndex, Accion)
            If BestTarget = 0 Then
                BestTarget = AllyInjured(NpcIndex)
                If BestTarget <> 0 Then
                    Accion = eHealerAction.HealAlly
                End If
            End If
        End If
    End If
    If BestTarget = 0 Then
        If Npclist(NpcIndex).flags.Paralizado = 0 Then
            BestTarget = AllyParalyzed(NpcIndex, True)
            If BestTarget <> 0 Then
                Accion = eHealerAction.RescueFarAlly
            End If
        End If
    End If
    Dim CanMove As Boolean
    Call HealerPerformAction(NpcIndex, BestTarget, Accion, CanMove)
    If CanMove Then Call HealerMove(NpcIndex)
    Exit Sub
ErrorHandler:
    LogError ("Error en AI_Healer. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Function HealerBestTarget(ByVal NpcIndex As Integer, ByRef Accion As Byte) As Integer
    On Error GoTo ErrorHandler
    Dim Userindex           As Integer
    Dim Counter             As Long
    Dim PetIndex            As Integer
    Dim BestTarget          As Integer
    Dim BestTargetInvisible As Boolean
    Dim NpcX                As Integer
    Dim NpcY                As Integer
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    Dim CounterStart As Long
    Dim CounterEnd   As Long
    Dim CounterStep  As Long
    CounterStep = RandomNumber(0, 1)
    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = Areas.ConnGroups(CenterPos.Map).Count()
    Else
        CounterStart = Areas.ConnGroups(CenterPos.Map).Count()
        CounterEnd = 1
        CounterStep = -1
    End If
    For Counter = CounterStart To CounterEnd Step CounterStep
        Userindex = Areas.ConnGroups(CenterPos.Map).Item(Counter)
        If UserAtacable(Userindex, False) Then
            PetIndex = CheckNearUserPets(NpcIndex, Userindex)
            If PetIndex <> 0 Then
                BestTarget = PetIndex
                Accion = eHealerAction.ParalizePet
                Exit For
            End If
            If InVisionRange(Userindex, NpcX, NpcY) Then
                With UserList(Userindex)
                    If .flags.Paralizado = 0 Then
                        If (.flags.invisible = 0) And (.flags.Oculto = 0) Then
                            BestTarget = Userindex
                            Accion = eHealerAction.ParalizeUser
                        End If
                    ElseIf BestTarget = 0 Then
                        BestTarget = Userindex
                        Accion = eHealerAction.AttackUser
                    End If
                End With
            End If
        End If
    Next Counter
    HealerBestTarget = BestTarget
    Exit Function
ErrorHandler:
    LogError ("Error en HealerBestTarget. Error: " & Err.Number & " - " & Err.description)
End Function

Private Sub HealerPerformAction(ByVal NpcIndex As Integer, ByVal BestTarget As Integer, ByVal Accion As Byte, ByRef CanMove As Boolean)
    With Npclist(NpcIndex)
        Select Case Accion
            Case eHealerAction.HealAllyParalisis
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eHealerSpells.RemoveParalisis), True)
                CanMove = False
                
            Case eHealerAction.ParalizePet
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eHealerSpells.ParalizeNpc), True)
                CanMove = True
                
            Case eHealerAction.ParalizeUser
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePalabrasMagicas(.Spells(eHealerSpells.ParalizeUser), .Char.CharIndex))
                    
                Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eHealerSpells.ParalizeUser), True)
                CanMove = False
                
            Case eHealerAction.AttackUser
                Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eHealerSpells.Tormenta), True, True)
                CanMove = False
                
            Case eHealerAction.HealAlly
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eHealerSpells.Heal), True)
                CanMove = True
                
            Case eHealerAction.RescueFarAlly
                Call GreedyWalkTo(NpcIndex, .Pos.Map, Npclist(BestTarget).Pos.X, Npclist(BestTarget).Pos.Y)
                CanMove = False
                 
            Case Else
                CanMove = True
        End Select
    End With
End Sub

Private Sub HealerMove(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Then Exit Sub
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        Call ReturnToCenter(NpcIndex)
        If TargetClose(.Pos) Then
            Call ReturnToCenter(NpcIndex)
        End If
    End With
End Sub

Public Sub AI_SpellCaster(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim BestTarget As Integer
    Dim Accion     As Byte
    If SpellCasterSacrifice(NpcIndex) Then Exit Sub
    BestTarget = SpellCasterBestTarget(NpcIndex, Accion)
    If BestTarget <> 0 Then
        Call SpellCasterPerformAction(NpcIndex, BestTarget, Accion)
    End If
    Call SpellCasterMove(NpcIndex, BestTarget)
    Exit Sub
ErrorHandler:
    LogError ("Error en AI_SpellCaster. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Function SpellCasterSacrifice(ByVal NpcIndex As Integer) As Boolean
    With Npclist(NpcIndex)
        If .Stats.MinHp <= START_COUNTDOWN_MIN_HP Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARGRANDE, INFINITE_LOOPS))
            If IntervaloPermiteAtacarNpc(NpcIndex) Then
                WandBreakCounter = WandBreakCounter - 1
                Call SpellCasterDestroyWand(NpcIndex, eSpellCasterSpells.Apocalipsis)
                If WandBreakCounter = 0 Then Call MuereNpc(NpcIndex, 0)
            End If
            SpellCasterSacrifice = True
        Else
            If WandBreakCounter <> MAX_WAND_BREAK_VALUE Then
                WandBreakCounter = MAX_WAND_BREAK_VALUE
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If
        End If
    End With
End Function

Private Function SpellCasterBestTarget(ByVal NpcIndex As Integer, ByRef Accion As Byte) As Integer
    On Error GoTo ErrorHandler
    Dim Userindex           As Integer
    Dim Counter             As Long
    Dim BestTarget          As Integer
    Dim NpcX                As Integer
    Dim NpcY                As Integer
    Dim BestTargetInvisible As Boolean
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    Dim CounterStart As Long
    Dim CounterEnd   As Long
    Dim CounterStep  As Long
    CounterStep = RandomNumber(0, 1)
    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = Areas.ConnGroups(CenterPos.Map).Count()
    Else
        CounterStart = Areas.ConnGroups(CenterPos.Map).Count()
        CounterEnd = 1
        CounterStep = -1
    End If
    For Counter = CounterStart To CounterEnd Step CounterStep
        Userindex = Areas.ConnGroups(CenterPos.Map).Item(Counter)
        If InVisionRange(Userindex, NpcX, NpcY) Then
            If UserAtacable(Userindex, False) Then
                With UserList(Userindex)
                    If (.flags.invisible = 1) Or (.flags.Oculto = 1) Then
                        If (RandomNumber(1, 100) <= 35) Then
                            BestTarget = Userindex
                            Accion = eSpellCasterAction.RemoveInvi
                            Exit For
                        End If
                        If .flags.Paralizado = 1 Then
                            BestTarget = Userindex
                            BestTargetInvisible = True
                            Accion = eSpellCasterAction.Attack
                        End If
                    ElseIf (.flags.Paralizado = 1) Then
                        If Not BestTargetInvisible Or BestTarget = 0 Then
                            BestTarget = Userindex
                            Accion = eSpellCasterAction.Attack
                        End If
                    ElseIf BestTarget = 0 Then
                        BestTarget = Userindex
                        Accion = eSpellCasterAction.Attack
                    End If
                End With
            End If
        End If
    Next Counter
    SpellCasterBestTarget = BestTarget
    Exit Function
ErrorHandler:
    LogError ("Error en SpellCasterBestTarget. Error: " & Err.Number & " - " & Err.description)
End Function

Private Sub SpellCasterPerformAction(ByVal NpcIndex As Integer, ByVal BestTarget As Integer, ByVal Accion As Byte)
    With Npclist(NpcIndex)
        Select Case Accion
            Case eSpellCasterAction.Attack
                Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eSpellCasterSpells.Apocalipsis), True, True)
            
            Case eSpellCasterAction.RemoveInvi
                Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eSpellCasterSpells.RemoInvi), True, True)
        End Select
    End With
End Sub

Private Sub SpellCasterMove(ByVal NpcIndex As Integer, ByVal BestTarget As Integer)
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Then Exit Sub
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        If (BestTarget = 0) And (IntervaloPermiteAtacarNpc(NpcIndex)) Then Call ReturnToCenter(NpcIndex)
        If TargetClose(.Pos) Then
            Call ReturnToCenter(NpcIndex)
        End If
    End With
End Sub

Private Sub SpellCasterDestroyWand(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer)
    On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        Select Case WandBreakCounter
            Case 5
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Rahma", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 4
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("vortax", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 3
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Zill", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 2
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("yaka E'nta", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 1
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Korata!!", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 0
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(vbNullString, .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))
                Call DealWandDamage(NpcIndex, .Spells(SpellIndex))
        End Select
    End With
    Exit Sub
ErrorHandler:
    LogError ("Error en SpellCasterDestroyWand. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub DealWandDamage(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer)
    On Error GoTo ErrorHandler
    Dim Userindex As Integer
    Dim Counter   As Long
    Dim Distancia As Integer
    Dim Danio     As Integer
    Dim X         As Integer
    Dim Y         As Integer
    With Npclist(NpcIndex)
        X = .Pos.X
        Y = .Pos.Y
        For Counter = 1 To Areas.ConnGroups(CenterPos.Map).Count()
            Userindex = Areas.ConnGroups(CenterPos.Map).Item(Counter)
            If Userindex <> 0 Then
                With UserList(Userindex)
                    If UserAtacable(Userindex, False, False) Then
                        Distancia = UserDistance(Userindex, X, Y)
                        Danio = Abs(Int(880 / (Distancia ^ (3 / 7))))
                        .Stats.MinHp = .Stats.MinHp - Danio
                        Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).Name & " te ha quitado " & Danio & " puntos de vida al romper su vara.", FontTypeNames.FONTTYPE_FIGHT)
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y))
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                        If .Stats.MinHp < 1 Then
                            .Stats.MinHp = 0
                            Call UserDie(Userindex)
                        Else
                            Call WriteUpdateHP(Userindex)
                            If .NroMascotas > 0 Then Call DealWandDamageToPets(NpcIndex, Userindex, SpellIndex)
                        End If
                    End If
                End With
            End If
        Next Counter
    End With
    Exit Sub
ErrorHandler:
    LogError ("Error en DealWandDamage. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub DealWandDamageToPets(ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByVal SpellIndex As Integer)
    On Error GoTo ErrorHandler
    Dim PetIndex   As Integer
    Dim PetCounter As Long
    Dim Distancia  As Integer
    Dim Danio      As Integer
    Dim NpcX       As Integer
    Dim NpcY       As Integer
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    For PetCounter = 1 To MAXMASCOTAS
        PetIndex = UserList(Userindex).MascotasIndex(PetCounter)
        If PetIndex > 0 Then
            With Npclist(PetIndex)
                Distancia = NpcDistance(PetIndex, NpcX, NpcY)
                Danio = Abs(Int(880 / (Distancia ^ (3 / 7))))
                .Stats.MinHp = .Stats.MinHp - Danio
                Call SendData(SendTarget.ToNPCArea, PetIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y))
                Call SendData(SendTarget.ToNPCArea, PetIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0
                    Call MuereNpc(PetIndex, 0)
                End If
            End With
        End If
    Next PetCounter
    Exit Sub
ErrorHandler:
    LogError ("Error en DealWandDamageToPets. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub AI_SwordMaster(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim BestTarget As Integer
    BestTarget = SwordMasterBestTarget(NpcIndex)
    Call SwordMasterMove(NpcIndex, BestTarget)
    Call SwordMasterAttack(NpcIndex)
    Exit Sub
ErrorHandler:
    LogError ("Error en AI_SwordMaster. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Function SwordMasterBestTarget(ByVal NpcIndex As Integer) As Integer
    On Error GoTo ErrorHandler
    Dim Userindex          As Integer
    Dim Counter            As Long
    Dim BestTarget         As Integer
    Dim BestTargetDistance As Integer
    Dim NpcX               As Integer
    Dim NpcY               As Integer
    Dim Distance           As Integer
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    Dim CounterStart As Long
    Dim CounterEnd   As Long
    Dim CounterStep  As Long
    CounterStep = RandomNumber(0, 1)
    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = Areas.ConnGroups(CenterPos.Map).Count()
    Else
        CounterStart = Areas.ConnGroups(CenterPos.Map).Count()
        CounterEnd = 1
        CounterStep = -1
    End If
    For Counter = CounterStart To CounterEnd Step CounterStep
        Userindex = Areas.ConnGroups(CenterPos.Map).Item(Counter)
        If InVisionRange(Userindex, NpcX, NpcY) Then
            If UserAtacable(Userindex) Then
                If UserReachable(NpcIndex, Userindex) Then
                    If BestTarget <> 0 Then
                        Distance = UserDistance(Userindex, NpcX, NpcY)
                        If Distance < BestTargetDistance Then
                            BestTarget = Userindex
                            BestTargetDistance = Distance
                        End If
                    Else
                        BestTarget = Userindex
                        BestTargetDistance = UserDistance(Userindex, NpcX, NpcY)
                    End If
                End If
            End If
        End If
    Next Counter
    SwordMasterBestTarget = BestTarget
    Exit Function
ErrorHandler:
    LogError ("Error en SwordMasterBestTarget. Error: " & Err.Number & " - " & Err.description)
End Function

Private Sub SwordMasterMove(ByVal NpcIndex As Integer, ByVal BestTarget As Integer)
    On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Then Exit Sub
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        If FarFromTeam(NpcIndex) Then
            Call ReturnToCenter(NpcIndex)
            Exit Sub
        End If
        If BestTarget > 0 Then
            With UserList(BestTarget)
                Call GreedyWalkTo(NpcIndex, CenterPos.Map, .Pos.X, .Pos.Y)
            End With
        Else
            Call ReturnToCenter(NpcIndex)
        End If
    End With
    Exit Sub
ErrorHandler:
    LogError ("Error en SwordMasterMove. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub SwordMasterAttack(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim headingloop As Byte
    Dim NpcPos      As WorldPos
    Dim Userindex   As Integer
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            NpcPos = .Pos
            Call HeadtoPos(headingloop, NpcPos)
            Userindex = MapData(NpcPos.Map, NpcPos.X, NpcPos.Y).Userindex
            If Userindex > 0 Then
                If UserAtacable(Userindex, False) Then
                    If NpcAtacaUser(NpcIndex, Userindex) Then
                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                        Exit For
                    End If
                End If
            End If
        Next headingloop
    End With
    Exit Sub
ErrorHandler:
    LogError ("Error en SwordMasterAttack. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub AI_Shooter(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim BestTarget As Integer
    BestTarget = ShooterBestTarget(NpcIndex)
    If BestTarget > 0 Then
        Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, Npclist(NpcIndex).Spells(1))
    End If
    Call ShooterMove(NpcIndex, BestTarget)
    Exit Sub
ErrorHandler:
    LogError ("Error en AI_Shooter. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Function ShooterBestTarget(ByVal NpcIndex As Integer) As Integer
    On Error GoTo ErrorHandler
    Dim Userindex          As Integer
    Dim Counter            As Long
    Dim BestTarget         As Integer
    Dim BestTargetDistance As Integer
    Dim NpcX               As Integer
    Dim NpcY               As Integer
    Dim Distance           As Integer
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    Dim CounterStart As Long
    Dim CounterEnd   As Long
    Dim CounterStep  As Long
    CounterStep = RandomNumber(0, 1)
    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = Areas.ConnGroups(CenterPos.Map).Count()
    Else
        CounterStart = Areas.ConnGroups(CenterPos.Map).Count()
        CounterEnd = 1
        CounterStep = -1
    End If
    For Counter = CounterStart To CounterEnd Step CounterStep
        Userindex = Areas.ConnGroups(CenterPos.Map).Item(Counter)
        If InVisionRange(Userindex, NpcX, NpcY) Then
            If UserAtacable(Userindex) Then
                If EsClaseMagica(Userindex) Then
                    BestTarget = Userindex
                    Exit For
                Else
                    If BestTarget <> 0 Then
                        Distance = UserDistance(Userindex, NpcX, NpcY)
                        If Distance < BestTargetDistance Then
                            BestTarget = Userindex
                            BestTargetDistance = Distance
                        End If
                    Else
                        BestTarget = Userindex
                        BestTargetDistance = UserDistance(Userindex, NpcX, NpcY)
                    End If
                End If
            End If
        End If
    Next Counter
    ShooterBestTarget = BestTarget
    Exit Function
ErrorHandler:
    LogError ("Error en ShooterBestTarget. Error: " & Err.Number & " - " & Err.description)
End Function

Private Sub ShooterMove(ByVal NpcIndex As Integer, ByVal BestTarget As Integer)
    On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Then Exit Sub
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        If FarFromTeam(NpcIndex) Then
            Call ReturnToCenter(NpcIndex)
            Exit Sub
        End If
        If Abs(CenterPos.X - .Pos.X) > 4 Or Abs(CenterPos.Y - .Pos.Y) > 4 Then
            Call ReturnToCenter(NpcIndex)
            Exit Sub
        End If
        If BestTarget > 0 Then
            Call GreedyWalkTo(NpcIndex, CenterPos.Map, CenterPos.X + ((UserList(BestTarget).Pos.X - CenterPos.X) / 2), CenterPos.Y + ((UserList(BestTarget).Pos.Y - CenterPos.Y) / 2))
            Exit Sub
        Else
            Dim Counter   As Long
            Dim Userindex As Integer
            For Counter = 1 To Areas.ConnGroups(CenterPos.Map).Count()
                Userindex = Areas.ConnGroups(CenterPos.Map).Item(Counter)
                If InVisionRange(Userindex, CenterPos.X, CenterPos.Y, True) Then
                    If UserAtacable(Userindex) Then
                        With UserList(Userindex)
                            Call GreedyWalkTo(NpcIndex, CenterPos.Map, .Pos.X, .Pos.Y)
                        End With
                        Exit Sub
                    End If
                End If
            Next Counter
        End If
        If IntervaloPermiteAtacarNpc(NpcIndex) Then Call ReturnToCenter(NpcIndex)
    End With
    Exit Sub
ErrorHandler:
    LogError ("Error en ShooterMove. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub AI_Thief(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim BestTarget As Integer
    Dim Action     As Byte
    BestTarget = ThiefBestTarget(NpcIndex, Action)
    Call ThiefPerfomAction(NpcIndex, BestTarget, Action)
    Exit Sub
ErrorHandler:
    LogError ("Error en AI_Thief. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Function ThiefBestTarget(ByVal NpcIndex As Integer, ByRef Action As Byte) As Integer
    On Error GoTo ErrorHandler
    Dim Userindex          As Integer
    Dim Counter            As Long
    Dim BestTarget         As Integer
    Dim BestTargetDistance As Integer
    Dim NpcX               As Integer
    Dim NpcY               As Integer
    Dim Distance           As Integer
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    Dim CounterStart As Long
    Dim CounterEnd   As Long
    Dim CounterStep  As Long
    CounterStep = RandomNumber(0, 1)
    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = Areas.ConnGroups(CenterPos.Map).Count()
    Else
        CounterStart = Areas.ConnGroups(CenterPos.Map).Count()
        CounterEnd = 1
        CounterStep = -1
    End If
    For Counter = CounterStart To CounterEnd Step CounterStep
        Userindex = Areas.ConnGroups(CenterPos.Map).Item(Counter)
        If InVisionRange(Userindex, NpcX, NpcY) Then
            If UserAtacable(Userindex, False) Then
                If UserReachable(NpcIndex, Userindex) Then
                    If UserList(Userindex).Clase = eClass.Hunter Then
                        BestTarget = Userindex
                        BestTargetDistance = UserDistance(Userindex, NpcX, NpcY)
                        Exit For
                    End If
                    If BestTarget <> 0 Then
                        Distance = UserDistance(Userindex, NpcX, NpcY)
                        If Distance < BestTargetDistance Then
                            BestTarget = Userindex
                            BestTargetDistance = Distance
                        End If
                    Else
                        BestTarget = Userindex
                        BestTargetDistance = UserDistance(Userindex, NpcX, NpcY)
                    End If
                End If
            End If
        End If
    Next Counter
    If BestTarget = 0 Then
        Action = eThiefAction.None
    Else
        If BestTargetDistance <= THIEF_STEAL_DISTANCE Then
            Action = eThiefAction.Steal
        Else
            Action = eThiefAction.Attack
        End If
    End If
    ThiefBestTarget = BestTarget
    Exit Function
ErrorHandler:
    LogError ("Error en ThiefBestTarget. Error: " & Err.Number & " - " & Err.description)
End Function

Private Sub ThiefPerfomAction(ByVal NpcIndex As Integer, ByVal BestTarget As Integer, ByVal Action As Byte)
    On Error GoTo ErrorHandler
    Dim WeaponEqpSlot   As Byte
    Dim MunicionEqpSlot As Byte
    With Npclist(NpcIndex)
        Select Case Action
            Case eThiefAction.Attack, eThiefAction.Steal
                If IntervaloPermiteAtacarNpc(NpcIndex) Then
                    If Action = eThiefAction.Attack Then
                        Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eThiefSpells.Arrow), False, True)
                        If RandomNumber(1, 100) < 10 Then
                            Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eThiefSpells.Paralisis), False, True)
                            Call WriteConsoleMsg(BestTarget, "" & .Name & " te ha paralizado con su golpe!!", FontTypeNames.FONTTYPE_VENENO)
                        ElseIf RandomNumber(1, 100) < 22 Then
                            WeaponEqpSlot = UserList(BestTarget).Invent.WeaponEqpSlot
                            If WeaponEqpSlot <> 0 Then Call Desequipar(BestTarget, WeaponEqpSlot)
                        End If
                    ElseIf Not EsGm(BestTarget) Then
                        If RandomNumber(1, 100) < 16 Then
                            WeaponEqpSlot = UserList(BestTarget).Invent.WeaponEqpSlot
                            If WeaponEqpSlot <> 0 Then
                                If SnatchItem(NpcIndex, BestTarget, WeaponEqpSlot) Then
                                    Call WriteConsoleMsg(BestTarget, "" & .Name & " te ha arrebatado tu arma!!", FontTypeNames.FONTTYPE_VENENO)
                                End If
                            Else
                                MunicionEqpSlot = UserList(BestTarget).Invent.MunicionEqpSlot
                                If MunicionEqpSlot <> 0 Then
                                    If SnatchItem(NpcIndex, BestTarget, MunicionEqpSlot) Then
                                        Call WriteConsoleMsg(BestTarget, "" & .Name & " te ha arrebatado tus municiones!!", FontTypeNames.FONTTYPE_VENENO)
                                    End If
                                End If
                            End If
                        ElseIf RandomNumber(1, 100) < 16 Then
                            Dim Slot As Byte
                            Slot = RandomNumber(1, UserList(BestTarget).CurrentInventorySlots)
                            If ThiefStealITem(NpcIndex, BestTarget, Slot) Then
                                Call WriteConsoleMsg(BestTarget, "" & .Name & " te esta robando!!", FontTypeNames.FONTTYPE_VENENO)
                            End If
                        End If
                    End If
                End If
                If .flags.invisible = 0 Then
                    If RandomNumber(1, 100) < 25 Then Call ThiefTurnInvisible(NpcIndex, True)
                Else
                    If RandomNumber(1, 100) < 13 Then Call ThiefTurnInvisible(NpcIndex, False)
                End If
                Call GreedyWalkTo(NpcIndex, CenterPos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
                
            Case eThiefAction.None
                If .flags.invisible = 1 Then Call ThiefTurnInvisible(NpcIndex, False)
                Call ReturnToCenter(NpcIndex)
        End Select
    End With
    Exit Sub
ErrorHandler:
    LogError ("Error en ThiefPerfomAction. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Function SnatchItem(ByVal NpcIndex As Integer, ByVal targetIndex As Integer, ByVal Slot As Byte) As Boolean
    On Error GoTo ErrorHandler
    Dim DropObj As obj
    Dim nPos    As WorldPos
    With UserList(targetIndex)
        DropObj.Amount = .Invent.Object(Slot).Amount
        DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex
        Call Tilelibre(.Pos, nPos, DropObj, False, True)
        If nPos.X <> 0 And nPos.Y <> 0 Then
            Call MakeObj(DropObj, nPos.Map, nPos.X, nPos.Y)
            Call QuitarUserInvItem(targetIndex, Slot, DropObj.Amount)
            Call UpdateUserInv(False, targetIndex, Slot)
            SnatchItem = True
        End If
    End With
    Exit Function
ErrorHandler:
    LogError ("Error en SnatchItem. Error: " & Err.Number & " - " & Err.description)
End Function

Private Function ThiefStealITem(ByVal NpcIndex As Integer, ByVal targetIndex As Integer, ByVal Slot As Byte) As Boolean
    Dim StolenObj As obj
    Dim NroItems  As Integer
    If UserList(targetIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Function
    If Not ObjEsRobable(targetIndex, Slot) Then Exit Function
    NroItems = Npclist(NpcIndex).Invent.NroItems
    If NroItems = 20 Then
        ThiefStealITem = SnatchItem(NpcIndex, targetIndex, Slot)
    Else
        StolenObj.Amount = UserList(targetIndex).Invent.Object(Slot).Amount
        StolenObj.ObjIndex = UserList(targetIndex).Invent.Object(Slot).ObjIndex
        With Npclist(NpcIndex)
            NroItems = NroItems + 1
            .Invent.Object(NroItems).Amount = StolenObj.Amount
            .Invent.Object(NroItems).ObjIndex = StolenObj.ObjIndex
            .Invent.NroItems = NroItems
        End With
        Call QuitarUserInvItem(targetIndex, Slot, StolenObj.Amount)
        Call UpdateUserInv(False, targetIndex, Slot)
        ThiefStealITem = True
    End If
End Function

Private Sub ThiefTurnInvisible(ByVal NpcIndex As Integer, ByVal TurnVisible As Boolean)
    With Npclist(NpcIndex)
        If TurnVisible Then
            .flags.invisible = 1
        Else
            .flags.invisible = 0
        End If
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetInvisible(.Char.CharIndex, TurnVisible))
    End With
End Sub

Private Function TargetClose(ByRef Pos As WorldPos) As Boolean
    Dim Userindex   As Integer
    Dim nPos        As WorldPos
    Dim headingloop As Byte
    For headingloop = eHeading.NORTH To eHeading.WEST
        nPos = Pos
        Call HeadtoPos(headingloop, nPos)
        Userindex = MapData(nPos.Map, nPos.X, nPos.Y).Userindex
        If Userindex <> 0 Then
            If UserAtacable(Userindex) Then
                TargetClose = True
                Exit Function
            End If
        End If
    Next headingloop
End Function

Private Function CheckNearUserPets(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Integer
    Dim PetCounter As Long
    Dim PetIndex   As Integer
    With UserList(Userindex)
        If .NroMascotas = 0 Then Exit Function
        For PetCounter = 1 To MAXMASCOTAS
            PetIndex = .MascotasIndex(PetCounter)
            If PetIndex > 0 Then
                With Npclist(PetIndex)
                    If InVisionRangeNpc(NpcIndex, .Pos.X, .Pos.Y) Then
                        If .flags.Paralizado = 0 Then
                            CheckNearUserPets = PetIndex
                            Exit Function
                        End If
                    End If
                End With
            End If
        Next PetCounter
    End With
End Function

Private Function EsClaseMagica(ByVal Userindex As Integer) As Boolean
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        EsClaseMagica = .Clase = eClass.Mage Or .Clase = eClass.Cleric Or .Clase = eClass.Druid Or .Clase = eClass.Bard
    End With
    Exit Function
ErrorHandler:
    LogError ("Error en EsClaseMagica. Error: " & Err.Number & " - " & Err.description)
End Function

Private Sub GreedyWalkTo(ByVal NpcIndex As Integer, ByVal TargetMap As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer)
    On Error GoTo ErrorHandler
    Dim NpcX      As Integer
    Dim NpcY      As Integer
    Dim RandomDir As Integer
    With Npclist(NpcIndex).Pos
        If .Map <> TargetMap Then Exit Sub
        NpcX = .X
        NpcY = .Y
    End With
    If (NpcX = TargetX And NpcY = TargetY) Then Exit Sub
    If (NpcX > TargetX) Then
        If (NpcY < TargetY) Then
            RandomDir = RandomNumber(0, 9)
            If ((RandomDir And 1) = 0) Then
                If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                End If
            Else
                If MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                End If
            End If
        ElseIf (NpcY > TargetY) Then
            RandomDir = RandomNumber(0, 9)
            If ((RandomDir And 1) = 0) Then
                If MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                End If
            Else
                If MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                End If
            End If
        Else
            If MoveNPCChar(NpcIndex, eHeading.WEST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                Exit Sub
            Else
                Call MoveFailed(NpcIndex)
            End If
        End If
    ElseIf (NpcX < TargetX) Then
        If (NpcY < TargetY) Then
            RandomDir = RandomNumber(0, 9)
            If ((RandomDir And 1) = 0) Then
                If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                End If
            Else
                If MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                End If
            End If
        ElseIf (NpcY > TargetY) Then
            RandomDir = RandomNumber(0, 9)
            If ((RandomDir And 1) = 0) Then
                If MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                End If
            Else
                If MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                End If
            End If
        Else
            If MoveNPCChar(NpcIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                Exit Sub
            Else
                Call MoveFailed(NpcIndex)
            End If
        End If
    Else
        If (NpcY > TargetY) Then
            If MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                Exit Sub
            Else
                Call MoveFailed(NpcIndex)
            End If
        Else
            If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                Exit Sub
            Else
                Call MoveFailed(NpcIndex)
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    LogError ("Error en GreedyWalkTo. Error: " & Err.Number & " - " & Err.description)
End Sub

Private Sub MoveFailed(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        If IntervaloPermiteAtacarNpc(NpcIndex) Then
            If (RandomNumber(1, 100) > 95) Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Maldito bastardo, Ven aqui!", .Char.CharIndex, vbYellow))
            End If
        End If
    End With
End Sub

Private Sub ReturnToCenter(ByVal NpcIndex As Integer)
    Call GreedyWalkTo(NpcIndex, CenterPos.Map, CenterPos.X, CenterPos.Y)
End Sub

Private Function FarFromTeam(ByVal NpcIndex) As Boolean
On Error GoTo ErrorHandler
    If Npclist(NpcIndex).Pos.Map <> CenterPos.Map Then Exit Function
    FarFromTeam = (Abs(Npclist(NpcIndex).Pos.Y - CenterPos.Y) > MAX_DISTANCE) Or (Abs(Npclist(NpcIndex).Pos.X - CenterPos.X) > MAX_DISTANCE)
    Exit Function
ErrorHandler:
    LogError ("Error en FarFromTeam. Error: " & Err.Number & " - " & Err.description)
End Function

Private Function UserReachable(ByVal NpcIndex As Integer, ByVal Userindex As Integer) As Boolean
    On Error GoTo ErrorHandler
    If Npclist(NpcIndex).Pos.Map <> CenterPos.Map Then Exit Function
    UserReachable = (Abs(UserList(Userindex).Pos.X - CenterPos.X) < MAX_DISTANCE) And (Abs(UserList(Userindex).Pos.Y - CenterPos.Y) < MAX_DISTANCE)
    Exit Function
ErrorHandler:
    LogError ("Error en UserReachable. Error: " & Err.Number & " - " & Err.description)
End Function

Private Function InVisionRange(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal ExtendedRange As Boolean = False) As Boolean
    Dim Rango As Integer
    Rango = val(IIf(ExtendedRange, 2, 1))
    With UserList(Userindex)
        InVisionRange = Abs(.Pos.X - X) <= RANGO_VISION_X * Rango And Abs(.Pos.Y - Y) <= RANGO_VISION_Y * Rango
    End With
End Function

Private Function InVisionRangeNpc(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal ExtendedRange As Boolean = False) As Boolean
    Dim Rango As Integer
    Rango = val(IIf(ExtendedRange, 2, 1))
    With Npclist(NpcIndex)
        InVisionRangeNpc = Abs(.Pos.X - X) <= RANGO_VISION_NPC_X * Rango And Abs(.Pos.Y - Y) <= RANGO_VISION_NPC_Y * Rango
    End With
End Function

Private Function UserAtacable(ByVal Userindex As Integer, Optional ByVal CheckVisibility As Boolean = True, Optional ByVal AttackAdmin As Boolean = True) As Boolean
    With UserList(Userindex).flags
        UserAtacable = Not .EnConsulta And .AdminInvisible = 0 And .AdminPerseguible And .Muerto = 0
        If CheckVisibility Then
            UserAtacable = UserAtacable And .Oculto = 0 And .invisible = 0
        End If
        If Not AttackAdmin Then
            UserAtacable = UserAtacable And (Not EsGm(Userindex))
        End If
    End With
End Function

Private Function UserDistance(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
    With UserList(Userindex)
        UserDistance = Abs(X - .Pos.X) + Abs(Y - .Pos.Y)
    End With
End Function

Private Function NpcDistance(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
    With Npclist(NpcIndex)
        NpcDistance = Abs(X - .Pos.X) + Abs(Y - .Pos.Y)
    End With
End Function

Public Sub MuerePretoriano(ByVal NpcIndex As Integer)
    Dim PretoIndex As Integer
    For PretoIndex = 1 To NRO_PRETORIANOS
        If Pretorianos(PretoIndex).NpcIndex = NpcIndex Then
            Pretorianos(PretoIndex).NpcIndex = 0
            ClanMembersAlive = ClanMembersAlive - 1
            If ClanMembersAlive = 0 Then
                If RespawnClan Then
                    Dim NewSpawnPos As WorldPos
                    If CenterPos.X = LeftSpawnPos.X Then
                        NewSpawnPos = RightSpawnPos
                    Else
                        NewSpawnPos = LeftSpawnPos
                    End If
                    Call SpawnClan(NewSpawnPos.Map, NewSpawnPos.X, NewSpawnPos.Y, ClanIndex, True)
                Else
                    ClanActive = False
                End If
            End If
            Exit Sub
        End If
    Next PretoIndex
End Sub

Private Function AllyParalyzed(ByVal NpcIndex As Integer, ByVal ExtendedRange As Boolean) As Integer
    Dim PretoIndex As Integer
    Dim AllyIndex  As Integer
    For PretoIndex = 1 To NRO_PRETORIANOS
        AllyIndex = Pretorianos(PretoIndex).NpcIndex
        If AllyIndex <> 0 Then
            If Npclist(AllyIndex).flags.Paralizado = 1 Then
                If InVisionRangeNpc(AllyIndex, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, ExtendedRange) Then
                    AllyParalyzed = AllyIndex
                    Exit Function
                End If
            End If
        End If
    Next PretoIndex
End Function

Private Function AllyInjured(ByVal NpcIndex As Integer) As Integer
    Dim PretoIndex As Integer
    Dim AllyIndex  As Integer
    For PretoIndex = 1 To NRO_PRETORIANOS
        AllyIndex = Pretorianos(PretoIndex).NpcIndex
        If AllyIndex <> 0 Then
            With Npclist(AllyIndex)
                If .Stats.MinHp < .Stats.MaxHp Then
                    If InVisionRangeNpc(NpcIndex, .Pos.X, .Pos.Y, False) Then
                        AllyInjured = AllyIndex
                        Exit Function
                    End If
                End If
            End With
        End If
    Next PretoIndex
End Function

Private Function AllyPoisoned(ByVal NpcIndex As Integer) As Integer
    Dim PretoIndex As Integer
    Dim AllyIndex  As Integer
    For PretoIndex = 1 To NRO_PRETORIANOS
        AllyIndex = Pretorianos(PretoIndex).NpcIndex
        If AllyIndex <> 0 Then
            If Npclist(AllyIndex).flags.Envenenado = 1 Then
                If InVisionRangeNpc(AllyIndex, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
                    AllyPoisoned = AllyIndex
                    Exit Function
                End If
            End If
        End If
    Next PretoIndex
End Function

Private Function AllyDead() As Integer
    Dim PretoIndex As Integer
    Dim AllyIndex  As Integer
    For PretoIndex = 1 To NRO_PRETORIANOS
        If Pretorianos(PretoIndex).NpcIndex = 0 Then
            AllyDead = PretoIndex
            Exit Function
        End If
    Next PretoIndex
End Function

Public Function CanAtackMember(ByVal NpcIndex As Integer) As Boolean
    If Pretorianos(1).NpcIndex = NpcIndex Then
        If ClanMembersAlive <> 1 Then Exit Function
    End If
    CanAtackMember = True
End Function

Public Property Get ClanMap() As Integer
    ClanMap = CenterPos.Map
End Property

Public Property Get Active() As Boolean
    Active = ClanActive
End Property

Public Sub DeleteClan()
    Dim PretoIndex As Long
    Dim NpcIndex   As Integer
    For PretoIndex = 1 To NRO_PRETORIANOS
        NpcIndex = Pretorianos(PretoIndex).NpcIndex
        If NpcIndex <> 0 Then Call QuitarNPC(NpcIndex)
        Pretorianos(PretoIndex).NpcIndex = 0
    Next PretoIndex
    ClanActive = False
End Sub
