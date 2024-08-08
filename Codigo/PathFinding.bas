Attribute VB_Name = "PathFinding"
Option Explicit

Private Const rows   As Integer = 100
Private Const COLUMS As Integer = 100
Private Const MAXINT As Integer = 1000
Private TmpArray(1 To rows, 1 To COLUMS) As tIntermidiateWork

Private Type tIntermidiateWork
    DistV As Integer
    PrevV As tVertice
End Type

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer) As Boolean
    Limites = ((vcolu >= 1) And (vcolu <= COLUMS) And (vfila >= 1) And (vfila <= rows))
End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
    With MapData(Map, row, Col)
        IsWalkable = ((.Blocked Or .NpcIndex) = 0)
        If .Userindex <> 0 Then
            If .Userindex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False
        End If
    End With
End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef t() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
    Dim V As tVertice
    Dim j As Integer
    j = vfila - 1
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
            With t(j, vcolu)
                If .DistV = MAXINT Then
                    .DistV = t(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    V.X = vcolu
                    V.Y = j
                    Call Push(V)
                End If
            End With
        End If
    End If
    j = vfila + 1
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
            With t(j, vcolu)
                If .DistV = MAXINT Then
                    .DistV = t(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    V.X = vcolu
                    V.Y = j
                    Call Push(V)
                End If
            End With
        End If
    End If
    j = vcolu - 1
    If Limites(vfila, j) Then
        If IsWalkable(MapIndex, vfila, j, NpcIndex) Then
            With t(vfila, j)
                If .DistV = MAXINT Then
                    .DistV = t(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    V.X = j
                    V.Y = vfila
                    Call Push(V)
                End If
            End With
        End If
    End If
    j = vcolu + 1
    If Limites(vfila, j) Then
        If IsWalkable(MapIndex, vfila, j, NpcIndex) Then
            With t(vfila, j)
                If .DistV = MAXINT Then
                    .DistV = t(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    V.X = j
                    V.Y = vfila
                    Call Push(V)
                End If
            End With
        End If
    End If
End Sub

Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
    Dim cur_npc_pos As tVertice
    Dim tar_npc_pos As tVertice
    Dim V           As tVertice
    Dim NpcMap      As Integer
    Dim steps       As Integer
    With Npclist(NpcIndex)
        NpcMap = .Pos.Map
        cur_npc_pos.X = .Pos.Y
        cur_npc_pos.Y = .Pos.X
        tar_npc_pos.X = .PFINFO.Target.X
        tar_npc_pos.Y = .PFINFO.Target.Y
        Call InitializeTable(TmpArray, cur_npc_pos)
        Call InitQueue
        Call Push(cur_npc_pos)
        Do While (Not IsEmpty)
            If steps > MaxSteps Then Exit Do
            V = Pop
            If (V.X = tar_npc_pos.X) And (V.Y = tar_npc_pos.Y) Then Exit Do
            Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
        Loop
        Call MakePath(NpcIndex)
    End With
End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
    Dim Pasos As Integer
    Dim miV   As tVertice
    Dim i     As Integer
        With Npclist(NpcIndex)
        Pasos = TmpArray(.PFINFO.Target.Y, .PFINFO.Target.X).DistV
        .PFINFO.PathLenght = Pasos
        If Pasos = MAXINT Then
            .PFINFO.NoPath = True
            .PFINFO.PathLenght = 0
            Exit Sub
        End If
        ReDim .PFINFO.Path(1 To Pasos) As tVertice
        miV.X = .PFINFO.Target.X
        miV.Y = .PFINFO.Target.Y
        For i = Pasos To 1 Step -1
            .PFINFO.Path(i) = miV
            miV = TmpArray(miV.Y, miV.X).PrevV
        Next i
        .PFINFO.CurPos = 1
        .PFINFO.NoPath = False
    End With
End Sub

Private Sub InitializeTable(ByRef t() As tIntermidiateWork, ByRef S As tVertice, Optional ByVal MaxSteps As Integer = 30)
    Dim j As Integer, K As Integer
    Const anymap = 1
    For j = S.Y - MaxSteps To S.Y + MaxSteps
        For K = S.X - MaxSteps To S.X + MaxSteps
            If InMapBounds(anymap, j, K) Then
                With t(j, K)
                    .DistV = MAXINT
                    .PrevV.X = 0
                    .PrevV.Y = 0
                End With
            End If
        Next K
    Next j
    t(S.Y, S.X).DistV = 0
End Sub
