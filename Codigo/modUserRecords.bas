Attribute VB_Name = "modUserRecords"
Option Explicit

Public Sub LoadRecords()
    Dim Reader As clsIniManager
    Dim tmpStr As String
    Dim i      As Long
    Dim j      As Long
    Set Reader = New clsIniManager
    If Not FileExist(DatPath & "RECORDS.DAT") Then
        Call CreateRecordsFile
    End If
    Call Reader.Initialize(DatPath & "RECORDS.DAT")
    NumRecords = Reader.GetValue("INIT", "NumRecords")
    If NumRecords Then ReDim Records(1 To NumRecords)
    For i = 1 To NumRecords
        With Records(i)
            .Usuario = Reader.GetValue("RECORD" & i, "Usuario")
            .Creador = Reader.GetValue("RECORD" & i, "Creador")
            .Fecha = Reader.GetValue("RECORD" & i, "Fecha")
            .Motivo = Reader.GetValue("RECORD" & i, "Motivo")
            .NumObs = val(Reader.GetValue("RECORD" & i, "NumObs"))
            If .NumObs Then ReDim .Obs(1 To .NumObs)
            For j = 1 To .NumObs
                tmpStr = Reader.GetValue("RECORD" & i, "Obs" & j)
                .Obs(j).Creador = ReadField(1, tmpStr, 45)
                .Obs(j).Fecha = ReadField(2, tmpStr, 45)
                .Obs(j).Detalles = ReadField(3, tmpStr, 45)
            Next j
        End With
    Next i
End Sub

Public Sub SaveRecords()
    Dim Writer As clsIniManager
    Dim tmpStr As String
    Dim i      As Long
    Dim j      As Long
    Set Writer = New clsIniManager
    Call Writer.ChangeValue("INIT", "NumRecords", NumRecords)
    For i = 1 To NumRecords
        With Records(i)
            Call Writer.ChangeValue("RECORD" & i, "Usuario", .Usuario)
            Call Writer.ChangeValue("RECORD" & i, "Creador", .Creador)
            Call Writer.ChangeValue("RECORD" & i, "Fecha", .Fecha)
            Call Writer.ChangeValue("RECORD" & i, "Motivo", .Motivo)
            Call Writer.ChangeValue("RECORD" & i, "NumObs", .NumObs)
            For j = 1 To .NumObs
                tmpStr = .Obs(j).Creador & "-" & .Obs(j).Fecha & "-" & .Obs(j).Detalles
                Call Writer.ChangeValue("RECORD" & i, "Obs" & j, tmpStr)
            Next j
        End With
    Next i
    Call Writer.DumpFile(DatPath & "RECORDS.DAT")
End Sub

Public Sub AddRecord(ByVal Userindex As Integer, ByVal Nickname As String, ByVal Reason As String)
    NumRecords = NumRecords + 1
    ReDim Preserve Records(1 To NumRecords)
    With Records(NumRecords)
        .Usuario = UCase$(Nickname)
        .Fecha = Format(Now, "DD/MM/YYYY hh:mm:ss")
        .Creador = UCase$(UserList(Userindex).Name)
        .Motivo = Reason
        .NumObs = 0
    End With
End Sub

Public Sub AddObs(ByVal Userindex As Integer, ByVal RecordIndex As Integer, ByVal Obs As String)
    With Records(RecordIndex)
        .NumObs = .NumObs + 1
        ReDim Preserve .Obs(1 To .NumObs)
        .Obs(.NumObs).Creador = UCase$(UserList(Userindex).Name)
        .Obs(.NumObs).Fecha = Now
        .Obs(.NumObs).Detalles = Obs
    End With
End Sub

Public Sub RemoveRecord(ByVal RecordIndex As Integer)
    Dim i As Long
    If RecordIndex = NumRecords Then
        NumRecords = NumRecords - 1
        If NumRecords > 0 Then
            ReDim Preserve Records(1 To NumRecords)
        End If
    Else
        NumRecords = NumRecords - 1
        For i = RecordIndex To NumRecords
            Records(i) = Records(i + 1)
        Next i
        ReDim Preserve Records(1 To NumRecords)
    End If
End Sub

Public Sub CreateRecordsFile()
    Dim intFile As Integer
    intFile = FreeFile
    Open DatPath & "RECORDS.DAT" For Output As #intFile
    Print #intFile, "[INIT]"
    Print #intFile, "NumRecords=0"
    Close #intFile
End Sub
