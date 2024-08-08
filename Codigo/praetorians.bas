Attribute VB_Name = "PraetoriansCoopNPC"
Public MAPA_PRETORIANO           As Integer
Public PRETORIANO_X              As Byte
Public PRETORIANO_Y              As Byte
Public PretorianAIOffset(1 To 7) As Integer
Public PretorianDatNumbers()     As Integer
Public Const SONIDO_DRAGON_VIVO  As Integer = 30

Public Enum ePretorianAI
    King = 1
    Healer
    SpellCaster
    SwordMaster
    Shooter
    Thief
    Last
End Enum

Public Sub LoadPretorianData()
    Dim PretorianDat As String
    PretorianDat = DatPath & "Pretorianos.dat"
    Dim NroCombinaciones As Integer
    NroCombinaciones = val(GetVar(PretorianDat, "MAIN", "Combinaciones"))
    ReDim PretorianDatNumbers(1 To NroCombinaciones)
    Dim TempInt        As Integer
    Dim Counter        As Long
    Dim PretorianIndex As Integer
    PretorianIndex = 1
    TempInt = val(GetVar(PretorianDat, "KING", "Cantidad"))
    PretorianAIOffset(ePretorianAI.King) = 1
    For Counter = 1 To TempInt
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "KING", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "KING", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1
    Next Counter
    TempInt = val(GetVar(PretorianDat, "HEALER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Healer) = PretorianIndex
    For Counter = 1 To TempInt
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "HEALER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "HEALER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1
    Next Counter
    TempInt = val(GetVar(PretorianDat, "SPELLCASTER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.SpellCaster) = PretorianIndex
    For Counter = 1 To TempInt
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SPELLCASTER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SPELLCASTER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1
    Next Counter
    TempInt = val(GetVar(PretorianDat, "SWORDSWINGER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.SwordMaster) = PretorianIndex
    For Counter = 1 To TempInt
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SWORDSWINGER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SWORDSWINGER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1
    Next Counter
    TempInt = val(GetVar(PretorianDat, "LONGRANGE", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Shooter) = PretorianIndex
    For Counter = 1 To TempInt
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "LONGRANGE", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "LONGRANGE", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1
    Next Counter
    TempInt = val(GetVar(PretorianDat, "THIEF", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Thief) = PretorianIndex
    For Counter = 1 To TempInt
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "THIEF", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "THIEF", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1
    Next Counter
    PretorianAIOffset(ePretorianAI.Last) = PretorianIndex
    ReDim ClanPretoriano(ePretorianType.Default To ePretorianType.Custom) As clsClanPretoriano
    Set ClanPretoriano(ePretorianType.Default) = New clsClanPretoriano
    Set ClanPretoriano(ePretorianType.Custom) = New clsClanPretoriano
End Sub

