VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Base 0

Private Const NOT_ENOUGH_DATA  As Long = vbObjectError + 9
Private Const NOT_ENOUGH_SPACE As Long = vbObjectError + 10
Private Const DATA_BUFFER      As Long = 10240
Dim data()                     As Byte
Dim queueCapacity              As Long
Dim queueLength                As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal Length As Long)

Private Sub Class_Initialize()
    ReDim data(DATA_BUFFER - 1) As Byte
    queueCapacity = DATA_BUFFER
End Sub

Private Sub Class_Terminate()
    Erase data
End Sub

Public Sub CopyBuffer(ByRef source As clsByteQueue)
    If source.Length = 0 Then
        'Clear the list and exit
        Call RemoveData(Length)
        Exit Sub
    End If
    queueCapacity = source.Capacity
    ReDim data(queueCapacity - 1) As Byte
    Dim buf() As Byte
    ReDim buf(source.Length - 1) As Byte
    Call source.PeekBlock(buf, source.Length)
    queueLength = 0
    Call WriteBlock(buf, source.Length)
End Sub

Private Function min(ByVal val1 As Long, ByVal val2 As Long) As Long
    If val1 < val2 Then
        min = val1
    Else
        min = val2
    End If
End Function

Private Function WriteData(ByRef buf() As Byte, ByVal dataLength As Long) As Long
    If queueCapacity - queueLength - dataLength < 0 Then
        Call Err.Raise(NOT_ENOUGH_SPACE)
        Exit Function
    End If
    Call CopyMemory(data(queueLength), buf(0), dataLength)
    queueLength = queueLength + dataLength
    WriteData = dataLength
End Function

Private Function ReadData(ByRef buf() As Byte, ByVal dataLength As Long) As Long
    If dataLength > queueLength Then
        Call Err.Raise(NOT_ENOUGH_DATA)
        Exit Function
    End If
    Call CopyMemory(buf(0), data(0), dataLength)
    ReadData = dataLength
End Function

Private Function RemoveData(ByVal dataLength As Long) As Long
    RemoveData = min(dataLength, queueLength)
    If RemoveData <> queueCapacity Then Call CopyMemory(data(0), data(RemoveData), queueLength - RemoveData)
    queueLength = queueLength - RemoveData
End Function

Public Function WriteByte(ByVal Value As Byte) As Long
    Dim buf(0) As Byte
    buf(0) = Value
    WriteByte = WriteData(buf, 1)
End Function

Public Function WriteInteger(ByVal Value As Integer) As Long
    Dim buf(1) As Byte
    Call CopyMemory(buf(0), Value, 2)
    WriteInteger = WriteData(buf, 2)
End Function

Public Function WriteLong(ByVal Value As Long) As Long
    Dim buf(3) As Byte
    Call CopyMemory(buf(0), Value, 4)
    WriteLong = WriteData(buf, 4)
End Function

Public Function WriteSingle(ByVal Value As Single) As Long
    Dim buf(3) As Byte
    Call CopyMemory(buf(0), Value, 4)
    WriteSingle = WriteData(buf, 4)
End Function

Public Function WriteDouble(ByVal Value As Double) As Long
    Dim buf(7) As Byte
    Call CopyMemory(buf(0), Value, 8)
    WriteDouble = WriteData(buf, 8)
End Function

Public Function WriteBoolean(ByVal Value As Boolean) As Long
    Dim buf(0) As Byte
    If Value Then buf(0) = 1
    WriteBoolean = WriteData(buf, 1)
End Function

Public Function WriteASCIIStringFixed(ByVal Value As String) As Long
    Dim buf() As Byte
    ReDim buf(Len(Value) - 1) As Byte
    Call CopyMemory(buf(0), ByVal StrPtr(StrConv(Value, vbFromUnicode)), Len(Value))
    WriteASCIIStringFixed = WriteData(buf, Len(Value))
End Function

Public Function WriteUnicodeStringFixed(ByVal Value As String) As Long
    Dim buf() As Byte
    ReDim buf(LenB(Value)) As Byte
    Call CopyMemory(buf(0), ByVal StrPtr(Value), LenB(Value))
    WriteUnicodeStringFixed = WriteData(buf, LenB(Value))
End Function

Public Function WriteASCIIString(ByVal Value As String) As Long
    Dim buf() As Byte
    ReDim buf(Len(Value) + 1) As Byte
    Call CopyMemory(buf(0), CInt(Len(Value)), 2)
    If Len(Value) > 0 Then
        Call CopyMemory(buf(2), ByVal StrPtr(StrConv(Value, vbFromUnicode)), Len(Value))
    End If
    WriteASCIIString = WriteData(buf, Len(Value) + 2)
End Function

Public Function WriteUnicodeString(ByVal Value As String) As Long
    Dim buf() As Byte
    ReDim buf(LenB(Value) + 1) As Byte
    Call CopyMemory(buf(0), CInt(Len(Value)), 2)
    If Len(Value) > 0 Then
        Call CopyMemory(buf(2), ByVal StrPtr(Value), LenB(Value))
    End If
    WriteUnicodeString = WriteData(buf, LenB(Value) + 2)
End Function

Public Function WriteBlock(ByRef Value() As Byte, Optional ByVal Length As Long = -1) As Long
    If Length > UBound(Value()) + 1 Or Length < 0 Then Length = UBound(Value()) + 1
    WriteBlock = WriteData(Value, Length)
End Function

Public Function ReadByte() As Byte
    Dim buf(0) As Byte
    Call RemoveData(ReadData(buf, 1))
    ReadByte = buf(0)
End Function

Public Function ReadInteger() As Integer
    Dim buf(1) As Byte
    Call RemoveData(ReadData(buf, 2))
    Call CopyMemory(ReadInteger, buf(0), 2)
End Function

Public Function ReadLong() As Long
    Dim buf(3) As Byte
    Call RemoveData(ReadData(buf, 4))
    Call CopyMemory(ReadLong, buf(0), 4)
End Function

Public Function ReadSingle() As Single
    Dim buf(3) As Byte
    Call RemoveData(ReadData(buf, 4))
    Call CopyMemory(ReadSingle, buf(0), 4)
End Function

Public Function ReadDouble() As Double
    Dim buf(7) As Byte
    Call RemoveData(ReadData(buf, 8))
    Call CopyMemory(ReadDouble, buf(0), 8)
End Function

Public Function ReadBoolean() As Boolean
    Dim buf(0) As Byte
    Call RemoveData(ReadData(buf, 1))
    If buf(0) = 1 Then ReadBoolean = True
End Function

Public Function ReadASCIIStringFixed(ByVal Length As Long) As String
    If Length <= 0 Then Exit Function
    If queueLength >= Length Then
        Dim buf() As Byte
        ReDim buf(Length - 1) As Byte
        Call RemoveData(ReadData(buf, Length))
        ReadASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function ReadUnicodeStringFixed(ByVal Length As Long) As String
    If Length <= 0 Then Exit Function
    If queueLength >= Length * 2 Then
        Dim buf() As Byte
        ReDim buf(Length * 2 - 1) As Byte
        Call RemoveData(ReadData(buf, Length * 2))
        ReadUnicodeStringFixed = buf
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function ReadASCIIString() As String
    Dim buf(1) As Byte
    Dim Length As Integer
    If queueLength > 1 Then
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)
        If queueLength >= Length + 2 Then
            Call RemoveData(2)
            If Length > 0 Then
                Dim buf2() As Byte
                ReDim buf2(Length - 1) As Byte
                Call RemoveData(ReadData(buf2, Length))
                ReadASCIIString = StrConv(buf2, vbUnicode)
            End If
        Else
            Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function ReadUnicodeString() As String
    Dim buf(1) As Byte
    Dim Length As Integer
    If queueLength > 1 Then
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)
        If queueLength >= Length * 2 + 2 Then
            Call RemoveData(2)
            Dim buf2() As Byte
            ReDim buf2(Length * 2 - 1) As Byte
            Call RemoveData(ReadData(buf2, Length * 2))
            ReadUnicodeString = buf2
        Else
            Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function ReadBlock(ByRef block() As Byte, ByVal dataLength As Long) As Long
    If dataLength > 0 Then ReadBlock = RemoveData(ReadData(block(), dataLength))
End Function

Public Function PeekByte() As Byte
    Dim buf(0) As Byte
    Call ReadData(buf, 1)
    PeekByte = buf(0)
End Function

Public Function PeekInteger() As Integer
    Dim buf(1) As Byte
    Call ReadData(buf, 2)
    Call CopyMemory(PeekInteger, buf(0), 2)
End Function

Public Function PeekLong() As Long
    Dim buf(3) As Byte
    Call ReadData(buf, 4)
    Call CopyMemory(PeekLong, buf(0), 4)
End Function

Public Function PeekSingle() As Single
    Dim buf(3) As Byte
    Call ReadData(buf, 4)
    Call CopyMemory(PeekSingle, buf(0), 4)
End Function

Public Function PeekDouble() As Double
    Dim buf(7) As Byte
    Call ReadData(buf, 8)
    Call CopyMemory(PeekDouble, buf(0), 8)
End Function

Public Function PeekBoolean() As Boolean
    Dim buf(0) As Byte
    Call ReadData(buf, 1)
    If buf(0) = 1 Then PeekBoolean = True
End Function

Public Function PeekASCIIStringFixed(ByVal Length As Long) As String
    If Length <= 0 Then Exit Function
    If queueLength >= Length Then
        Dim buf() As Byte
        ReDim buf(Length - 1) As Byte
        Call ReadData(buf, Length)
        PeekASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekUnicodeStringFixed(ByVal Length As Long) As String
    If Length <= 0 Then Exit Function
    If queueLength >= Length * 2 Then
        Dim buf() As Byte
        ReDim buf(Length * 2 - 1) As Byte
        Call ReadData(buf, Length * 2)
        PeekUnicodeStringFixed = buf
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekASCIIString() As String
    Dim buf(1) As Byte
    Dim Length As Integer
    If queueLength > 1 Then
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)
        If queueLength >= Length + 2 Then
            Dim buf2() As Byte
            ReDim buf2(Length + 1) As Byte
            Call ReadData(buf2, Length + 2)
            If Length > 0 Then
                Dim buf3() As Byte
                ReDim buf3(Length - 1) As Byte
                Call CopyMemory(buf3(0), buf2(2), Length)
                PeekASCIIString = StrConv(buf3, vbUnicode)
            End If
        Else
            Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekUnicodeString() As String
    Dim buf(1) As Byte
    Dim Length As Integer
    If queueLength > 1 Then
        Call ReadData(buf, 2)
        Call CopyMemory(Length, buf(0), 2)
        If queueLength >= Length * 2 + 2 Then
            Dim buf2() As Byte
            ReDim buf2(Length * 2 + 1) As Byte
            Call ReadData(buf2, Length * 2 + 2)
            Dim buf3() As Byte
            ReDim buf3(Length * 2 - 1) As Byte
            Call CopyMemory(buf3(0), buf2(2), Length * 2)
            PeekUnicodeString = buf3
        Else
            Call Err.Raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.Raise(NOT_ENOUGH_DATA)
    End If
End Function

Public Function PeekBlock(ByRef block() As Byte, ByVal dataLength As Long) As Long
    If dataLength > 0 Then PeekBlock = ReadData(block(), dataLength)
End Function

Public Property Get Capacity() As Long
    Capacity = queueCapacity
End Property

Public Property Let Capacity(ByVal Value As Long)
    queueCapacity = Value
    If Length > Value Then queueLength = Value
    ReDim Preserve data(queueCapacity - 1) As Byte
End Property

Public Property Get Length() As Long
    Length = queueLength
End Property

Public Property Get NotEnoughDataErrCode() As Long
    NotEnoughDataErrCode = NOT_ENOUGH_DATA
End Property

Public Property Get NotEnoughSpaceErrCode() As Long
    NotEnoughSpaceErrCode = NOT_ENOUGH_SPACE
End Property
