VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private m_sString    As String
Private m_iChunkSize As Long
Private m_iPos       As Long
Private m_iLen       As Long

Public Property Get Length() As Long
    Length = m_iPos \ 2
End Property

Public Property Get Capacity() As Long
    Capacity = m_iLen \ 2
End Property

Public Property Get ChunkSize() As Long
    ChunkSize = m_iChunkSize \ 2
End Property

Public Property Let ChunkSize(ByVal iChunkSize As Long)
    m_iChunkSize = iChunkSize * 2
End Property

Public Property Get toString() As String
    If m_iPos > 0 Then
        toString = Left$(m_sString, m_iPos \ 2)
    End If
End Property

Public Property Let TheString(ByRef sThis As String)
    Dim lLen As Long
    lLen = LenB(sThis)
    If lLen = 0 Then
        'Clear
        m_sString = ""
        m_iPos = 0
        m_iLen = 0
    Else
        If m_iLen < lLen Then
            Do
                m_sString = m_sString & Space$(m_iChunkSize \ 2)
                m_iLen = m_iLen + m_iChunkSize
            Loop While m_iLen < lLen
        End If
        CopyMemory ByVal StrPtr(m_sString), ByVal StrPtr(sThis), lLen
        m_iPos = lLen
    End If
End Property

Public Sub Clear()
    m_sString = ""
    m_iPos = 0
    m_iLen = 0
End Sub

Public Sub AppendNL(ByRef sThis As String)
    Append sThis
    Append vbCrLf
End Sub

Public Sub Append(ByRef sThis As String)
    Dim lLen        As Long
    Dim lLenPlusPos As Long
    lLen = LenB(sThis)
    lLenPlusPos = lLen + m_iPos
    If lLenPlusPos > m_iLen Then
        Dim lTemp As Long
        lTemp = m_iLen
        Do While lTemp < lLenPlusPos
            lTemp = lTemp + m_iChunkSize
        Loop
        m_sString = m_sString & Space$((lTemp - m_iLen) \ 2)
        m_iLen = lTemp
    End If
    CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
    m_iPos = m_iPos + lLen
End Sub

Public Sub AppendByVal(ByVal sThis As String)
    Append sThis
End Sub

Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)
    Dim lLen  As Long
    Dim lPos  As Long
    Dim lSize As Long
    If (iIndex * 2 > m_iPos) Then
        Err.Raise 9
    Else
        lLen = LenB(sThis)
        If (m_iPos + lLen) > m_iLen Then
            m_sString = m_sString & Space$(m_iChunkSize \ 2)
            m_iLen = m_iLen + m_iChunkSize
        End If
        lPos = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
        lSize = m_iPos - iIndex * 2
        CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
        CopyMemory ByVal lPos, ByVal StrPtr(sThis), lLen
        m_iPos = m_iPos + lLen
    End If
End Sub

Public Sub InsertByVal(ByVal iIndex As Long, ByVal sThis As String)
    Insert iIndex, sThis
End Sub

Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
    Dim lSrc  As Long
    Dim lDst  As Long
    Dim lSize As Long
    If (iIndex * 2 > m_iPos) Then
        Err.Raise 9
    Else
        If ((iIndex + lLen) * 2 > m_iPos) Then
            Err.Raise 9
        Else
            lSrc = UnsignedAdd(StrPtr(m_sString), (iIndex + lLen) * 2)
            lDst = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
            lSize = (m_iPos - (iIndex + lLen) * 2)
            CopyMemory ByVal lDst, ByVal lSrc, lSize
            m_iPos = m_iPos - lLen * 2
        End If
    End If
End Sub

Public Function Find(ByVal sToFind As String, Optional ByVal lStartIndex As Long = 1, Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
    Dim lInstr As Long
    If (lStartIndex > 0) Then
        lInstr = InStr(lStartIndex, m_sString, sToFind, compare)
    Else
        lInstr = InStr(m_sString, sToFind, compare)
    End If
    If (lInstr < m_iPos \ 2) Then
        Find = lInstr
    End If
End Function

Public Sub HeapMinimize()
    Dim iLen As Long
    If (m_iLen - m_iPos) > m_iChunkSize Then
        iLen = m_iLen
        Do While (iLen - m_iPos) > m_iChunkSize
            iLen = iLen - m_iChunkSize
        Loop
        m_sString = Left$(m_sString, iLen \ 2)
        m_iLen = iLen
    End If
End Sub

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
    If Start And &H80000000 Then
        UnsignedAdd = Start + Incr
    ElseIf (Start Or &H80000000) < -Incr Then
        UnsignedAdd = Start + Incr
    Else
        UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
    End If
End Function

Private Sub Class_Initialize()
    m_iChunkSize = 16384
End Sub
