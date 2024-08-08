VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FRENTE = 1
Private MAXELEM As Integer
Private Cola    As Collection

Public Property Let MaxLenght(ByVal max As Integer)
    MAXELEM = max
End Property

Public Sub Reset()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To Me.Longitud
        Cola.Remove FRENTE
    Next i
End Sub

Public Property Get Longitud() As Integer
    Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
    IndexValido = i >= 1 And i <= Me.Longitud
End Function

Private Sub Class_Initialize()
    Set Cola = New Collection
End Sub

Public Function VerElemento(ByVal index As Integer, Optional ByVal ConvertUpper As Boolean = True) As String
    On Error Resume Next
    If IndexValido(index) Then
        If ConvertUpper Then
            VerElemento = UCase$(Cola.Item(index))
        Else
            VerElemento = Cola.Item(index)
        End If
    Else
        VerElemento = 0
    End If
End Function

Public Sub Push(ByVal Detalle As String, Optional ByVal ConvertUpper As Boolean = True)
    On Error Resume Next
    Dim aux As String
    aux = time$ & " "
    If ConvertUpper Then
        aux = aux & UCase$(Detalle)
    Else
        aux = aux & Detalle
    End If
    If MAXELEM <> 0 Then
        If Cola.Count = MAXELEM Then
            Call Cola.Remove(FRENTE)
        End If
    End If
    Call Cola.Add(aux)
End Sub

Public Function Pop() As String
    On Error Resume Next
    If Cola.Count > 0 Then
        Pop = Cola(FRENTE)
        Call Cola.Remove(FRENTE)
    Else
        Pop = 0
    End If
End Function

Public Function PopByVal() As String
    On Error Resume Next
    If Cola.Count > 0 Then
        PopByVal = Cola.Item(1)
    Else
        PopByVal = 0
    End If
End Function

Public Function Existe(ByVal Nombre As String) As Boolean
    On Error Resume Next
    Dim V                  As String
    Dim i                  As Integer
    Dim NombreEnMayusculas As String
    NombreEnMayusculas = UCase$(Nombre)
    For i = 1 To Me.Longitud
        V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
        If V = NombreEnMayusculas Then
            Existe = True
            Exit Function
        End If
    Next
    Existe = False
End Function

Public Sub Quitar(ByVal Nombre As String)
    On Error Resume Next
    Dim V                  As String
    Dim i                  As Integer
    Dim NombreEnMayusculas As String
    NombreEnMayusculas = UCase$(Nombre)
    For i = 1 To Me.Longitud
        V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
        If V = NombreEnMayusculas Then
            Call Cola.Remove(i)
            Exit Sub
        End If
    Next i
End Sub

Public Sub QuitarIndex(ByVal index As Integer)
    On Error Resume Next
    If IndexValido(index) Then Call Cola.Remove(index)
End Sub

Private Sub Class_Terminate()
    Set Cola = Nothing
End Sub
