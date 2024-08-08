Attribute VB_Name = "Characters"
Option Explicit

Public Const INVALID_INDEX As Integer = 0

Public Function CharIndexToUserIndex(ByVal CharIndex As Integer) As Integer
    CharIndexToUserIndex = CharList(CharIndex)
    If CharIndexToUserIndex < 1 Or CharIndexToUserIndex > MaxUsers Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
    If UserList(CharIndexToUserIndex).Char.CharIndex <> CharIndex Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
End Function
