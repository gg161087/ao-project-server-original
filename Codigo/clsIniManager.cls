VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type ChildNode
    key As String
    Value As String
End Type

Private Type MainNode
    Name As String
    values() As ChildNode
    numValues As Integer
End Type

Private fileData() As MainNode
Private MainNodes  As Long

Private Sub Class_Terminate()
    Dim i As Long
    If MainNodes Then
        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        Erase fileData
    End If
    MainNodes = 0
End Sub

Public Sub Initialize(ByVal File As String)
    Dim handle As Integer
    Dim Text   As String
    Dim Pos    As Long
    Call Class_Terminate
    handle = FreeFile
    Open File For Input As handle
    Do Until EOF(handle)
        Line Input #handle, Text
        If Len(Text) Then
            If Left$(Text, 1) = "[" Then
                Pos = InStr(2, Text, "]")
                If Pos Then
                    ReDim Preserve fileData(MainNodes) As MainNode
                    fileData(MainNodes).Name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
                    MainNodes = MainNodes + 1
                End If
            Else
                Pos = InStr(2, Text, "=")
                If Pos Then
                    If MainNodes Then
                        With fileData(MainNodes - 1)
                            ReDim Preserve .values(.numValues) As ChildNode
                            .values(.numValues).Value = Right$(Text, Len(Text) - Pos)
                            .values(.numValues).key = UCase$(Left$(Text, Pos - 1))
                            .numValues = .numValues + 1
                        End With
                    End If
                End If
            End If
        End If
    Loop
    Close handle
    Dim i As Long
    If MainNodes Then
        Call SortMainNodes(0, MainNodes - 1)
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i
    End If
End Sub

Private Sub SortChildNodes(ByRef Node As MainNode, ByVal First As Integer, ByVal Last As Integer)
    Dim min  As Integer
    Dim max  As Integer
    Dim comp As String
    Dim temp As ChildNode
    min = First
    max = Last
    With Node
        comp = .values((min + max) \ 2).key
        Do While min <= max
            Do While .values(min).key < comp And min < Last
                min = min + 1
            Loop
            Do While .values(max).key > comp And max > First
                max = max - 1
            Loop
            If min <= max Then
                temp = .values(min)
                .values(min) = .values(max)
                .values(max) = temp
                min = min + 1
                max = max - 1
            End If
        Loop
    End With
    If First < max Then SortChildNodes Node, First, max
    If min < Last Then SortChildNodes Node, min, Last
End Sub

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
    Dim min  As Integer
    Dim max  As Integer
    Dim comp As String
    Dim temp As MainNode
    min = First
    max = Last
    comp = fileData((min + max) \ 2).Name
    Do While min <= max
        Do While fileData(min).Name < comp And min < Last
            min = min + 1
        Loop
        Do While fileData(max).Name > comp And max > First
            max = max - 1
        Loop
        If min <= max Then
            temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    If First < max Then SortMainNodes First, max
    If min < Last Then SortMainNodes min, Last
End Sub

Public Function GetValue(ByVal Main As String, ByVal key As String) As String
    Dim i As Long
    Dim j As Long
    i = FindMain(UCase$(Main))
    If i >= 0 Then
        j = FindKey(fileData(i), UCase$(key))
        If j >= 0 Then GetValue = fileData(i).values(j).Value
    End If
End Function

Public Sub ChangeValue(ByVal Main As String, ByVal key As String, ByVal Value As String)
    Dim i As Long
    Dim j As Long
    i = FindMain(UCase$(Main))
    If i < 0 Then i = AddNode(Main)
    j = FindKey(fileData(i), UCase$(key))
    If j < 0 Then
        Call AddKey(Main, key, Value)
    Else
        fileData(i).values(j).Value = Value
    End If
End Sub

Private Function FindKey(ByRef Node As MainNode, ByVal key As String) As Long
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    min = 0
    max = Node.numValues - 1
    Do While min <= max
        mid = (min + max) \ 2
        If Node.values(mid).key < key Then
            min = mid + 1
        ElseIf Node.values(mid).key > key Then
            max = mid - 1
        Else
            FindKey = mid
            Exit Function
        End If
    Loop
    FindKey = Not min
End Function

Private Function FindMain(ByVal Name As String) As Long
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    min = 0
    max = MainNodes - 1
    Do While min <= max
        mid = (min + max) \ 2
        If fileData(mid).Name < Name Then
            min = mid + 1
        ElseIf fileData(mid).Name > Name Then
            max = mid - 1
        Else
            FindMain = mid
            Exit Function
        End If
    Loop
    FindMain = Not min
End Function

Public Function KeyExists(ByVal Name As String) As Boolean
    KeyExists = FindMain(UCase$(Name)) >= 0
End Function

Private Function AddNode(ByRef Name As String) As Long
    Dim i       As Long
    Dim MainPos As Long
    Name = UCase$(Name)
    ReDim Preserve fileData(MainNodes) As MainNode
    If MainNodes Then
        i = MainNodes - 1
        MainPos = Not FindMain(Name)
        Do While i >= MainPos
            fileData(i + 1) = fileData(i)
            i = i - 1
        Loop
        fileData(MainPos).numValues = 0
        Erase fileData(MainPos).values
    End If
    fileData(MainPos).Name = Name
    MainNodes = MainNodes + 1
    AddNode = MainPos
End Function

Private Function AddKey(ByRef Main As String, ByRef key As String, Optional ByRef Value As String = vbNullString) As Long
    Dim MainPos As Long
    Dim KeyPos  As Long
    Dim i       As Long
    Main = UCase$(Main)
    key = UCase$(key)
    MainPos = FindMain(Main)
    If MainPos < 0 Then MainPos = AddNode(Main)
    With fileData(MainPos)
        ReDim Preserve .values(.numValues) As ChildNode
        If .numValues > 0 Then
            i = .numValues - 1
            KeyPos = Not FindKey(fileData(MainPos), key)
            Do While i >= KeyPos
                .values(i + 1) = .values(i)
                i = i - 1
            Loop
        End If
        .values(KeyPos).key = key
        .values(KeyPos).Value = Value
        .numValues = .numValues + 1
    End With
    AddKey = KeyPos
End Function

Public Sub DumpFile(ByRef File As String)
    Dim hFile As Integer
    Dim i     As Long
    Dim j     As Long
    hFile = FreeFile()
    Open File For Output As hFile
    For i = 0 To MainNodes - 1
        With fileData(i)
            Print #hFile, "[" & .Name & "]"
            For j = 0 To .numValues - 1
                Print #hFile, .values(j).key & "=" & .values(j).Value
            Next j
            Print #hFile, ""
        End With
    Next i
    Close hFile
End Sub
