Attribute VB_Name = "modForum"
Option Explicit

Public Const MAX_MENSAJES_FORO As Byte = 30
Public Const MAX_ANUNCIOS_FORO As Byte = 5
Public Const FORO_REAL_ID      As String = "REAL"
Public Const FORO_CAOS_ID      As String = "CAOS"
Private NumForos As Integer
Private Foros()  As tForo

Public Type tPost
    sTitulo As String
    sPost As String
    Autor As String
End Type

Public Type tForo
    vsPost(1 To MAX_MENSAJES_FORO) As tPost
    vsAnuncio(1 To MAX_ANUNCIOS_FORO) As tPost
    CantPosts As Byte
    CantAnuncios As Byte
    ID As String
End Type

Public Sub AddForum(ByVal sForoID As String)
    Dim ForumPath As String
    Dim PostPath  As String
    Dim PostIndex As Integer
    Dim FileIndex As Integer
    NumForos = NumForos + 1
    ReDim Preserve Foros(1 To NumForos) As tForo
    ForumPath = App.Path & "\foros\" & sForoID & ".for"
    With Foros(NumForos)
        .ID = sForoID
        If FileExist(ForumPath, vbNormal) Then
            .CantPosts = val(GetVar(ForumPath, "INFO", "CantMSG"))
            .CantAnuncios = val(GetVar(ForumPath, "INFO", "CantAnuncios"))
            For PostIndex = 1 To .CantPosts
                FileIndex = FreeFile
                PostPath = App.Path & "\foros\" & sForoID & PostIndex & ".for"
                Open PostPath For Input Shared As #FileIndex
                Input #FileIndex, .vsPost(PostIndex).sTitulo
                Input #FileIndex, .vsPost(PostIndex).Autor
                Input #FileIndex, .vsPost(PostIndex).sPost
                Close #FileIndex
            Next PostIndex
            For PostIndex = 1 To .CantAnuncios
                FileIndex = FreeFile
                PostPath = App.Path & "\foros\" & sForoID & PostIndex & "a.for"
                Open PostPath For Input Shared As #FileIndex
                Input #FileIndex, .vsAnuncio(PostIndex).sTitulo
                Input #FileIndex, .vsAnuncio(PostIndex).Autor
                Input #FileIndex, .vsAnuncio(PostIndex).sPost
                Close #FileIndex
            Next PostIndex
        End If
    End With
End Sub

Public Function GetForumIndex(ByRef sForoID As String) As Integer
    Dim ForumIndex As Integer
    For ForumIndex = 1 To NumForos
        If Foros(ForumIndex).ID = sForoID Then
            GetForumIndex = ForumIndex
            Exit Function
        End If
    Next ForumIndex
End Function

Public Sub AddPost(ByVal ForumIndex As Integer, ByRef Post As String, ByRef Autor As String, ByRef Titulo As String, ByVal bAnuncio As Boolean)
    With Foros(ForumIndex)
        If bAnuncio Then
            If .CantAnuncios < MAX_ANUNCIOS_FORO Then .CantAnuncios = .CantAnuncios + 1
            Call MoveArray(ForumIndex, bAnuncio)
            With .vsAnuncio(1)
                .sTitulo = Titulo
                .Autor = Autor
                .sPost = Post
            End With
        Else
            If .CantPosts < MAX_MENSAJES_FORO Then .CantPosts = .CantPosts + 1
            Call MoveArray(ForumIndex, bAnuncio)
            With .vsPost(1)
                .sTitulo = Titulo
                .Autor = Autor
                .sPost = Post
            End With
        End If
    End With
End Sub

Public Sub SaveForums()
    Dim ForumIndex As Integer
    For ForumIndex = 1 To NumForos
        Call SaveForum(ForumIndex)
    Next ForumIndex
End Sub

Private Sub SaveForum(ByVal ForumIndex As Integer)
    Dim PostIndex As Integer
    Dim FileIndex As Integer
    Dim PostPath  As String
    Call CleanForum(ForumIndex)
    With Foros(ForumIndex)
        Call WriteVar(App.Path & "\Foros\" & .ID & ".for", "INFO", "CantMSG", .CantPosts)
        Call WriteVar(App.Path & "\Foros\" & .ID & ".for", "INFO", "CantAnuncios", .CantAnuncios)
        For PostIndex = 1 To .CantPosts
            PostPath = App.Path & "\Foros\" & .ID & PostIndex & ".for"
            FileIndex = FreeFile()
            Open PostPath For Output As FileIndex
            With .vsPost(PostIndex)
                Print #FileIndex, .sTitulo
                Print #FileIndex, .Autor
                Print #FileIndex, .sPost
            End With
            Close #FileIndex
        Next PostIndex
        For PostIndex = 1 To .CantAnuncios
            PostPath = App.Path & "\Foros\" & .ID & PostIndex & "a.for"
            FileIndex = FreeFile()
            Open PostPath For Output As FileIndex
            With .vsAnuncio(PostIndex)
                Print #FileIndex, .sTitulo
                Print #FileIndex, .Autor
                Print #FileIndex, .sPost
            End With
            Close #FileIndex
        Next PostIndex
    End With
End Sub

Public Sub CleanForum(ByVal ForumIndex As Integer)
    Dim PostIndex As Integer
    Dim NumPost   As Integer
    Dim ForumPath As String
    With Foros(ForumIndex)
        ForumPath = App.Path & "\Foros\" & .ID & ".for"
        If FileExist(ForumPath, vbNormal) Then
            NumPost = val(GetVar(ForumPath, "INFO", "CantMSG"))
            For PostIndex = 1 To NumPost
                Kill App.Path & "\Foros\" & .ID & PostIndex & ".for"
            Next PostIndex
            NumPost = val(GetVar(ForumPath, "INFO", "CantAnuncios"))
            For PostIndex = 1 To NumPost
                Kill App.Path & "\Foros\" & .ID & PostIndex & "a.for"
            Next PostIndex
            Kill App.Path & "\Foros\" & .ID & ".for"
        End If
    End With
End Sub

Public Function SendPosts(ByVal Userindex As Integer, ByRef ForoID As String) As Boolean
    Dim ForumIndex As Integer
    Dim PostIndex  As Integer
    Dim bEsGm      As Boolean
    ForumIndex = GetForumIndex(ForoID)
    If ForumIndex > 0 Then
        With Foros(ForumIndex)
            For PostIndex = 1 To .CantPosts
                With .vsPost(PostIndex)
                    Call WriteAddForumMsg(Userindex, eForumMsgType.ieGeneral, .sTitulo, .Autor, .sPost)
                End With
            Next PostIndex
            For PostIndex = 1 To .CantAnuncios
                With .vsAnuncio(PostIndex)
                    Call WriteAddForumMsg(Userindex, eForumMsgType.ieGENERAL_STICKY, .sTitulo, .Autor, .sPost)
                End With
            Next PostIndex
        End With
        bEsGm = EsGm(Userindex)
        If esCaos(Userindex) Or bEsGm Then
            ForumIndex = GetForumIndex(FORO_CAOS_ID)
            With Foros(ForumIndex)
                For PostIndex = 1 To .CantPosts
                    With .vsPost(PostIndex)
                        Call WriteAddForumMsg(Userindex, eForumMsgType.ieCAOS, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
                For PostIndex = 1 To .CantAnuncios
                    With .vsAnuncio(PostIndex)
                        Call WriteAddForumMsg(Userindex, eForumMsgType.ieCAOS_STICKY, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
            End With
        End If
        If esArmada(Userindex) Or bEsGm Then
            ForumIndex = GetForumIndex(FORO_REAL_ID)
            With Foros(ForumIndex)
                For PostIndex = 1 To .CantPosts
                    With .vsPost(PostIndex)
                        Call WriteAddForumMsg(Userindex, eForumMsgType.ieREAL, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
                For PostIndex = 1 To .CantAnuncios
                    With .vsAnuncio(PostIndex)
                        Call WriteAddForumMsg(Userindex, eForumMsgType.ieREAL_STICKY, .sTitulo, .Autor, .sPost)
                    End With
                Next PostIndex
            End With
        End If
        SendPosts = True
    End If
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
    End Select
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
    End Select
End Function

Public Sub ResetForums()
    ReDim Foros(1 To 1) As tForo
    NumForos = 0
End Sub

Private Sub MoveArray(ByVal ForumIndex As Integer, ByVal Sticky As Boolean)
    Dim i As Long
    With Foros(ForumIndex)
        If Sticky Then
            For i = .CantAnuncios To 2 Step -1
                .vsAnuncio(i).sTitulo = .vsAnuncio(i - 1).sTitulo
                .vsAnuncio(i).sPost = .vsAnuncio(i - 1).sPost
                .vsAnuncio(i).Autor = .vsAnuncio(i - 1).Autor
            Next i
        Else
            For i = .CantPosts To 2 Step -1
                .vsPost(i).sTitulo = .vsPost(i - 1).sTitulo
                .vsPost(i).sPost = .vsPost(i - 1).sPost
                .vsPost(i).Autor = .vsPost(i - 1).Autor
            Next i
        End If
    End With
End Sub
