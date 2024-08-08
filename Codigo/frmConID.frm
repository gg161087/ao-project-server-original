VERSION 5.00
Begin VB.Form frmConID 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ConID"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Liberar todos los slots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3495
      Width           =   4290
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ver estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3030
      Width           =   4290
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3975
      Width           =   4290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   180
      TabIndex        =   4
      Top             =   2430
      Width           =   4230
   End
End
Attribute VB_Name = "frmConID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    List1.Clear
    Dim c As Integer
    Dim i As Integer
    For i = 1 To MaxUsers
        List1.AddItem "UserIndex " & i & " -- " & UserList(i).ConnID
        If UserList(i).ConnID <> -1 Then c = c + 1
    Next i
    If c = MaxUsers Then
        Label1.Caption = "No hay slots vacios!"
    Else
        Label1.Caption = "Hay " & MaxUsers - c & " slots vacios!"
    End If
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida And Not UserList(i).flags.UserLogged Then Call CloseSocket(i)
    Next i
End Sub

