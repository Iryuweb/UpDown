VERSION 5.00
Begin VB.Form frmMulti 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Up And Down Expansion Pack 2 MultiPlayer Debug Mode"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8760
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      Caption         =   "방 참가"
      Height          =   2535
      Left            =   4440
      TabIndex        =   6
      Top             =   720
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "생성"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "이름 : "
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "방 참가"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "참가"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "이름 : "
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "아이피 입력 : "
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Caption         =   "멀티플레이어 시험중입니다."
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "아이피를 입력하세요.", vbExclamation
        If Text2.Text = "" Then
            MsgBox "이름을 입력하세요.", vbExclamation
        End If
    Else
        If Text2.Text = "" Then
            MsgBox "이름을 입력하세요.", vbExclamation
        Else
        Center.seorcl = "cl"
        Center.cli_nickname = Text2.Text
        Center.ip = Text1.Text
        Unload Me
        frmMultiMain.Show
        End If
    End If
End Sub

Private Sub Command2_Click()
    Center.Dest = InputBox("목표를 입력하세요:", "생성을 위한 조건들", 0)
    If Val(Center.Dest) < 0 Then
        MsgBox "목표는 0 이하가 될수 없습니다. 1로 조정합니다.", vbExclamation
        Center.Dest = 1
    End If
    Center.m_life = InputBox("목숨갯수를 입력하세요:", "생성을 위한 조건들", 0)
    If Val(Center.m_life) < 0 Then
        MsgBox "목숨은 0 이하가 될수 없습니다. 1로 조정합니다.", vbExclamation
        Center.m_life = 1
    End If
    Center.Length = InputBox("범위를 입력하세요:", "생성을 위한 조건들", 0)
    If Val(Center.Length) < 0 Then
        MsgBox "범위는 0 이하가 될수 없습니다. 1로 조정합니다.", vbExclamation
        Center.Length = 1
    Else
        Center.seorcl = "se"
        Center.ser_nickname = Text3.Text
        Unload Me
        frmMultiMain.Show
    End If
End Sub

