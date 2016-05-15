VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMultiMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "UpDown EP02 MultiPlayer Debug Test Mode"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      Caption         =   "MultiPlayer 전용 공간"
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6255
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label11 
         Caption         =   "진도율 : "
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "진도율 :"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3240
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   2760
         X2              =   2760
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Label Label7 
         Caption         =   "님이 맞추신 횟수"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "nickname_c"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "님이 맞추신 횟수"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "nickname_s"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "연결 끊기(&D)"
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   10000
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock Sock_Cli 
      Left            =   3000
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2520
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   135
      Left            =   7200
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "목숨 갯수:"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmMultiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Se_Life As Integer
Dim Cl_Life As Integer
Dim Se_ANN As Integer
Dim Cl_ANN As Integer
Dim Anwser As Integer
Dim Se As String
Dim Cl As String

Public Sub ANN()
    If Val(Text1.Text) > Val(Label12.Caption) Then
        MsgBox "다운", vbInformation
        Text2.Text = Text2.Text - 1
        If Center.seorcl = "se" Then
            Se_Life = Text2.Text
            Winsock2.SendData Se_Life
        ElseIf Center.seorcl = "cl" Then
            Cl_Life = Text2.Text
            Sock_Cli.SendData Cl_Life
        End If
    ElseIf Val(Text1.Text) > Val(Label12.Caption) Then
        MsgBox "업", vbInformation
        Text2.Text = Text2.Text - 1
        If Center.seorcl = "se" Then
            Se_Life = Text2.Text
            Winsock2.SendData Se_Life
        ElseIf Center.seorcl = "cl" Then
            Cl_Life = Text2.Text
            Sock_Cli.SendData Cl_Life
    ElseIf Val(Text1.Text) = Val(Label12.Caption) Then
        MsgBox "정답입니다!", vbInformation
        If Center.seorcl = "se" Then
            Se_ANN = Se_ANN + 1
            Winsock2.SendData Se_ANN
        ElseIf Center.seorcl = "cl" Then
            Cl_ANN = Cl_ANN + 1
            Sock_Cli.SendData Cl_ANN
        Else
            MsgBox "알수 없는 값", vbCritical
            End
            End If
        End If
    End If
End Sub
Private Sub Command2_Click()
    Winsock2.Close
    Sock_Cli.Close
    Unload Me
    Form1.Show
End Sub
Public Sub Gen()

End Sub
Private Sub Form_Load()
    Randomize
    Label12.Caption = Int(Rnd * Center.Length) + 1
    If Center.seorcl = "cl" Then
        Sock_Cli.Close
        Sock_Cli.Connect Center.ip, 895
    ElseIf Center.seorcl = "se" Then
        Winsock2.Close
        Winsock2.LocalPort = 895
        Winsock2.Listen
    Else
        MsgBox "Null", vbCritical
        End
    End If
    Label1.Caption = Center.Dest
    Label2.Caption = Center.m_life
End Sub

Private Sub Sock_Cli_Close()
    MsgBox "연결에 실패했습니다. 처음으로 돌아갑니다.", vbCritical
    Center.ip = ""
    Unload Me
    frmMulti.Show
End Sub

Private Sub Sock_Cli_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "연결에 실패했습니다. 처음으로 돌아갑니다.", vbCritical
    Center.ip = ""
    Unload Me
    frmMulti.Show
End Sub

