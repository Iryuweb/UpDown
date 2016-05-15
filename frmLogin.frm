VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "로그인"
   ClientHeight    =   1845
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3885
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1090.087
   ScaleMode       =   0  '사용자
   ScaleWidth      =   3647.805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "인증서 로그인"
      Height          =   735
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1529
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   390
      Left            =   1320
      TabIndex        =   5
      Top             =   1260
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  '사용 못함
      Left            =   1529
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "사용자 이름(&U):"
      Height          =   375
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Caption         =   "암호(&P):"
      Height          =   270
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   660
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Public login As Boolean


Private Sub cmdCancel_Click()
    '전역 변수를 False로 설정하여
    '실패한 로드인을 표시합니다.
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
If txtUserName.Text = "Admin" Then
Else
MsgBox "아이디가 다릅니다."
End If
If txtPassword.Text = "admin2049" Then
MsgBox "성공적으로 로그인되었습니다!"
LoginSucceeded = True
Unload Me
Form2.Show

Else
MsgBox "아이디 또는 비밀번호가 틀립니다."
End If

End Sub


Private Sub Command1_Click()
On Error GoTo Err_Command1_Click
Dim strtemp As String
Dim FN As Integer
FN = FreeFile
With CommonDialog1
 .DialogTitle = "인증서 찾기"
 .Filter = "인증서파일|*.ak"
 .ShowOpen
Open CommonDialog1.FileName For Input As #FN
Line Input #FN, strtemp
txtUserName.Text = strtemp
If txtUserName.Text = "Username = Developer No.1" Then
MsgBox "성공적으로 로그인 되었습니다!", vbInformation
Unload Me
Form2.Show
Else
MsgBox "인증서가 맞지않아 로그인에 실패하였습니다.", vbCritical
Unload Me
End If
End With
Err_Command1_Click:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
If Form2.ShowInTaskbar Then
 MsgBox "이미 켜져있습니다.", vbCritical
End If

End Sub
