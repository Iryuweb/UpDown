VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "로그인"
   ClientHeight    =   1845
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4245
   Icon            =   "2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1090.087
   ScaleMode       =   0  '사용자
   ScaleWidth      =   3985.826
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.OptionButton Option2 
      Caption         =   "일반 로그인"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "인증서 로그인"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "로그인"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "admin2049"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Admin"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
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


Private Sub Command1_Click()
    If Option1.Value = True Then
        On Error GoTo Err_Command1_Click
        Dim strtemp As String
        Dim FN As Integer
        FN = FreeFile
        With CommonDialog1
         .DialogTitle = "인증서 찾기"
         .Filter = "인증서파일|*.ak"
         .ShowOpen
        End With
        Open CommonDialog1.FileName For Input As #FN
        Line Input #FN, strtemp
        Label2.Caption = strtemp
        If Label2.Caption = "Username = Developer No.1" Then
        
        MsgBox "성공적으로 로그인 되었습니다!", vbInformation
        Unload Me
        Form2.Show
        Else
        MsgBox "인증서가 맞지않아 로그인에 실패하였습니다.", vbCritical
        Unload Me
        End If
    ElseIf Option2.Value = False Then
        MsgBox "보안상 존재하지 않는 페이지입니다.", vbCritical
    End If
Err_Command1_Click:
If Err.Description = "" Then
Else
MsgBox Err.Description
End If
Exit Sub
End Sub

