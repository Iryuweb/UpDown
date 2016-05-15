VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Developer Tools v2.0"
   ClientHeight    =   3720
   ClientLeft      =   5295
   ClientTop       =   5385
   ClientWidth     =   8175
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command14 
      Caption         =   "Road Scenario"
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "오답 처리(다운)"
      Height          =   495
      Left            =   4920
      TabIndex        =   26
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "오답 처리(업)"
      Height          =   495
      Left            =   4920
      TabIndex        =   25
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "정답 처리"
      Height          =   495
      Left            =   4920
      TabIndex        =   24
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "확인"
      Height          =   300
      Left            =   3000
      TabIndex        =   23
      Top             =   2505
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1440
      TabIndex        =   22
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "확인"
      Height          =   300
      Left            =   3000
      TabIndex        =   19
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1440
      TabIndex        =   18
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "새로고침"
      Height          =   495
      Left            =   6480
      TabIndex        =   16
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "확인"
      Height          =   300
      Left            =   3000
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "목숨 무한"
      Height          =   495
      Left            =   6480
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "확인"
      Height          =   300
      Left            =   3000
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1440
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "목숨 1개 제거"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "목숨 1개 추가"
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "프로그램 종료"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "문제 넘어가기"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "퍼센트 조절 : "
      Height          =   255
      Left            =   300
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   1455
      Left            =   11880
      TabIndex        =   20
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "정답 조절 : "
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "%%%"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "현재의 범위 : "
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "범위 조절 :"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "목숨 조절 : "
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "%%%"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "기반 : Up and Down Expansion Pack No.2 - 3.x - Code Name : Blue Ocean"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label3 
      Caption         =   "현재의 답 : "
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Randomize
frmMain.labe.Caption = Int(Rnd * frmMain.Label7.Caption) + 1
Me.Label2.Caption = frmMain.labe.Caption
End Sub

Private Sub Command10_Click()
    If 100 < Val(Text4.Text) Then
        MsgBox "100(%) 이하의 값을 하십시오.", vbCritical
    Else
        frmMain.ProgressBar1.Value = Me.Text4.Text
        frmMain.Label9.Caption = Me.Text4.Text / 100 * 300
    End If
End Sub

Private Sub Command11_Click()
    frmMain.D.Text = frmMain.labe.Caption
    frmMain.Answerchk
End Sub

Private Sub Command12_Click()
    frmMain.D.Text = frmMain.labe.Caption - 1
    frmMain.Answerchk
End Sub

Private Sub Command13_Click()
    frmMain.D.Text = frmMain.labe.Caption + 1
    frmMain.Answerchk
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
frmMain.Text1.Text = frmMain.Text1.Text + 1
End Sub



Private Sub Command4_Click()
frmMain.Text1.Text = frmMain.Text1.Text - 1
End Sub

Private Sub Command5_Click()
frmMain.Text1.Text = Text1.Text

End Sub

Private Sub Command6_Click()
Label2.Caption = frmMain.labe.Caption
Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command7_Click()
frmMain.Label7.Caption = Text2.Text
Randomize
frmMain.labe.Caption = Int(Rnd * frmMain.Label7.Caption) + 1
Me.Label2.Caption = frmMain.labe.Caption
Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command8_Click()
frmMain.Text1.Text = "1E+124"
End Sub

Private Sub Command9_Click()
frmMain.labe.Caption = Me.Text3.Text
End Sub

Private Sub Form_Load()
Label2.Caption = frmMain.labe
End Sub

