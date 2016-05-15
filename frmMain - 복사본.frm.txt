VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "UpDown 3.1.2 , Name : Blue Ocean (EP2)"
   ClientHeight    =   2145
   ClientLeft      =   5070
   ClientTop       =   2475
   ClientWidth     =   6645
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6645
   Begin VB.Frame Frame3 
      Caption         =   "MultiPlay"
      Height          =   3375
      Left            =   840
      TabIndex        =   44
      Top             =   4920
      Width           =   6375
      Begin VB.ListBox List3 
         Height          =   1320
         Left            =   360
         TabIndex        =   54
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "전송"
         Height          =   255
         Left            =   5640
         TabIndex        =   52
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2520
         TabIndex        =   51
         Top             =   2880
         Width           =   3135
      End
      Begin VB.ListBox List2 
         Height          =   2220
         Left            =   2520
         TabIndex        =   49
         Top             =   600
         Width           =   3615
      End
      Begin VB.ListBox List1 
         Height          =   780
         Left            =   360
         TabIndex        =   48
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label22 
         Caption         =   "게임 상황 : "
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label21 
         Caption         =   "채팅 : "
         Height          =   255
         Left            =   2520
         TabIndex        =   50
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label20 
         Caption         =   "참가자 목록:"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   360
         Width           =   4575
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10200
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "정보"
      Height          =   375
      Left            =   4800
      TabIndex        =   28
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1440
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   480
      Width           =   1140
   End
   Begin VB.CommandButton Command22 
      Caption         =   "난이도 변경(부활)"
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command21 
      Caption         =   "마 지 막 목숨"
      Height          =   420
      Left            =   10260
      TabIndex        =   20
      Top             =   10980
      Width           =   3975
   End
   Begin VB.CommandButton Command20 
      Caption         =   "목숨"
      Height          =   555
      Left            =   12480
      TabIndex        =   19
      Top             =   7680
      Width           =   1005
   End
   Begin VB.CommandButton Command19 
      Caption         =   "목숨"
      Height          =   555
      Left            =   13080
      TabIndex        =   18
      Top             =   9480
      Width           =   1005
   End
   Begin VB.CommandButton Command18 
      Caption         =   "목숨"
      Height          =   555
      Left            =   13080
      TabIndex        =   17
      Top             =   7860
      Width           =   1005
   End
   Begin VB.CommandButton Command17 
      Caption         =   "목숨"
      Height          =   555
      Left            =   14070
      TabIndex        =   16
      Top             =   7860
      Width           =   1005
   End
   Begin VB.CommandButton Command16 
      Caption         =   "목숨"
      Height          =   555
      Left            =   11100
      TabIndex        =   15
      Top             =   8400
      Width           =   1005
   End
   Begin VB.CommandButton Command15 
      Caption         =   "목숨"
      Height          =   555
      Left            =   13200
      TabIndex        =   14
      Top             =   8040
      Width           =   1005
   End
   Begin VB.CommandButton Command14 
      Caption         =   "목숨"
      Height          =   555
      Left            =   13080
      TabIndex        =   13
      Top             =   8400
      Width           =   1005
   End
   Begin VB.CommandButton Command13 
      Caption         =   "목숨"
      Height          =   555
      Left            =   14070
      TabIndex        =   12
      Top             =   8400
      Width           =   1005
   End
   Begin VB.CommandButton Command12 
      Caption         =   "목숨"
      Height          =   555
      Left            =   11100
      TabIndex        =   11
      Top             =   8940
      Width           =   1005
   End
   Begin VB.CommandButton Command11 
      Caption         =   "목숨"
      Height          =   555
      Left            =   14160
      TabIndex        =   10
      Top             =   7800
      Width           =   1005
   End
   Begin VB.CommandButton Command10 
      Caption         =   "목숨"
      Height          =   555
      Left            =   13080
      TabIndex        =   9
      Top             =   8940
      Width           =   1005
   End
   Begin VB.CommandButton Command9 
      Caption         =   "목숨"
      Height          =   555
      Left            =   14070
      TabIndex        =   8
      Top             =   8940
      Width           =   1005
   End
   Begin VB.CommandButton Command8 
      Caption         =   "목숨"
      Height          =   555
      Left            =   11100
      TabIndex        =   7
      Top             =   9480
      Width           =   1005
   End
   Begin VB.CommandButton Command7 
      Caption         =   "목숨"
      Height          =   555
      Left            =   9240
      TabIndex        =   6
      Top             =   11160
      Width           =   1005
   End
   Begin VB.CommandButton Command6 
      Caption         =   "목숨"
      Height          =   555
      Left            =   14280
      TabIndex        =   5
      Top             =   8280
      Width           =   1005
   End
   Begin VB.CommandButton Command5 
      Caption         =   "목숨"
      Height          =   555
      Left            =   11100
      TabIndex        =   4
      Top             =   7860
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "다시하기"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox D 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "미사용"
      Height          =   3495
      Left            =   10440
      TabIndex        =   39
      Top             =   7200
      Width           =   6135
   End
   Begin VB.Label Label23 
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label Label19 
      Caption         =   "0"
      Height          =   255
      Left            =   9600
      TabIndex        =   46
      Top             =   6600
      Width           =   4455
   End
   Begin VB.Label Label18 
      Caption         =   "멀티플 활성화 :"
      Height          =   255
      Left            =   8280
      TabIndex        =   45
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   255
      Left            =   9600
      TabIndex        =   43
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label Label16 
      Caption         =   "시나리오 단계 : "
      Height          =   255
      Left            =   8280
      TabIndex        =   42
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "300"
      Height          =   255
      Left            =   9600
      TabIndex        =   41
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label Label14 
      Caption         =   "시나리오 목표 : "
      Height          =   375
      Left            =   8280
      TabIndex        =   40
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "초기 목숨 : "
      Height          =   495
      Left            =   8280
      TabIndex        =   38
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label labe 
      Caption         =   "Label1"
      Height          =   255
      Left            =   8880
      TabIndex        =   0
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label Label12 
      Caption         =   "정답 : "
      Height          =   255
      Left            =   8280
      TabIndex        =   37
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "범위 : "
      Height          =   255
      Left            =   8280
      TabIndex        =   36
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   735
      Left            =   11400
      TabIndex        =   35
      Top             =   8040
      Width           =   4455
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   34
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "맞춘 횟수 : "
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   9240
      TabIndex        =   22
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "현재 시간 : "
      Height          =   255
      Left            =   8280
      TabIndex        =   30
      Top             =   4800
      Width           =   4575
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   2760
      TabIndex        =   29
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      TabIndex        =   23
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label7 
      Caption         =   "10"
      Height          =   735
      Left            =   8880
      TabIndex        =   21
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   1095
      Left            =   11520
      TabIndex        =   31
      Top             =   9120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub CheckLifeChar()
    Dim Life As Integer
        Life = Life - 1
        If Err.Description = "형식이 일치하지 않습니다." Then
            MsgBox "생명 갯수를 숫자로 입력하십시오.", vbCritical
            Unload Me
            Form1.Show
        End If
End Sub
Public Sub answerchk()
    If D.Text = "terribleterribledamage" Then
            MsgBox "목숨 무한이되었습니다."
            Text1.Text = "1E+25"
    ElseIf D.Text = "add" Then
            Text1.Text = Text1.Text + 1
    ElseIf D.Text = "Exit" Then
            End
    ElseIf D.Text = "DEVMODE!@#$%^&*" Then
            MsgBox "개발자 모드로 실행합니다.", vbExclamation
            frmLogin.Show
    ElseIf D.Text = "Pass" Then
            Pass
    ElseIf D.Text = "View" Then
            View
    ElseIf D.Text = "End" Then
            MsgBox "종료합니다."
            End
    ElseIf Val(D.Text) > Val(labe.Caption) Then
            MsgBox " 다운 ", vbCritical
            Text1.Text = Text1.Text - 1
    ElseIf Val(D.Text) < Val(labe.Caption) Then
            MsgBox " 업 ", vbCritical
    
   
            Text1.Text = Text1.Text - 1
    Else
        If 100 > Val(ProgressBar1.Value) Then
            MsgBox " 으아 맞추다니.. 쳇 다음문제! ", vbInformation
            Text1.Text = Text1.Text + 1
            Randomize
            labe.Caption = Int(Rnd * Label7.Caption) + 1
            D.Text = ""
            Form2.Label2.Caption = Me.labe.Caption
            Label9.Caption = Label9.Caption + 1
            ProgressBar1.Value = Label9.Caption / Label15.Caption * 100
        End If
    End If
End Sub


Private Sub Command1_Click()
    If Label19.Caption = "1" Then
        
    ElseIf Label19.Caption = "0" Then
        answerchk
    Else
        MsgBox "첨자 오류. 종료합니다.", vbCritical
        End
    End If
End Sub
Public Sub Manswerchk()
    If Val(D.Text) > Val(labe.Caption) Then
            MsgBox " 다운 ", vbCritical
            Text1.Text = Text1.Text - 1
    ElseIf Val(D.Text) < Val(labe.Caption) Then
            MsgBox " 업 ", vbCritical
            Text1.Text = Text1.Text - 1
    Else
        If 100 > Val(ProgressBar1.Value) Then
            MsgBox " 으아 맞추다니.. 쳇 다음문제! ", vbInformation
            Text1.Text = Text1.Text + 1
            Randomize
            labe.Caption = Int(Rnd * Label7.Caption) + 1
            D.Text = ""
            Form2.Label2.Caption = Me.labe.Caption
            Label9.Caption = Label9.Caption + 1
            ProgressBar1.Value = Label9.Caption / Label15.Caption * 100
    End If
    End If
    
    
End Sub
Private Sub Command10_Click()
    Command10.Enabled = False
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
End Sub



Private Sub Command2_Click()
    MsgBox "포기했습니다! 정답 : " & labe.Caption
    Randomize
    labe.Caption = Int(Rnd * Label7.Caption) + 1
    Text1.Text = Text1.Text
    D.Text = ""
    Text1.Text = Me.Label1.Caption
    Form2.Label2.Caption = Me.labe.Caption
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    MsgBox "ㅉㅉ 못맞추다니"""
    Command5.Enabled = True
    Command6.Enabled = True
    Command18.Enabled = True
    Command17.Enabled = True
    Command16.Enabled = True
    Command15.Enabled = True
    Command14.Enabled = True
    Command13.Enabled = True
    Command12.Enabled = True
    Command11.Enabled = True
    Command10.Enabled = True
    Command9.Enabled = True
    Command8.Enabled = True
    Command7.Enabled = True
    Command19.Enabled = True
    Command20.Enabled = True
    Randomize
    labe.Caption = Int(Rnd * Label7.Caption) + 1
End Sub

Private Sub Command22_Click()
    MsgBox "너무 어려운 가벼? ㅋㅋ", vbInformation
    Unload Me
    Form1.Show
    Me.Height = 2070
End Sub

Public Sub View()
    MsgBox "정답은 : " & labe.Caption
End Sub

Public Sub Pass()
    MsgBox "다음 문제로 넘어갔습니다."
    Randomize
    labe.Caption = Int(Rnd * Label7.Caption) + 1
    Form2.Label2.Caption = Me.labe.Caption
End Sub

Private Sub Command23_Click()
    Pass
End Sub

Private Sub Command24_Click()
    Text1.Text = Text1.Text + 1
End Sub

Private Sub Command25_Click()
    End
End Sub

Private Sub Command3_Click()
    frmAbout.Show
End Sub

Private Sub Command4_Click()
    If Check1.Value = 1 Then
        View
        D.Text = labe.Caption
    Else
        View
    End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
End Sub

Private Sub D_Change()
    If Text1.Text = "0" Then
        MsgBox "목숨이없습니다. 프로그램을 종료합니다.", vbCritical
        End
    End If
End Sub

Private Sub D_Click()
    D.Text = ""
End Sub

Private Sub Form_Load()
    If Err.Description = "형식이 일치하지 않습니다." Then
        MsgBox "오류 종료합니다."
        End
    Else
        CheckLifeChar
        Label5.Caption = "위치 : " & App.Path & "\" & App.EXEName & ".exe"
        Text1.Locked = True
        Text1.Text = Form1.Label8.Caption
        Randomize
        labe.Caption = Int(Rnd * Label7.Caption) + 1
        Form2.Label2.Caption = Me.labe.Caption
    End If
End Sub

Public Sub MF()

End Sub
