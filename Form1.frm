VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Up And Down EP02 v3.2.0 with Multiplayer v0.01"
   ClientHeight    =   3750
   ClientLeft      =   4020
   ClientTop       =   1425
   ClientWidth     =   8235
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8235
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   180
      Left            =   5880
      TabIndex        =   29
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   5760
      TabIndex        =   27
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "BGM"
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "치트 난이도(1~1E+25)                                     치트족 다모여라~~!!"
      Enabled         =   0   'False
      Height          =   855
      Left            =   600
      TabIndex        =   25
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command16 
      Caption         =   "멀티(&M)"
      Height          =   2655
      Left            =   0
      TabIndex        =   24
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "스페셜 난이도 "
      Height          =   855
      Left            =   2760
      TabIndex        =   23
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command14 
      Caption         =   "추천1~120"
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "하수(1~10)"
      Height          =   495
      Left            =   1440
      TabIndex        =   21
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "중수(1~25)"
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "고수(1~154)"
      Height          =   495
      Left            =   3120
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "굿 (1~450)"
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "확인"
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   5760
      TabIndex        =   16
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   2535
      Left            =   10560
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   11640
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11040
      Top             =   12360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "공식 캠페인(시나리오)"
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "종료(&E)"
      Height          =   2655
      Left            =   4800
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "신(0~1030)"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "초인(1~784)"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "고급자(1~143)"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "중급자(1~100)"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "초보자(1~70)"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   5760
      TabIndex        =   9
      ToolTipText     =   "1234567890"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "목숨 갯수를 입력(숫자만)"
      Height          =   615
      Left            =   5640
      TabIndex        =   10
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "목표 조절하기"
      Height          =   255
      Left            =   6240
      TabIndex        =   28
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "범위 조절하기"
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "life"
      Height          =   1095
      Left            =   8520
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   735
      Left            =   7680
      TabIndex        =   7
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   975
      Left            =   10560
      TabIndex        =   6
      Top             =   11640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "난이도 "
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    frmMain.Text1.Text = Me.Text1.Text
    Unload Me
    frmMain.Show
    frmMain.Label7.Caption = "70"
    Label8.Caption = Text1.Text
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command10_Click()
    frmMain.Text1.Text = Me.Text1.Text
    Unload Me
    frmMain.Show
    frmMain.Label7.Caption = "25"
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command11_Click()
    MsgBox "시나리오에 오신걸 환영합니다.", vbInformation
    MsgBox "1단계 10번 맞춰라!, 범위는 50!, 목숨은 총 50개!"
    frmMain.Label7.Caption = "50"
    frmMain.Label15.Caption = "10"
    frmMain.Label17.Caption = "1"
    frmMain.Text1.Text = "50"
    frmMain.Command22.Enabled = False
    frmMain.Command2.Enabled = False
    frmMain.Command3.Enabled = False
    Unload Me
    frmMain.Show
End Sub

Private Sub Command12_Click()
    Label8.Caption = Text5.Text
    frmMain.Text1.Text = Me.Text1.Text
    frmMain.Label7.Caption = Me.Label8.Caption
    Unload Me
    frmMain.Show
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command13_Click()
    frmMain.Text1.Text = Me.Text1.Text
    Unload Me
    frmMain.Show
    frmMain.Label7.Caption = "10"
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command14_Click()
    frmMain.Text1.Text = Me.Text1.Text
    Unload Me
    frmMain.Show
    frmMain.Label7.Caption = "120"
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command15_Click()
    MsgBox "스페셜 난이도 활성화!! 즐거운 시간되세욬", vbInformation
    Randomize
    frmMain.Label7.Caption = Int(Rnd * 500) + 1
    If Text1.Text = "" And "0" Then
        MsgBox "목숨 갯수가 입력되지 않았습니다."
    Else
        frmMain.Text1.Text = Me.Text1.Text
        Unload Me
        frmMain.Show
        Label8.Caption = Text1.Text
        Form2.Label7.Caption = frmMain.Label7.Caption
    End If
End Sub

Private Sub Command17_Click()
    'frmMain.Label15.Caption = Text2.Text
  frmMulti.Show
Center.Show

End Sub

Private Sub Command16_Click()
    MsgBox "현재 멀티는 기술상으로 오류가 나와, 개발중단되었습니다.", vbCritical
End Sub

Private Sub Command2_Click()
    If Text1.Text = "목숨 갯수를 입력" Then
        MsgBox "목숨 갯수가 입력되지 않았습니다."
    Else
        frmMain.Text1.Text = Me.Text1.Text
        Unload Me
        frmMain.Show
        frmMain.Label7.Caption = "100"
        Label8.Caption = Text1.Text
        Form2.Label7.Caption = frmMain.Label7.Caption
    End If
End Sub

Private Sub Command20_Click()
    If Text1.Text = "목숨 갯수를 입력" Then
        MsgBox "목숨 갯수가 입력되지 않았습니다."
    Else
        frmMain.Text1.Text = Me.Text1.Text
        Unload Me
        frmMain.Show
        frmMain.Label7.Caption = "1E+25"
        MsgBox "치트 난이도 안내 - 현재 테스트중입니다. 치트 아는분만 하시는걸 추천합니다.", vbExclamation
        Label8.Caption = Text1.Text
        Form2.Label7.Caption = frmMain.Label7.Caption
        frmMain.Height = 4275
    End If
End Sub

Private Sub Command3_Click()
    
    frmMain.Text1.Text = Me.Text1.Text
    Unload Me
    frmMain.Show
    frmMain.Label7.Caption = "143"
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption

End Sub

Private Sub Command4_Click()
    If Text1.Text = "목숨 갯수를 입력" Then
        MsgBox "목숨 갯수가 입력되지 않았습니다."
    Else
        frmMain.Text1.Text = Me.Text1.Text
        Unload Me
        frmMain.Show
        frmMain.Label7.Caption = "784"
        MsgBox "고급자 다함? ㅋㅋ", vbQuestion
        Label8.Caption = Text1.Text
        Form2.Label7.Caption = frmMain.Label7.Caption
    End If
End Sub

Private Sub Command5_Click()
    If Text1.Text = "목숨 갯수를 입력" Then
        MsgBox "목숨 갯수가 입력되지 않았습니다."
    Else
        frmMain.Text1.Text = Me.Text1.Text
        Unload Me
        frmMain.Show
        frmMain.Label7.Caption = "1030"
        MsgBox "용케도 선택 하셨군 ㅋㅋ", vbInformation
        Form2.Label7.Caption = frmMain.Label7.Caption
    End If
End Sub

Private Sub Command6_Click()
    frmMain.Text1.Text = Me.Text1.Text
    Unload Me
    frmMain.Show
    frmMain.Label7.Caption = "450"
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command7_Click()
    frmMain.Text1.Text = Me.Text1.Text
    Unload Me
    frmMain.Show
    frmMain.Label7.Caption = "154"
    Label8.Caption = Text1.Text
    Form2.Label7.Caption = frmMain.Label7.Caption
End Sub

Private Sub Command8_Click()
    MsgBox "종료합니다", vbCritical
    MsgBox "5", vbCritical
    
    MsgBox "4", vbCritical
    
    MsgBox "3", vbCritical
    
    MsgBox "2", vbCritical
    
    MsgBox "1", vbCritical
    
    MsgBox "펑", vbInformation
    End
End Sub

Private Sub Command9_Click()
    frmBGM.Show
    
End Sub

'Private Sub Command9_Click()
'If Val(Text2.Text) > Val("20000") Then
'Text1.Text = Text2.Text / "100"
'MsgBox "당신은 " & Text2.Text & "원 내셨으며, 처리된 목숨갯수는 " & Text1.Text & "개 입니다. 즐겜요!", vbInformation
'Text1.Text = Me.Text1.Text + 1
'MsgBox "20000원이상 내셔서 목숨한개 더 추가되었습니다!"
'Else
'Text1.Text = Text2.Text / "100"
'MsgBox "당신은 " & Text2.Text & "원 내셨으며, 처리된 목숨갯수는 " & Text1.Text & "개 입니다. 즐겜요!", vbInformation
'End If
'End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_Change()
    frmMain.Text1.Text = Me.Text1.Text
    Me.Label8.Caption = Me.Text1.Text
    Me.Label11.Caption = Me.Text1.Text
    frmMain.Label1.Caption = Me.Text1.Text
End Sub

Private Sub Text1_Click()
    Text1.Text = ""
End Sub

Private Sub Text2_Change()
    Text1.Enabled = False
End Sub

