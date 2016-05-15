VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmBGM 
   BorderStyle     =   1  '단일 고정
   Caption         =   "BGM Technology"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7890
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows 기본값
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "열기"
      Height          =   975
      Left            =   7080
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6855
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   12091
      _cy             =   1191
   End
End
Attribute VB_Name = "frmBGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    With CommonDialog1
        .Filter = "All File|*.*"
        .ShowOpen
    End With
        Text1.Text = CommonDialog1.FileName
        WindowsMediaPlayer1.URL = Text1.Text
End Sub
