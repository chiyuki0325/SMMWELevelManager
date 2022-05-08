VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5985
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "文字文字"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "By 是一刀斩哒"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SMMWE 关卡管理器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StdPic As New stdPicEx2

Private Sub Form_Load()
    Me.Caption = ConstStr("title") & " " & App.Major & "." & App.Minor & "." & App.Revision & " - " & ConstStr("about")    '窗口标题
    Image1.Picture = StdPic.LoadPictureEx(App.Path & "\assets\about.png")
    Label1.Caption = ConstStr("about_text")(1)
    Label2.Caption = ConstStr("about_text")(2) & " " & ConstStr("version") & ": " & App.Major & "." & App.Minor & "." & App.Revision
    Label3.Caption = ConstStr("about_text")(3) & vbCrLf & vbCrLf & ConstStr("about_text")(4)
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

