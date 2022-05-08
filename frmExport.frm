VERSION 5.00
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Begin VB.Form frmExport 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdDesktop 
      Caption         =   "桌面"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   2760
      Style           =   1  'Checkbox
      TabIndex        =   4
      Tag             =   "HW"
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导入"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   735
   End
   Begin VB6ResizerLib2.VB6Resizer VB6Resizer1 
      Left            =   5160
      Top             =   2400
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.DirListBox Dirs 
      Height          =   2940
      Left            =   120
      TabIndex        =   1
      Tag             =   "H"
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox Drvs 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDesktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case cmdDesktop.Caption
        Case ConstStr("desktop")
            Dirs.Path = Environ("UserProfile") & "\Desktop"
        Case ConstStr("appdir")
            Dirs.Path = App.Path
        Case ConstStr("qq")
            Dirs.Path = Environ("UserProfile") & "\Documents\Tencent Files"
        End Select
    Case 2
        Select Case cmdDesktop.Caption
        Case ConstStr("desktop")
            cmdDesktop.Caption = ConstStr("appdir")
        Case ConstStr("appdir")
            cmdDesktop.Caption = ConstStr("qq")
        Case ConstStr("qq")
            cmdDesktop.Caption = ConstStr("desktop")
        End Select
    End Select
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub
Private Sub cmdExport_Click()
    Dim ItemNum As Long, RetText As String, fso As New FileSystemObject
    For ItemNum = 0 To List1.ListCount - 1
        If List1.Selected(ItemNum) Then
            If List1.List(ItemNum) <> "" Then
                RetText = RetText & List1.List(ItemNum) & " "
                fso.CopyFile LevelPath & "\" & List1.List(ItemNum), Dirs.Path & "\" & List1.List(ItemNum), True
            End If
        End If
    Next
    If RetText <> "" Then MsgBox RetText & ConstStr("export_completed")
    Me.Hide
    Unload Me
End Sub


Private Sub Drvs_Change()
    Dirs.Path = UCase(Split(Drvs.Drive, ":")(0)) & ":\"
End Sub


Private Sub Form_Load()
    Me.Caption = ConstStr("title") & " " & App.Major & "." & App.Minor & "." & App.Revision & " - " & ConstStr("t_export")    '窗口标题
    cmdExport.Caption = ConstStr("export")
    cmdCancel.Caption = ConstStr("cancel")
    cmdDesktop.Caption = ConstStr("desktop")
    cmdDesktop.ToolTipText = ConstStr("right_change")
    List1.ToolTipText = ConstStr("drag_tooltip")
    List1.Clear
    Dim LevelFile As Variant
    For Each LevelFile In GetFileList(LevelPath, "*.swe")
        List1.AddItem LevelFile
    Next LevelFile
    Dirs.Path = App.Path
End Sub

