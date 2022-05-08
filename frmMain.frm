VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "SMMWE Level Manager"
   ClientHeight    =   5520
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9105
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   9105
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame frm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "关卡详情"
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   7200
      TabIndex        =   1
      Tag             =   "LH"
      Top             =   0
      Width           =   1815
      Begin VB.Image imgStage3 
         Height          =   255
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   255
      End
      Begin VB.Image imgStage2 
         Height          =   375
         Left            =   960
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   735
      End
      Begin VB.Image imgStage1 
         Height          =   375
         Left            =   120
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label lblLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Click2Show"
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Tag             =   "H"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Image imgThumbnail 
         Height          =   975
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1575
      End
   End
   Begin VBCCR17.ListView lst 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Tag             =   "HW"
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9340
      VisualTheme     =   1
      Icons           =   "imgGameStyle"
      SmallIcons      =   "imgGameStyle"
      ColumnHeaderIcons=   "imgGameStyle"
      GroupIcons      =   "imgGameStyle"
      View            =   3
      LabelEdit       =   1
   End
   Begin VBCCR17.ImageList imgGameStyle 
      Left            =   8040
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      UseBackColor    =   -1  'True
      UseMaskColor    =   0   'False
      InitListImages  =   "frmMain.frx":FDDA
   End
   Begin VB6ResizerLib2.VB6Resizer Resizer 
      Left            =   8640
      Top             =   5040
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.Menu mImport 
      Caption         =   "导入"
   End
   Begin VB.Menu mExport 
      Caption         =   "导出"
   End
   Begin VB.Menu mDelete 
      Caption         =   "删除"
   End
   Begin VB.Menu mMore 
      Caption         =   "更多"
      Begin VB.Menu mOpenLevelFolder 
         Caption         =   "打开关卡文件夹"
      End
      Begin VB.Menu mClearCache 
         Caption         =   "清理缓存"
      End
      Begin VB.Menu mRefresh 
         Caption         =   "刷新"
      End
   End
   Begin VB.Menu mAccount 
      Caption         =   "登录"
      Visible         =   0   'False
   End
   Begin VB.Menu mAbout 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StdPic As New stdPicEx2

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    Dim LevelList() As String, SingleLevel As Variant, i As Long, GameStyle As String, LevelName As String
    '加载配置
    frmDummy.Show
    DoEvents
    Set Conf = JSON.parse(ReadTextFile(App.Path & "\conf.json"))    '加载配置文件
    Set ConstStr = JSON.parse(ReadTextFile(App.Path & "\assets\locale-" & Conf("locale") & ".json"))    '加载语言文件
    frmDummy.Label1.Caption = ConstStr("loading") & ConstStr("loading_configuration")
    DoEvents
    Set ConstStr = ConstStr("locale")
    Me.Caption = ConstStr("title") & " " & App.Major & "." & App.Minor & "." & App.Revision & " - " & ConstStr("t_local_levels")    '窗口标题
    LevelPath = Environ("UserProfile") & "\AppData\Local\SMM_WE\Niveles"    '关卡路径
    If Dir(LevelPath, vbDirectory) = "" Then MsgBox ConstStr("err_need_a_level"), vbCritical: End    '关卡文件夹是否存在
    If Dir(App.Path & "\cache", vbDirectory) = "" Then MkDir App.Path & "\cache"    '创建缓存文件夹
    LoadLocalLevels
    lblLevel.Caption = ConstStr("click_to_show")
    frm.Caption = ConstStr("level_details")
    mImport.Caption = ConstStr("import")
    mExport.Caption = ConstStr("export")
    mAbout.Caption = ConstStr("about")
    mAccount.Caption = ConstStr("login")
    mDelete.Caption = ConstStr("delete")
    mMore.Caption = ConstStr("more")
    mRefresh.Caption = ConstStr("refresh")
    mOpenLevelFolder.Caption = ConstStr("open_level_folder")
    mClearCache.Caption = ConstStr("clear_cache")
    frmDummy.Hide
    Unload frmDummy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lst_Click()
    Dim SingleMeta As Object, TagClass As String
    Set SingleMeta = LevelMeta(lst.SelectedItem.Tag)("MAIN")("AJUSTES")(1)
    lblLevel.Caption = ConstStr("author") & ": " & SingleMeta("user")
    If SingleMeta("etiqueta2") = -1 Then
        lblLevel.Caption = lblLevel.Caption & vbCrLf & ConstStr("tag") & ": " & ConstStr("tags")(SingleMeta("etiqueta1") + 1)
    Else
        lblLevel.Caption = lblLevel.Caption & vbCrLf & ConstStr("tag") & ": " & vbCrLf & ConstStr("tags")(SingleMeta("etiqueta1") + 1) & ", " & ConstStr("tags")(SingleMeta("etiqueta2") + 1)
    End If
    lblLevel.Caption = lblLevel.Caption & vbCrLf & ConstStr("modify_date") & ": " & SingleMeta("date") & " " & SingleMeta("time")
    lblLevel.Caption = lblLevel.Caption & vbCrLf & ConstStr("timer") & ": " & SingleMeta("cronometro")
    lblLevel.Caption = lblLevel.Caption & vbCrLf & ConstStr("autoscroll") & ": " & SingleMeta("autoavance") & "x"
    lblLevel.Caption = lblLevel.Caption & vbCrLf & ConstStr("size") & ": " & GetFileSize(LevelPath & "\" & Base64Decode(lst.SelectedItem.Tag))
    If Conf("locale") = "es_ES" Then
        TagClass = "es"
    Else
        TagClass = "en"
    End If
    If CheckFileExists(App.Path & "\assets\tags_" & TagClass & "\tags-" & CStr(SingleMeta("etiqueta1")) & "-" & TagClass & ".png") Then
        imgThumbnail.Picture = StdPic.LoadPictureEx(App.Path & "\assets\tags_" & TagClass & "\tags-" & CStr(SingleMeta("etiqueta1")) & "-" & TagClass & ".png")
    Else
        Select Case SingleMeta("etiqueta1")
        Case 9
            imgThumbnail.Picture = StdPic.LoadPictureEx(App.Path & "\assets\tags_" & TagClass & "\tags-8-" & TagClass & ".png")
        Case 10
            imgThumbnail.Picture = StdPic.LoadPictureEx(App.Path & "\assets\tags_" & TagClass & "\tags-1-" & TagClass & ".png")
        Case 11
            imgThumbnail.Picture = StdPic.LoadPictureEx(App.Path & "\assets\tags_" & TagClass & "\tags-7-" & TagClass & ".png")
        Case 12
            imgThumbnail.Picture = StdPic.LoadPictureEx(App.Path & "\assets\tags_" & TagClass & "\tags-6-" & TagClass & ".png")
        Case 13
            imgThumbnail.Picture = StdPic.LoadPictureEx(App.Path & "\assets\tags_" & TagClass & "\tags-3-" & TagClass & ".png")
        Case 14
            imgThumbnail.Picture = StdPic.LoadPictureEx(App.Path & "\assets\tags_" & TagClass & "\tags-1-" & TagClass & ".png")
        End Select
    End If
    imgStage1.Picture = StdPic.LoadPictureEx(App.Path & "\assets\game_styles\game_style-" & ConstStr("game_styles")(CInt(SingleMeta("apariencia")) + 1) & ".png")
    imgStage2.Picture = StdPic.LoadPictureEx(App.Path & "\assets\stages\stage-" & SingleMeta("entorno") & ".png")
    imgStage3.Picture = StdPic.LoadPictureEx(App.Path & "\assets\day_night\day_night-" & SingleMeta("modo_noche") & ".png")
End Sub

Private Sub mAbout_Click()
    frmAbout.Show
End Sub


Private Sub mClearCache_Click()
    MsgBox ConstStr("clear_cache_text")(1) & GetFolderSize(App.Path & "\cache") & ConstStr("clear_cache_text")(2)
    ShellAndWait "cmd /c rd /s /q """ & App.Path & "\cache"""
    Do Until Dir(App.Path & "\cache", vbDirectory) = ""
        Sleep 200
    Loop
    MkDir App.Path & "\cache"
End Sub

Private Sub mDelete_Click()
'删除关卡
    Dim LevelName As String
    LevelName = Base64Decode(lst.SelectedItem.Tag)
    If MsgBox(ConstStr("delete_text")(1) & LevelName & ConstStr("delete_text")(2), vbYesNo) = vbYes Then
        Kill LevelPath & "\" & LevelName
        Kill App.Path & "\cache\" & LevelName & ".cache"
        MsgBox LevelName & " " & ConstStr("delete_text")(3)
        LoadLocalLevels
    End If
End Sub

Private Sub mImport_Click()
    frmImport.Show
End Sub
Private Sub mExport_Click()
    frmExport.Show
End Sub

Private Sub mOpenLevelFolder_Click()
    Shell "cmd /c start """" """ & LevelPath & """"
End Sub

Private Sub mRefresh_Click()
    LoadLocalLevels
End Sub

Private Sub Resizer_AfterResize()
    On Error Resume Next
    lst.ColumnHeaders(3).Width = frmMain.Width - 4750
End Sub

Public Sub LoadLocalLevels()
'加载列标头
    LevelMeta.RemoveAll
    lst.ListItems.Clear
    lst.ColumnHeaders.Clear
    lst.ColumnHeaders.Add 1, "Icon", ConstStr("game_style"), 1100
    lst.ColumnHeaders.Add 2, "Maker", ConstStr("author"), 1300
    lst.ColumnHeaders.Add 3, "Level", ConstStr("level"), 5000
    lst.FullRowSelect = True
    lst.GridLines = True
    '加载关卡
    LevelList = GetFileList(LevelPath, "*.swe")
    For Each SingleLevel In LevelList
        frmDummy.Label1.Caption = ConstStr("loading") & ConstStr("loading_level") & vbCrLf & SingleLevel
        DoEvents
        LevelMeta.Add Base64Encode(CStr(SingleLevel)), ReadLevel(CStr(SingleLevel))    '用解析函数加载关卡dictionary
    Next SingleLevel
    For Each SingleLevel In LevelMeta.keys
        i = i + 1
        GameStyle = ConstStr("game_styles")(CInt(LevelMeta(SingleLevel)("MAIN")("AJUSTES")(1)("apariencia")) + 1)
        lst.ListItems.Add i, CStr(i), GameStyle, , GameStyle
        lst.ListItems(i).SubItems(1) = LevelMeta(SingleLevel)("MAIN")("AJUSTES")(1)("user")
        LevelName = Base64Decode(CStr(SingleLevel))
        lst.ListItems(i).SubItems(2) = Replace(LevelName, ".swe", "")
        frmDummy.Label1.Caption = ConstStr("loading") & ConstStr("loading_level") & vbCrLf & LevelName
        DoEvents
        lst.ListItems(i).Tag = CStr(SingleLevel)
    Next SingleLevel
    On Error Resume Next
    lst.ColumnHeaders(3).Width = frmMain.Width - 4750
End Sub
