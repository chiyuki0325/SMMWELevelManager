Attribute VB_Name = "Module"
Public Conf As Object, ConstStr As Object
Attribute ConstStr.VB_VarUserMemId = 1073741824
Public LevelPath As String
Attribute LevelPath.VB_VarUserMemId = 1073741826
Public LevelMeta As New Dictionary
Attribute LevelMeta.VB_VarUserMemId = 1073741827


Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Function CheckFileExists(FilePath As String) As Boolean
'检查文件是否存在
    On Error GoTo Err
    If Len(FilePath) < 2 Then CheckFileExists = False: Exit Function
    If Dir$(FilePath, vbAllFileAttrib) <> vbNullString Then CheckFileExists = True
    Exit Function
Err:
    CheckFileExists = False
End Function

Public Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 0, True
    End With
End Sub

Public Function ReadTextFile(sFilePath As String) As String
    On Error Resume Next
    Dim handle As Integer
    If LenB(Dir$(sFilePath)) > 0 Then
        handle = FreeFile
        Open sFilePath For Binary As #handle
        ReadTextFile = Space$(LOF(handle))
        Get #handle, , ReadTextFile
        Close #handle
    End If
End Function

Function GetFileList(ByVal Path As String, Optional fExp As String = "*.*") As String()
    Dim fName As String, i As Long, FileName() As String
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    fName = Dir$(Path & fExp)
    i = 0
    Do While fName <> ""
        ReDim Preserve FileName(i) As String
        FileName(i) = fName
        fName = Dir$
        i = i + 1
    Loop
    If i <> 0 Then
        ReDim Preserve FileName(i - 1) As String
    End If
    GetFileList = FileName
End Function

Public Function ReadLevel(ByVal LevelName As String) As Object
    If CheckFileExists(App.Path & "\cache\" & LevelName & ".cache") Then
        Set ReadLevel = JSON.parse(ReadTextFile(App.Path & "\cache\" & LevelName & ".cache"))
    Else
        Dim Content As String
        Content = ReadTextFile(LevelPath & "\" & LevelName)
        Content = Base64Decode(Left(Content, Len(Content) - 40))
        Set ReadLevel = JSON.parse(Content)
        Open App.Path & "\cache\" & LevelName & ".cache" For Output As #2
        Print #2, Content;
        Close #2
    End If
End Function
Function GetFileSize(someFile)
    Dim fs
    Dim File
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set File = fs.GetFile(someFile)
    GetFileSize = FormatFileSize(File.Size)
    Set File = Nothing
    Set fs = Nothing
End Function
Function GetFolderSize(someFile)
    Dim fs
    Dim File
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set File = fs.GetFolder(someFile)
    GetFolderSize = FormatFileSize(File.Size)
    Set File = Nothing
    Set fs = Nothing
End Function
Function FormatFileSize(Size)
    Dim units
    Dim factor
    units = Array("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    factor = Log(Size) \ 7
    FormatFileSize = Round(Size / (1024 ^ factor), 2) & units(factor)
End Function

Public Function ChooseFile(ByVal frmTitle As String, ByVal fileDescription As String, ByVal fileFilter As String, ByVal onForm As Object) As String
'oleexp 选择文件
    On Error Resume Next
    Dim pChoose As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    Dim tFilt() As COMDLG_FILTERSPEC
    ReDim tFilt(0)
    tFilt(0).pszName = fileDescription
    tFilt(0).pszSpec = fileFilter
    With pChoose
        .SetFileTypes UBound(tFilt) + 1, VarPtr(tFilt(0))
        .SetTitle frmTitle
        .SetOptions FOS_FILEMUSTEXIST + FOS_DONTADDTORECENT
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseFile = BStrFromLPWStr(lpPath)
End Function

Public Function BStrFromLPWStr(lpWStr As Long) As String
    SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
End Function

Public Function ChooseDir(ByVal frmTitle As String, onForm As Object) As String
'oleexp 选择目录
    On Error Resume Next
    Dim pChooseDir As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    With pChooseDir
        .SetOptions FOS_PICKFOLDERS
        .SetTitle frmTitle
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseDir = BStrFromLPWStr(lpPath)
End Function



