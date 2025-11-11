Attribute VB_Name = "hlpFilePath"
Option Explicit
' ===== OneDrive/SharePoint URL → 本地路径 =====
' 用法示例：
'   Dim src As String, local As String
'   src = ThisWorkbook.FullName  ' 或 ActivePresentation.FullName / ActiveDocument.FullName
'   local = GetLocalPathFromOfficePath(src)
'   If Len(local) > 0 Then
'       ' 现在 local 就是可用的本地磁盘路径
'   Else
'       MsgBox "无法将云端 URL 映射到本地同步路径。"
'   End If

' -------- 目录工具 --------
' ===== 目录工具 =====

' 判断是否是 URL/HTTP(S) 路径
Private Function IsHttpPath(ByVal p As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(p))
    IsHttpPath = (Left$(s, 7) = "http://" Or Left$(s, 8) = "https://")
End Function

' 递归创建多级目录（稳健版）
Private Sub EnsureFolder(ByVal folderPath As String)
    Dim norm As String, parts() As String, i As Long, cur As String
    If Len(folderPath) = 0 Then Exit Sub
    
    ' 统一分隔符并去掉末尾反斜杠
    norm = Replace(folderPath, "/", "\")
    If Right$(norm, 1) = "\" Then norm = Left$(norm, Len(norm) - 1)
    If Len(norm) = 0 Then Exit Sub
    
    parts = Split(norm, "\")
    If UBound(parts) < 0 Then Exit Sub
    
    ' 构造盘符起点（如 "C:\"）
    If InStr(parts(0), ":") > 0 Then
        cur = parts(0) & "\"
        i = 1
    Else
        ' 网络盘或 UNC 起始（如 "\\server"）
        cur = parts(0) & "\"
        i = 1
    End If
    
    On Error Resume Next
    For i = i To UBound(parts)
        cur = cur & parts(i)
        If Dir(cur, vbDirectory) = "" Then MkDir cur
        cur = cur & "\"
    Next i
    On Error GoTo 0
End Sub

'**************测试程序*****************************************
Sub testChooseLocalOutFolder()
    Dim p As String
    
    p = ChooseLocalOutFolder("https://d.docs.live.net/6ad87b27b1c908f0/Temp/PowerPoint宏/公司新型电力系统综合性示范项目第二轮评审汇报-孙1 - 副本.pptm", "请选择本地保存文件夹")
    Debug.Print p
    
    p = ChooseLocalOutFolder("//.docs.live.net/6ad87b27b1c908f0/Temp/PowerPoint宏/公司新型电力系统综合性示范项目第二轮评审汇报-孙1 - 副本.pptm", "请选择本地保存文件夹")
    Debug.Print p

End Sub


' 结合 OneDrive/SharePoint：把 Office 路径转成本地同步路径；失败则弹框选择
' 返回：本地输出目录；若用户取消，返回空串
Public Function ChooseLocalOutFolder(ByVal officePath As String, _
                                      Optional ByVal dialogTitle As String = "请选择本地保存文件夹") As String
    Dim outFolder As String
    Dim fldr As FileDialog
    
    ' 1) 获取本地路径
    On Error Resume Next
    outFolder = GetLocalPathFromOfficePath(officePath)
    On Error GoTo 0
    
    ' 2) 如果为空（不成功），就让用户选本地文件夹
    If Len(outFolder) = 0 Or IsHttpPath(outFolder) Then
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
            .title = dialogTitle
            .AllowMultiSelect = False
            If .Show <> -1 Then
                ChooseLocalOutFolder = ""   ' 用户取消
                Exit Function
            End If
            outFolder = .SelectedItems(1)
        End With
    End If
    
    ' 3) 确保目录存在（可递归创建）
    Call EnsureFolder(outFolder)
    ChooseLocalOutFolder = outFolder
End Function


Public Function GetLocalPathFromOfficePath(ByVal officePath As String) As String
    On Error Resume Next
    Dim p As String: p = officePath
    
    ' 已经是本地盘符,无需转换
    If InStr(1, p, ":\", vbTextCompare) > 0 Or Left$(p, 2) = "\\" Then
        GetLocalPathFromOfficePath = p
        Exit Function
    End If
    
    ' 不是 URL，直接返回
    If Not IsHttpPath(officePath) Then
        GetLocalPathFromOfficePath = ""
        Debug.Print "非有效路径：" & officePath
        Exit Function
    End If
    
    '尝试将http开头的OneDrive路径转换成本地路径
    GetLocalPathFromOfficePath = OneDriveUrlToLocalPath(officePath)
    
End Function

' 提取 SharePoint/OneDrive for Business 中“Documents/”或“Shared Documents/”之后的相对路径
Private Function ExtractSharePointDocRelative(ByVal url As String) As String
    Dim u As String: u = LCase$(url)
    Dim k1 As String: k1 = "/documents/"
    Dim k2 As String: k2 = "/shared%20documents/"
    Dim pos As Long
    pos = InStr(1, u, k1, vbTextCompare)
    If pos > 0 Then
        ExtractSharePointDocRelative = Mid$(url, pos + Len(k1))
        Exit Function
    End If
    pos = InStr(1, u, k2, vbTextCompare)
    If pos > 0 Then
        ExtractSharePointDocRelative = Mid$(url, pos + Len(k2))
    End If
End Function

' URL 解码（最常用的 %xx 十六进制 和 + 为空格）
Private Function UrlDecode(ByVal s As String) As String
    Dim i As Long, out As String
    i = 1
    Do While i <= Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch = "%" And i + 2 <= Len(s) Then
            Dim hx As String
            hx = Mid$(s, i + 1, 2)
            If hx Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                out = out & Chr$(CLng("&H" & hx))
                i = i + 3
            Else
                out = out & ch
                i = i + 1
            End If
        ElseIf ch = "+" Then
            out = out & " "
            i = i + 1
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    UrlDecode = out
End Function

' 拼接路径（避免重复反斜杠）
Private Function AppendPath(ByVal base As String, ByVal rel As String) As String
    If Right$(base, 1) = "\" Then
        AppendPath = base & rel
    Else
        AppendPath = base & "\" & rel
    End If
End Function

Private Function FileOrFolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileOrFolderExists = (Len(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function GetParentFolder(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then GetParentFolder = Left$(p, i - 1)
End Function

' 经验性猜测 OneDrive 商业版根目录（当环境变量缺失时）
Private Function GuessOneDriveBusinessRoot() As String
    Dim base As String
    base = Environ$("UserProfile")
    If Len(base) = 0 Then Exit Function
    Dim f As String
    f = Dir$(base & "\OneDrive - *", vbDirectory)
    If Len(f) > 0 Then
        GuessOneDriveBusinessRoot = base & "\" & f
    End If
End Function

' 从 s 中提取 startKey 之后到下一个 "/" 之间的内容
Private Function ExtractBetween(ByVal s As String, ByVal startKey As String, ByVal delim As String) As String
    Dim p As Long, q As Long
    p = InStr(1, s, startKey, vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(startKey)
    q = InStr(p, s, delim)
    If q = 0 Then Exit Function
    ExtractBetween = Mid$(s, p, q - p)
End Function

'****************测试程序：OneDrive路径转换***********************
Sub testOneDriveUrlToLocalPath()
    Dim p As String
    p = OneDriveUrlToLocalPath("https://gyqg-my.sharepoint.cn/personal/zysyw_1dts_cn/Documents/SRNT 2024 Poster_JFT2_YS_Mar17_2024.pptx")
    Debug.Print p
    p = OneDriveUrlToLocalPath("https://d.docs.live.net/6ad87b27b1c908f0/Temp/PowerPoint宏/公司新型电力系统综合性示范项目第二轮评审汇报-孙1 - 副本.pptm")
    Debug.Print p
End Sub

' 将 OneDrive/SharePoint 的云端 URL 转成本地同步路径（若可用）
Public Function OneDriveUrlToLocalPath(ByVal url As String) As String
    Dim low As String: low = LCase$(url)
    Dim rest As String, localRoot As String

    ' 个人版 OneDrive：https://d.docs.live.net/<cid>/...
    If InStr(1, low, "d.docs.live.net/", vbTextCompare) > 0 Then
        localRoot = Environ$("OneDriveConsumer")                       ' 个人版本地根
        If Len(localRoot) > 0 Then
            ' 取出 <cid> 之后的路径部分
            Dim parts() As String, i As Long
            parts = Split(url, "/")                            ' ["https:","","d.docs.live.net","<cid>", ...]
            If UBound(parts) >= 4 Then
                For i = 4 To UBound(parts)
                    If Len(rest) > 0 Then rest = rest & "\"
                    rest = rest & parts(i)
                Next
                OneDriveUrlToLocalPath = localRoot & "\" & rest
                Exit Function
            End If
        End If
    End If

    ' 企业/教育版 OneDrive/SharePoint：...my.sharepoint.com/...
    If InStr(1, low, "my.sharepoint.com", vbTextCompare) > 0 Or InStr(1, low, ".sharepoint.", vbTextCompare) > 0 Then
        ' 常见环境变量：OneDriveCommercial（可能存在）
        localRoot = Environ$("OneDriveCommercial")
        If Len(localRoot) > 0 Then
            ' 经验规则：取站点后路径中 "Documents"（或中文“文档”）起的部分
            Dim p As Long, u As String: u = url
            p = InStr(1, u, "/Documents/", vbTextCompare)
            If p = 0 Then p = InStr(1, u, "/文档/", vbTextCompare)
            If p > 0 Then
                rest = Mid$(u, p + 1)                          ' 去掉前面的斜杠
                OneDriveUrlToLocalPath = localRoot & "\" & Replace(rest, "/", "\")
                Exit Function
            End If
        End If
    End If

    ' 转换不成功：返回空字符串
    OneDriveUrlToLocalPath = ""
End Function

'*********************测试程序：打印当前文件的路径********************************************
Sub PrintCurrentPptPath()
    Dim pres As Presentation
    If Presentations.Count = 0 Then
        MsgBox "当前没有打开的演示文稿。", vbExclamation
        Exit Sub
    End If

    Set pres = ActivePresentation

    If Len(pres.path) = 0 Then
        ' 未保存时 Path 为空
        MsgBox "当前文件尚未保存。" & vbCrLf & "文件名：" & pres.Name, vbInformation
    Else
        Debug.Print "目录："; pres.path
        Debug.Print "完整路径："; pres.FullName
        'MsgBox "目录：" & pres.path & vbCrLf & "完整路径：" & pres.FullName, vbInformation
    End If
End Sub
