Attribute VB_Name = "modRestoreNotesFromBackup"
Option Explicit

' 入口：从同名 _Notes.txt 文件恢复备注；如未找到会弹窗让你手动选
Public Sub RestoreNotesFromBackup()
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim notesPath As String
    
    If Not ConfirmAction("恢复所有备注", "将尝试从备份恢复全部备注。") Then Exit Sub

    notesPath = GetDefaultNotesPath(pres)
    If Dir(notesPath) = "" Then
        notesPath = PickNotesFile()
        If Len(notesPath) = 0 Then
            MsgBox "未选择备注备份文件。", vbInformation
            Exit Sub
        End If
    End If
    
    Dim content As String
    content = ReadTextAnsi(notesPath)
    If Len(content) = 0 Then
        MsgBox "备注备份文件为空或读取失败：" & vbCrLf & notesPath, vbExclamation
        Exit Sub
    End If
    
    Dim applied As Long, skipped As Long
    ApplyNotesFromText pres, content, applied, skipped
    
    Dim msg As String
    msg = "备注恢复完成。" & vbCrLf & _
          "成功写入页数：" & applied & vbCrLf & _
          "跳过（页号不存在或格式不合规）：" & skipped & vbCrLf & _
          "来源文件：" & notesPath
    MsgBox msg, vbInformation
End Sub

' 解析文本并写回每一页备注
Private Sub ApplyNotesFromText(ByVal pres As Presentation, ByVal content As String, _
                               ByRef applied As Long, ByRef skipped As Long)
    Dim LF As String: LF = vbLf
    ' 统一换行符为 LF
    content = Replace(content, vbCrLf, LF)
    content = Replace(content, vbCr, LF)
    
    Dim lines() As String
    lines = Split(content, LF)
    
    Dim i As Long
    Dim curSlide As Long: curSlide = 0
    Dim buf As Collection: Set buf = New Collection
    
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(lines(i))
        
        Dim sldNum As Long
        If IsHeaderLine(line, sldNum) Then
            ' 写入上一页
            If curSlide > 0 Then
                If sldNumExists(pres, curSlide) Then
                    WriteNotes pres.Slides(curSlide), JoinCollection(buf, vbCrLf)
                    applied = applied + 1
                Else
                    skipped = skipped + 1
                End If
            End If
            ' 开启新页
            curSlide = sldNum
            Set buf = New Collection
            Dim firstLine As String
            firstLine = GetTextAfterColon(lines(i))   ' 仅取冒号后的文本，不含“Pxx：”
            If Len(firstLine) > 0 Then buf.Add firstLine
        Else
            ' 累积内容（保持空行）
            'buf.Add lines(i)
        End If
    Next i
    
    ' 收尾：写入最后一页
    If curSlide > 0 Then
        If sldNumExists(pres, curSlide) Then
            WriteNotes pres.Slides(curSlide), JoinCollection(buf, vbCrLf)
            applied = applied + 1
        Else
            skipped = skipped + 1
        End If
    End If
End Sub

' 判断是否为页头行：形如 P12: 或 P12：
Private Function IsHeaderLine(ByVal s As String, ByRef slideNum As Long) As Boolean
    IsHeaderLine = False
    slideNum = 0
    If Len(s) < 3 Then Exit Function
    If UCase$(Left$(s, 1)) <> "P" Then Exit Function
    
    ' 找冒号（英文 or 中文全角）
    Dim pos As Long
    pos = InStr(2, s, ":")
    If pos = 0 Then pos = InStr(2, s, "：") ' 全角冒号
    If pos <= 2 Then Exit Function
    
    Dim numPart As String
    numPart = Mid$(s, 2, pos - 2)
    numPart = Trim$(numPart)
    If Not IsNumeric(numPart) Then Exit Function
    
    slideNum = CLng(numPart)
    If slideNum <= 0 Then Exit Function
    
    ' 允许行尾还有任何文字（例如 “P3：” 后紧跟备注第一行也行）
    IsHeaderLine = True
End Function

' 将文本写入指定幻灯片的备注区（覆盖原备注）
Private Sub WriteNotes(ByVal sld As Slide, ByVal noteText As String)
    Dim tr As TextRange
    Set tr = GetNotesTextRange(sld)
    tr.text = noteText
End Sub

' 获取 NotesPage 正文占位符的 TextRange（更健壮的查找方式）
Private Function GetNotesTextRange(ByVal sld As Slide) As TextRange
    Dim shp As Shape
    ' 优先通过占位符类型匹配正文
    On Error Resume Next
    Dim i As Long
    For i = 1 To sld.NotesPage.Shapes.Placeholders.Count
        Set shp = sld.NotesPage.Shapes.Placeholders(i)
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then
                Set GetNotesTextRange = shp.TextFrame.TextRange
                Exit Function
            End If
        End If
    Next i
    ' 兜底：取可编辑的文本框
    For Each shp In sld.NotesPage.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Or shp.TextFrame.TextRange.text <> "" Then
                Set GetNotesTextRange = shp.TextFrame.TextRange
                Exit Function
            End If
        End If
    Next shp
    ' 再兜底：新建一个文本框
    Set shp = sld.NotesPage.Shapes.AddTextbox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=36, Top:=36, Width:=500, Height:=400)
    Set GetNotesTextRange = shp.TextFrame.TextRange
End Function

' 拼接 Collection 为字符串
Private Function JoinCollection(col As Collection, ByVal sep As String) As String
    Dim i As Long
    Dim arr() As String
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i
    JoinCollection = Join(arr, sep)
End Function

' 是否存在该页号
Private Function sldNumExists(ByVal pres As Presentation, ByVal n As Long) As Boolean
    On Error GoTo NOPE
    Dim tmp As Slide
    Set tmp = pres.Slides(n)
    sldNumExists = True
    Exit Function
NOPE:
    sldNumExists = False
End Function

' 默认的 _Notes.txt 路径
Private Function GetDefaultNotesPath(ByVal pres As Presentation) As String
    Dim base As String
    Dim filePath As String
    base = pres.Name
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    filePath = GetLocalPathFromOfficePath(pres.path)
    GetDefaultNotesPath = filePath & IIf(filePath = "", "", "\") & _
                          base & "_Notes.txt"
End Function

' 选择文件对话框
Private Function PickNotesFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "选择备注备份文件（_Notes.txt）"
        .Filters.Clear
        .Filters.Add "文本文件 (*.txt)", "*.txt"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickNotesFile = .SelectedItems(1)
        Else
            PickNotesFile = ""
        End If
    End With
End Function

' 读取文本（优先 UTF-8；失败则尝试 ANSI/系统默认）
Private Function ReadTextFileUTF8(ByVal filePath As String) As String
    On Error GoTo Fallback
    Dim stm As Object ' ADODB.Stream
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Mode = 3 ' read/write
    stm.charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileUTF8 = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing
    Exit Function
Fallback:
    On Error GoTo 0
    ' 退化为普通打开（可能会丢失非 ASCII 字符）
    Dim ff As Integer: ff = FreeFile
    Dim s As String
    Open filePath For Binary As #ff
        s = Space$(LOF(ff))
        If LOF(ff) > 0 Then Get #ff, , s
    Close #ff
    ReadTextFileUTF8 = s
End Function

' 取一行里“P…:（或：）”后的文本；若无冒号则返回空串
Private Function GetTextAfterColon(ByVal s As String) As String
    Dim pos As Long
    pos = InStr(2, s, ":")
    If pos = 0 Then pos = InStr(2, s, "：") ' 支持全角冒号
    If pos > 0 Then
        ' 去掉前导空格，保留原始其余格式
        GetTextAfterColon = LTrim$(Mid$(s, pos + 1))
    Else
        GetTextAfterColon = ""
    End If
End Function

' 用系统默认代码页读取（与 Open...Print# 写出的 ANSI 完全匹配）
Private Function ReadTextAnsi(ByVal filePath As String) As String
    Dim ff As Integer, line As String, buf As String
    ff = FreeFile
    On Error GoTo EH
    Open filePath For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, line          ' 按行读取，VBA 自动用系统代码页解码成 Unicode
        buf = buf & line & vbCrLf     ' 还原换行
    Loop
    Close #ff
    ReadTextAnsi = buf
    Exit Function
EH:
    If ff <> 0 Then Close #ff
    ReadTextAnsi = ""
End Function
