Attribute VB_Name = "modExportNotes"
Option Explicit

Sub ExportNotes()
    Dim sld As Slide
    Dim shp As Shape
    Dim noteText As String
    Dim outFile As String
    Dim fileNum As Integer
    Dim baseName As String
    Dim outFolder As String
    Dim fldr As FileDialog

    '—— 1. 确定保存文件名基础部分 ——
    baseName = Left(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1)
    outFolder = GetLocalPathFromOfficePath(ActivePresentation.path)
    
    '—— 2. 判断 Path 是否已经为本地路径 ——
    If outFolder = "" _
       Or InStr(1, outFolder, "http", vbTextCompare) > 0 Then
        ' 未保存或打开自 URL：弹对话框让用户选本地文件夹
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
            .title = "请选择保存备注 TXT 的本地文件夹"
            .AllowMultiSelect = False
            If .Show <> -1 Then
                MsgBox "未选择文件夹，导出已取消。", vbExclamation
                Exit Sub
            End If
            outFolder = .SelectedItems(1)
        End With
    End If

    outFile = outFolder & "\" & baseName & "_Notes.txt"

    '—— 3. 打开文件写入 ——
    fileNum = FreeFile
    Open outFile For Output As #fileNum

    For Each sld In ActivePresentation.Slides
        noteText = GetNotesText(sld)

        Print #fileNum, "P" & sld.SlideIndex & ": " & noteText
        Print #fileNum,  ' 空一行
    Next sld

    Close #fileNum
    MsgBox "备注已导出到：" & vbCrLf & outFile, vbInformation, "导出完成"
End Sub


' 读取备注文本（仅取备注正文；排除页眉/页脚/日期/页码）
Public Function GetNotesText(ByVal sld As Slide) As String
    On Error GoTo EH
    Dim s As String, t As String, shp As Shape

    ' 1) 常规：备注正文通常是 Placeholders(2)
    On Error Resume Next
    s = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text
    On Error GoTo EH
    If Len(s) > 0 Then
        GetNotesText = CleanNoteText(s)
        Exit Function
    End If

    ' 2) 遍历形状：仅收集备注正文，占位符以外的纯文本框也收
    For Each shp In sld.NotesPage.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                If shp.Type = msoPlaceholder Then
                    Select Case shp.PlaceholderFormat.Type
                        Case ppPlaceholderBody
                            t = t & shp.TextFrame.TextRange.text & vbCrLf
                        Case ppPlaceholderHeader, ppPlaceholderFooter, ppPlaceholderDate, ppPlaceholderSlideNumber
                            ' 忽略
                        Case Else
                            ' 其它占位符忽略
                    End Select
                Else
                    ' 非占位符的文本框（有人把备注写在自加的文本框里）
                    't = t & shp.TextFrame.TextRange.text & vbCrLf '2025.11.3发现该问题，有的幻灯片修改了备注页，加入了其他的文本框，可以在备注视图中看到，这个没必要考虑
                End If
            End If
        End If
    Next shp

    GetNotesText = CleanNoteText(t)
    Exit Function
EH:
    GetNotesText = ""
End Function

' 去掉尾部多余换行/空白，并清理零宽/BOM
Private Function CleanNoteText(ByVal s As String) As String
    s = Replace$(s, ChrW(&HFEFF), "")               ' BOM
    s = Replace$(s, ChrW(&H200B), "")               ' ZWSP
    s = Replace$(s, ChrW(&H200C), "")
    s = Replace$(s, ChrW(&H200D), "")
    ' 去尾部 CR/LF
    Do While Len(s) > 0 And (Right$(s, 1) = vbCr Or Right$(s, 1) = vbLf)
        s = Left$(s, Len(s) - 1)
    Loop
    CleanNoteText = Trim$(s)
End Function
