Attribute VB_Name = "modExportModles"
Option Explicit

'——— 常量（晚绑定，不依赖引用）———
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

' 导出所有VBA模块
Public Sub ExportAllVbaModules_OnClick()
    Dim pres As Presentation
    If ActivePresentation Is Nothing Then
        MsgBox "未检测到活动演示文稿。", vbExclamation
        Exit Sub
    End If
    Set pres = ActivePresentation

    ' —— 可配置项 ——
    Dim outputRoot As String
    outputRoot = pres.path & "\vba_modules"                 ' 导出根目录
    Dim perPresSubfolder As Boolean: perPresSubfolder = False ' 是否按演示文稿建子目录
    Dim includeDocumentModules As Boolean: includeDocumentModules = False ' 是否导出文档模块（ThisPresentation等）
    Dim clearOutputFirst As Boolean: clearOutputFirst = True ' 导出前清空目标目录内旧的 .bas/.cls/.frm/.frx

    ' 开始导出
    ExportAllVbaModules pres, outputRoot, perPresSubfolder, includeDocumentModules, clearOutputFirst
End Sub

' 核心过程：批量导出
Public Sub ExportAllVbaModules( _
    ByVal pres As Presentation, _
    ByVal outputRoot As String, _
    Optional ByVal perPresentationSubfolder As Boolean = True, _
    Optional ByVal includeDocumentModules As Boolean = False, _
    Optional ByVal clearOutputFirst As Boolean = False)

    On Error GoTo EH

    ' 目录准备
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outDir As String
    If Len(outputRoot) = 0 Then
        outDir = pres.path & "\vba_modules"
    Else
        outDir = outputRoot
    End If
    If perPresentationSubfolder Then
        outDir = outDir & "\" & StripExt(fso.GetFileName(pres.FullName))
    End If
    EnsureFolder fso, outDir

    If clearOutputFirst Then
        DeletePatternIfExists fso, outDir, "*.bas"
        DeletePatternIfExists fso, outDir, "*.cls"
        DeletePatternIfExists fso, outDir, "*.frm"
        DeletePatternIfExists fso, outDir, "*.frx"
    End If

    ' VBIDE 工程
    Dim vbProj As Object, vbComp As Object
    Set vbProj = pres.VBProject

    Dim countExport As Long, countSkip As Long
    Dim ext As String, t As Long
    Dim targetPath As String

    For Each vbComp In vbProj.VBComponents
        t = vbComp.Type
        ext = ""

        Select Case t
            Case vbext_ct_StdModule:           ext = ".bas"
            Case vbext_ct_ClassModule:         ext = ".cls"
            Case vbext_ct_MSForm:              ext = ".frm"  ' 将自动带出 .frx
            Case vbext_ct_Document
                If includeDocumentModules Then ext = ".cls"   ' 文档模块按类模块导出
            Case Else
                ' 其他类型忽略
        End Select

        If Len(ext) > 0 Then
            targetPath = outDir & "\" & vbComp.Name & ext

            ' 导出（若文件被只读或占用，导出会抛错）
            On Error Resume Next
            vbComp.Export targetPath
            If Err.Number = 0 Then
                countExport = countExport + 1
            Else
                countSkip = countSkip + 1
                Debug.Print "导出失败 → "; vbComp.Name; " ："; Err.Number; " - "; Err.Description
                Err.Clear
            End If
            On Error GoTo EH
        Else
            countSkip = countSkip + 1
        End If
    Next vbComp

    MsgBox "导出完成：" & vbCrLf & _
           "输出目录: " & outDir & vbCrLf & _
           "成功导出: " & countExport & " 个组件" & vbCrLf & _
           "跳过/失败: " & countSkip & " 个组件", vbInformation
    Exit Sub

EH:
    MsgBox "导出过程中发生错误：" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

'——— 工具函数 ———

Private Sub EnsureFolder(fso As Object, ByVal folderPath As String)
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Sub DeletePatternIfExists(fso As Object, ByVal folderPath As String, ByVal pattern As String)
    On Error Resume Next
    Dim fld As Object, fil As Object
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set fld = fso.GetFolder(folderPath)
    For Each fil In fld.Files
        If LCase(fso.GetFileName(fil.path)) Like LCase(pattern) Then
            fil.Delete True
        End If
    Next fil
    On Error GoTo 0
End Sub

Private Function StripExt(ByVal fileName As String) As String
    Dim p As Long: p = InStrRev(fileName, ".")
    If p > 0 Then
        StripExt = Left$(fileName, p - 1)
    Else
        StripExt = fileName
    End If
End Function


