Attribute VB_Name = "modConvertNotesToAudio"
Option Explicit
' 依赖：FA_Config.bas、FA_TTS.bas（调用 FA_TTS_ToFile）

Public Const TAG_KEY As String = "TTS_AUDIO"
Public Const TAG_VAL As String = "FishAudio"

'将所有片子的备注转化为音频，保存和嵌入音频文件
Public Sub NotesToAudio_AllSlides()
    Dim pres As Presentation, baseDir As String, audioDir As String
    Dim i As Long, sld As Slide, notesText As String, mp3Path As String
    Dim shp As Shape, durSec As Double, advanceSec As Double

    Set pres = ActivePresentation
    baseDir = pres.path
    If Len(baseDir) = 0 Then
        MsgBox "请先保存演示文稿再运行。", vbExclamation: Exit Sub
    End If

    audioDir = baseDir & IIf(Right$(baseDir, 1) = "\", "", "\") & "audio"
    audioDir = ChooseLocalOutFolder(audioDir, "请选择保存音频的本地文件夹")
    If Len(audioDir) = 0 Then
        MsgBox "未选择文件夹，导出已取消。", vbExclamation
        Exit Sub
    End If

    For i = 1 To pres.Slides.Count
        Set sld = pres.Slides(i)

        If Not NotesToAudio(sld, audioDir) Then
            GoTo NextSlide
        End If

NextSlide:
    Next i

    MsgBox "已完成。音频目录：" & audioDir, vbInformation
    
End Sub

' 仅将“当前片子”的备注转为音频并嵌入
Public Sub NotesToAudio_CurrentSlide()

    Dim pres As Presentation
    Dim sld As Slide
    Dim baseDir As String, audioDir As String
    Dim notesText As String, mp3Path As String
    Dim shp As Shape

    Set pres = ActivePresentation
    If pres Is Nothing Then
        MsgBox "未找到活动演示文稿。", vbExclamation: Exit Sub
    End If

    baseDir = pres.path
    If Len(baseDir) = 0 Then
        MsgBox "请先保存演示文稿再运行。", vbExclamation
        Exit Sub
    End If

    ' 取得当前片子（放映视图或普通视图都兼容）
    Set sld = GetActiveSlideSafe()
    If sld Is Nothing Then
        MsgBox "无法取得当前片子。", vbExclamation
        Exit Sub
    End If

    ' audio 目录：PPT 同目录\audio
    audioDir = baseDir & IIf(Right$(baseDir, 1) = "\", "", "\") & "audio"
    audioDir = ChooseLocalOutFolder(audioDir, "请选择保存音频的本地文件夹")
    If Len(audioDir) = 0 Then
        MsgBox "未选择文件夹，导出已取消。", vbExclamation
        Exit Sub
    End If
    
    If NotesToAudio(sld, audioDir) Then
        MsgBox "已为当前片子嵌入新音频：" & audioDir, vbInformation
    Else
        MsgBox "没有为当前片子嵌入音频", vbInformation
    End If
End Sub

' slide的备注转为音频并把音频文件存在audioDir目录下
Private Function NotesToAudio(sld As Slide, audioDir As String) As Boolean
    On Error GoTo EH
    
    Dim notesText As String, mp3Path As String
    Dim shp As Shape
    
    ' 生成文件名：slideNN.mp3
    mp3Path = audioDir & "\slide" & VBA.format$(sld.SlideIndex, "00") & ".mp3"

    ' 读取备注
    notesText = Trim$(GetNotesText(sld))
    'Debug.Print notesText
    If Len(notesText) = 0 Then
        DeleteOldTtsAudio sld
        ' 备注为空：直接将放映时间设置为2秒
        setShow sld, 2#
        Debug.Print "[Skip] Slide " & sld.SlideIndex & " 备注为空 -> 设 2s"
        'MsgBox "备注为空：已将当前片子的放映时间设为 2 秒。", vbInformation
        '由于没有生成音频文件，返回不成功
        NotesToAudio = False
        Exit Function
    End If

    ' 文本转语音（使用你已有的 FA_TTS_ToFile）
    If Not FA_TTS_ToFile(notesText, mp3Path) Then
        Debug.Print "[ERR] Slide " & sld.SlideIndex & "：TTS 失败"
        MsgBox "TTS 失败，请查看立即窗口日志。", vbExclamation
        NotesToAudio = False
        Exit Function
    End If

    ' 为当前片子嵌入音频
    InsertAudio4Slide sld, mp3Path
    NotesToAudio = True
    Exit Function

EH:
    MsgBox "发生错误：" & Err.Number & " - " & Err.Description, vbExclamation
    NotesToAudio = False
End Function

' 向片子 sld 插入音频 mp3Path，并根据音频长度设置放映时间
' tailPadSec：音频播放结束后额外停留的秒数（默认 1 秒）
Public Sub InsertAudio4Slide(ByVal sld As Slide, ByVal mp3Path As String, _
                             Optional ByVal tailPadSec As Single = 1#)
    On Error GoTo EH

    Dim shp As Shape
    Dim eff As Effect
    Dim durSec As Double, advanceSec As Double

    ' 1) 校验文件
    If Len(Dir$(mp3Path, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then
        MsgBox "音频文件不存在：" & vbCrLf & mp3Path, vbExclamation
        Exit Sub
    End If

    ' 2) 清理旧音频（按自定义 Tag）
    DeleteOldTtsAudio sld

    ' 3) 插入音频（AddMediaObject2 → 失败回退 AddMediaObject）
    On Error Resume Next
    Set shp = sld.Shapes.AddMediaObject2(mp3Path, msoFalse, msoTrue, 50, 50)
    If shp Is Nothing Then
        Err.Clear
        Set shp = sld.Shapes.AddMediaObject(mp3Path, msoFalse, msoTrue, 50, 50)
    End If
    On Error GoTo EH

    If shp Is Nothing Then
        Debug.Print "[ERR] Slide " & sld.SlideIndex & "：插入音频失败 -> " & mp3Path
        MsgBox "插入音频失败。", vbExclamation
        Exit Sub
    End If

    ' 4) 标签 + 播放设置
    On Error Resume Next
    shp.Tags.Add TAG_KEY, TAG_VAL
    shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoTrue
    shp.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue
    On Error GoTo EH

    ' 5) 创建“播放媒体”动画：与上一项同时开播（无开播前延时）
    Set eff = sld.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious)
    eff.Timing.TriggerDelayTime = 0    ' 不做开播前延时

    ' 6) 获取音频时长（秒）
    durSec = GetDurationFromShapeSec(shp, mp3Path)

    ' 7) 放映时间 = 音频时长 + 播放后留白；且至少 2 秒
    advanceSec = IIf(durSec > 0#, durSec + tailPadSec, 2#)
    If advanceSec < 2# Then advanceSec = 2#

    ' 8) 设置放映时间
    setShow sld, advanceSec

    Debug.Print "[OK] Slide " & sld.SlideIndex & _
                " 音频=" & format$(durSec, "0.0") & "s, 留白=" & format$(tailPadSec, "0.0") & "s, 放映=" & format$(advanceSec, "0.0") & "s -> " & mp3Path
    Exit Sub

EH:
    Debug.Print "[ERR] InsertAudio4Slide -> "; Err.Number; " "; Err.Description
    MsgBox "插入音频出错：" & Err.Number & vbCrLf & Err.Description, vbExclamation
End Sub

'设置幻灯片放映时间
Private Sub setShow(sld As Slide, advanceSec As Double)
    With sld.SlideShowTransition
        .AdvanceOnTime = True
        .AdvanceTime = advanceSec
        ' 如需禁止点击翻页，取消下一行注释：
        ' .AdvanceOnMouseClick = False
    End With
End Sub

' 兼容放映视图/普通视图，获取“当前片子”
Private Function GetActiveSlideSafe() As Slide
    On Error Resume Next
    If SlideShowWindows.Count > 0 Then
        Set GetActiveSlideSafe = SlideShowWindows(1).View.Slide
        If Not GetActiveSlideSafe Is Nothing Then Exit Function
    End If
    If Not ActiveWindow Is Nothing Then
        Set GetActiveSlideSafe = ActiveWindow.View.Slide
    End If
End Function

'向片子sld插入音频文件mp3Path，并根据音频长度设置片子放映时间
Public Sub InsertAudio4Slide1(ByVal sld As Slide, ByVal mp3Path As String)
    Dim shp As Shape
    Dim durSec As Double, advanceSec As Double
    
    ' 先清理旧的音频文件（相同tag）
    DeleteOldTtsAudio sld
    
    Set shp = InsertAudio(sld, mp3Path)
    If shp Is Nothing Then
        Debug.Print "[ERR] Slide " & sld.SlideIndex & "：插入音频失败"
        MsgBox "插入音频失败。", vbExclamation
        Exit Sub
    End If
    
        ' ? 设置“开始：自动(A)” '2025.11增加
    Dim eff As Effect
    Set eff = sld.TimeLine.MainSequence.AddEffect(Shape:=shp, effectId:=msoAnimEffectMediaPlay, _
                                                  trigger:=msoAnimTriggerAfterPrevious)
    eff.Timing.TriggerDelayTime = 0   ' 0 秒延时 = 立即自动播放

    ' 获取时长（统一“秒”）
    durSec = GetDurationFromShapeSec(shp, mp3Path)

    ' 放映时间 = 音频 +1s，且至少 2s
    advanceSec = IIf(durSec > 0#, durSec + 1#, 2#)
    If advanceSec < 2# Then advanceSec = 2#
    sld.SlideShowTransition.AdvanceOnTime = True
    sld.SlideShowTransition.AdvanceTime = advanceSec
    ' 如需禁止点击翻页，取消下一行注释：
    ' sld.SlideShowTransition.AdvanceOnMouseClick = False

    Debug.Print "[OK] Slide " & sld.SlideIndex & _
                " 时长=" & VBA.format$(advanceSec, "0.0") & "s -> " & mp3Path
End Sub


' -------- 插入音频 --------
Private Function InsertAudio(ByVal sld As Slide, ByVal filePath As String) As Shape
    On Error GoTo EH
    Dim shp As Shape
    Set shp = sld.Shapes.AddMediaObject2(filePath, msoFalse, msoTrue, 50, 50)
    shp.Tags.Add TAG_KEY, TAG_VAL
    
    On Error Resume Next
    shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoTrue
    shp.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue
    On Error GoTo 0
    
    sld.TimeLine.MainSequence.AddEffect shp, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious
    Set InsertAudio = shp
    Exit Function
EH:
    Set InsertAudio = Nothing
End Function

' -------- 时长获取（Shape 优先，失败用 MCI）--------
Private Function GetDurationFromShape(ByVal shp As Shape) As Double
    On Error Resume Next
    Dim mf As Object, v As Variant
    If shp.Type <> msoMedia Then Exit Function

    Set mf = shp.MediaFormat

    ' 尝试获取 Duration
    Err.Clear
    v = CallByName(mf, "Duration", VbGet)
    If Err.Number <> 0 Then
        ' 有些版本用 Length
        Err.Clear
        v = CallByName(mf, "Length", VbGet)
    End If

    If IsNumeric(v) Then GetDurationFromShape = CDbl(v) Else GetDurationFromShape = 0#
End Function

' 统一返回“秒”。能取到形状时长就用；必要时与文件时长对齐；如疑似毫秒则自动 /1000
Private Function GetDurationFromShapeSec(ByVal shp As Shape, ByVal filePath As String) As Double
    On Error Resume Next

    Dim s As Double, m As Double, mf As Object, v As Variant
    If shp.Type = msoMedia Then
        Set mf = shp.MediaFormat
        v = CallByName(mf, "Duration", VbGet)
        If Err.Number <> 0 Then
            Err.Clear
            v = CallByName(mf, "Length", VbGet)
        End If
        If IsNumeric(v) Then s = CDbl(v)
    End If

    ' 文件级别备用（已是“秒”）
    m = FA_GetMp3DurationSec(filePath)

    If s <= 0# And m > 0# Then
        GetDurationFromShapeSec = m
        Exit Function
    End If

    If s > 0# And m > 0# Then
        ' 如果形状时长明显比文件时长大很多，判定为毫秒
        If s > 20# * m Then s = s / 1000#
        ' 两者仍差距较大时，优先用文件时长
        If Abs(s - m) > 1# Then s = m
        GetDurationFromShapeSec = s
        Exit Function
    End If

    ' 只有形状时长可用时：若看起来像毫秒（≥1000），做一次 /1000
    If s >= 1000# Then s = s / 1000#
    GetDurationFromShapeSec = s
End Function

