Attribute VB_Name = "oldInsertAudioBySlide"
Option Explicit

'''从audio子文件夹内的mp3文件按照标题顺序嵌入PPT中
Sub EmbedMp3FilesByOrder()
    Dim pptPath As String
    Dim audioDir As String
    Dim audioFilePath As String
    Dim i As Integer
    Dim oSlide As Slide
    Dim oShp As Shape
    Dim oEffect As Effect

    ' 获取当前演示文稿的路径
    pptPath = ActivePresentation.path
    ' 构建audio子目录的路径
    audioDir = pptPath & "\audio\"

    ' 检查audio目录是否存在
    If Len(Dir(audioDir, vbDirectory)) = 0 Then
        MsgBox "未找到audio目录。", vbCritical
        Exit Sub
    End If
    
    ' 按照幻灯片顺序嵌入音频文件
    For i = 1 To ActivePresentation.Slides.Count
        audioFilePath = audioDir & i & ".mp3"
        
        ' 检查文件是否存在
        If Len(Dir(audioFilePath)) > 0 Then
            Set oSlide = ActivePresentation.Slides(i)
            ' 在当前幻灯片中嵌入音频文件
            Set oShp = oSlide.Shapes.AddMediaObject2( _
                fileName:=audioFilePath, LinkToFile:=msoFalse, SaveWithDocument:=True, _
                Left:=5, Top:=5)
            ' 设置为开始时自动播放
            Set oEffect = oSlide.TimeLine.MainSequence.AddEffect(oShp, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious)
            ' 设置为放映时隐藏
            oShp.AnimationSettings.PlaySettings.HideWhileNotPlaying = True
        Else
            ' 如果对应编号的音频文件不存在，则跳出循环
            Exit For
        End If
    Next i

    ' 完成消息
    MsgBox "音频文件嵌入完成。已处理 " & i - 1 & " 个音频文件。", vbInformation
    
    MuteAllEmbeddedVideos
    
End Sub

Sub MuteAllEmbeddedVideos()
    Dim sld As Slide
    Dim shp As Shape

    ' 遍历当前演示文稿中的所有幻灯片
    For Each sld In ActivePresentation.Slides
        ' 遍历幻灯片中的所有形状
        For Each shp In sld.Shapes
            ' 检查形状是否为媒体对象
            If shp.Type = msoMedia Then
                ' 检查媒体类型是否为视频
                If shp.MediaType = ppMediaTypeMovie Then
                    ' 设置视频静音
                    shp.MediaFormat.Muted = True
                End If
            End If
        Next shp
    Next sld
    
    MsgBox "所有嵌入的视频已设置为静音状态。", vbInformation
End Sub
