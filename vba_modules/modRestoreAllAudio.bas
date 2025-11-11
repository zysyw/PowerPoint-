Attribute VB_Name = "modRestoreAllAudio"
Option Explicit

'将指定目录下的音频文件，重新嵌入到片子
Public Sub RestoreAllAudio()
    Dim pres As Presentation, baseDir As String, audioDir As String
    Dim i As Long, sld As Slide, notesText As String, mp3Path As String
    Dim shp As Shape, durSec As Double, advanceSec As Double
    
    If Not ConfirmAction("恢复所有备注", "将尝试从备份恢复全部备注。") Then Exit Sub

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

        mp3Path = audioDir & "\slide" & VBA.format$(i, "00") & ".mp3"

        If Not Len(Dir$(mp3Path, vbNormal Or vbReadOnly)) > 0 Then
            Debug.Print "[ERR] Slide " & i & "：恢复 失败"
            MsgBox "Slide" & i & "找不到音频文件：" & mp3Path, vbInformation
            GoTo NextSlide
        End If

        ' 为当前片子嵌入音频
        InsertAudio4Slide sld, mp3Path
NextSlide:
    Next i

        MsgBox "已完成。音频目录：" & audioDir, vbInformation
End Sub
