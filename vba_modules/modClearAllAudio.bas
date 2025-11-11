Attribute VB_Name = "modClearAllAudio"
Option Explicit

Sub DeleteAllAudio()
    Dim sld As Slide
    Dim i As Long, n As Long, deleted As Long

    ' 仅统计“带标签的 TTS 音频”
    n = CountTtsAudios(True)
    If n = 0 Then
        MsgBox "未发现已标记的 TTS 音频可删除。", vbInformation: Exit Sub
    End If

    If Not ConfirmAction("删除所有音频", "将删除 " & n & " 个已嵌入的 TTS 音频（仅删除带标签的项目）。") Then Exit Sub

    For Each sld In ActivePresentation.Slides
        DeleteOldTtsAudio sld
        deleted = deleted + 1
    Next sld

    MsgBox "已删除 " & deleted & " / " & n & " 个页面的 TTS 音频。", vbInformation, "操作完成"
End Sub

'-------删除旧音频----------
Public Sub DeleteOldTtsAudio(ByVal sld As Slide)
    Dim i As Long, shp As Shape
    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)
        If shp.Type = msoMedia Then
            On Error Resume Next                  ' 某些版本 shp.MediaType 可能抛错
            If shp.MediaType = ppMediaTypeSound Then
                If TagEquals(shp, TAG_KEY, TAG_VAL) Then
                    shp.Delete
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Sub
