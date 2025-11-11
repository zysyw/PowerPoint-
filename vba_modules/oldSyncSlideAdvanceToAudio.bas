Attribute VB_Name = "oldSyncSlideAdvanceToAudio"
Option Explicit

Sub SyncSlideAdvanceToAudio()
    Dim sld As Slide
    Dim shp As Shape
    Dim audioDur As Single
    
    For Each sld In ActivePresentation.Slides
        audioDur = 0
        
        ' 在幻灯片的 Shapes 中寻找第一个音频
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
              If shp.MediaType = ppMediaTypeSound Then
                ' MediaFormat.Length 返回媒体时长（毫秒）
                audioDur = shp.MediaFormat.Length / 1000   ' 转为秒
                Exit For
              End If
            End If
        Next shp
        
        ' 设置自动切换与音频时长同步
        With sld.SlideShowTransition
            .AdvanceOnClick = msoFalse
            If audioDur > 0 Then
                .AdvanceOnTime = msoTrue
                .AdvanceTime = audioDur
            Else
                .AdvanceOnTime = msoFalse
            End If
        End With
    Next sld
    
    MsgBox "已同步所有幻灯片的自动切换时间与音频时长。", vbInformation, "设置完成"
End Sub

