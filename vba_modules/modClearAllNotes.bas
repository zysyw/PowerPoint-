Attribute VB_Name = "modClearAllNotes"
Option Explicit

Sub ClearAllNotes()
    Dim sld As Slide, shp As Shape
    Dim cleared As Long, tp As PpPlaceholderType
    Dim n As Long: n = CountSlidesWithNotes()
    
    If Not ConfirmAction("清空所有备注", "将清空 " & n & " 张幻灯片的备注文本。") Then Exit Sub
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.NotesPage.Shapes
            If shp.Type = msoPlaceholder Then
                On Error Resume Next
                tp = shp.PlaceholderFormat.Type
                On Error GoTo 0
                ' 只针对“备注正文”占位符
                If tp = ppPlaceholderBody Then
                    If shp.HasTextFrame Then
                        shp.TextFrame.TextRange.text = ""
                        cleared = cleared + 1
                    End If
                    Exit For ' 本页备注已处理，下一页
                End If
            End If
        Next shp
    Next sld
    
    MsgBox "已清空备注正文，共处理 " & cleared & " 张幻灯片。"
End Sub

