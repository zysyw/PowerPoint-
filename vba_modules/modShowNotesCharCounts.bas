Attribute VB_Name = "modShowNotesCharCounts"
'===================== 备注字符统计（当前页 + 全文）=====================
Option Explicit

' ------------- 工具函数：统计字符数 -------------
Private Function CharCount(textIn As String) As Long
    CharCount = Len(textIn)
End Function

' ------------- 主过程：同时统计当前页与全部 -------------
Sub ShowNotesCharCounts()
    Dim sld As Slide
    Dim totalChars As Long
    Dim currentChars As Long
    Dim noteTxt As String
    Dim curSlideID As Long
    
    ' 获取当前所选幻灯片的 SlideID（比索引更可靠，放映视图也适用）
    On Error Resume Next
    curSlideID = ActiveWindow.View.Slide.SlideID
    On Error GoTo 0
    
    For Each sld In ActivePresentation.Slides
        On Error Resume Next   '防止没有备注占位符时报错
        noteTxt = GetNotesText(sld)
        On Error GoTo 0
        
        Dim thisChars As Long
        thisChars = CharCount(noteTxt)
        totalChars = totalChars + thisChars
        
        If sld.SlideID = curSlideID Then
            currentChars = thisChars
        End If
    Next sld
    
    MsgBox "当前幻灯片备注字符数: " & currentChars & vbCrLf & _
           "所有幻灯片备注字符总数: " & totalChars, _
           vbInformation, "备注统计"
End Sub
'======================================================================


