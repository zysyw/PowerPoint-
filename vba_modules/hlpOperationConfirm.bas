Attribute VB_Name = "hlpOperationConfirm"
Option Explicit

' === 通用确认框 ===
Public Function ConfirmAction(ByVal title As String, ByVal detail As String) As Boolean
    Dim msg As String, ans As VbMsgBoxResult
    msg = "此操作不可撤销！" & vbCrLf & vbCrLf & detail & vbCrLf & vbCrLf & "是否继续？"
    ans = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, "确认 - " & title)
    ConfirmAction = (ans = vbYes)
End Function

' === 统计将受影响的对象（使用标签） ===
Public Function CountTtsAudios(Optional onlyTagged As Boolean = True) As Long
    Dim sld As Slide, shp As Shape, n As Long
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If Not onlyTagged Or TagEquals(shp, TAG_KEY, TAG_VAL) Then n = n + 1
            End If
        Next
    Next
    CountTtsAudios = n
End Function

Public Function CountSlidesWithNotes() As Long
    Dim sld As Slide, n As Long
    For Each sld In ActivePresentation.Slides
        If Len(GetNotesText(sld)) > 0 Then n = n + 1
    Next
    CountSlidesWithNotes = n
End Function

Public Function TagEquals(ByVal shp As Shape, ByVal key As String, ByVal expect As String) As Boolean
    Dim v As String, j As Long
    On Error Resume Next
    v = shp.Tags(key)
    If Err.Number <> 0 Then
        Err.Clear
        For j = 1 To shp.Tags.Count
            If StrComp(shp.Tags.Name(j), key, vbTextCompare) = 0 Then
                v = shp.Tags.Value(j): Exit For
            End If
        Next
    End If
    TagEquals = (StrComp(v, expect, vbBinaryCompare) = 0)
End Function


' 兼容不同版本的 Tags 访问：优先用“按名取值”，失败则遍历
Private Function GetTagValueByName(ByVal shp As Shape, ByVal key As String) As String
    On Error Resume Next
    Dim v As String, j As Long
    ' 有些版本支持 shp.Tags("name")
    v = shp.Tags(key)
    If Err.Number = 0 And Len(v) > 0 Then
        GetTagValueByName = v
        Exit Function
    End If
    Err.Clear
    ' 通用遍历：通过 Name(index)/Value(index)
    For j = 1 To shp.Tags.Count
        If StrComp(shp.Tags.Name(j), key, vbTextCompare) = 0 Then
            GetTagValueByName = shp.Tags.Value(j)
            Exit Function
        End If
    Next j
End Function
