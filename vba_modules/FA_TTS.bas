Attribute VB_Name = "FA_TTS"
Option Explicit
' 依赖：FA_Config.bas

' 对外函数：把 text 合成为 MP3/WAV/Opus 文件
Public Function FA_TTS_ToFile(ByVal text As String, ByVal outPath As String, _
                              Optional ByVal format As String = "mp3", _
                              Optional ByVal sampleRate As Long = 44100, _
                              Optional ByVal mp3Bitrate As Long = 128, _
                              Optional ByVal temperature As Double = 0.6, _
                              Optional ByVal top_p As Double = 0.7, _
                              Optional ByVal latency As String = "normal", _
                              Optional ByVal timeoutMs As Long = 180000) As Boolean
    On Error GoTo EH

    Dim http As Object, url As String, body As String
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    url = TrimEndSlash(FA_BASE_URL) & "/v1/tts"

    body = "{""text"":" & QuoteJSON(text) & _
           ",""reference_id"":" & QuoteJSON(FA_REFERENCE_ID) & _
           ",""format"":" & QuoteJSON(LCase$(format)) & _
           ",""sample_rate"":" & CStr(sampleRate)

    If LCase$(format) = "mp3" Then body = body & ",""mp3_bitrate"":" & CStr(mp3Bitrate)

    body = body & ",""temperature"":" & JsonNum(temperature) & _
                 ",""top_p"":" & JsonNum(top_p) & _
                 ",""latency"":" & QuoteJSON(latency) & "}"

    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & FA_TOKEN
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "*/*"
    http.SetRequestHeader "model", FA_MODEL
    http.SetRequestHeader "User-Agent", "VBA-FishAudio/1.0"

    http.Send body

    If http.Status = 200 Then
        EnsureParentFolderExists outPath
        SaveBinaryFile outPath, http.ResponseBody
        FA_TTS_ToFile = True
    Else
        Debug.Print "[HTTP] TTS " & http.Status & " / " & Left$(http.ResponseText, 300)
        FA_TTS_ToFile = False
    End If
    Exit Function
EH:
    Debug.Print "[ERR] FA_TTS_ToFile -> "; Err.Number; " "; Err.Description
    FA_TTS_ToFile = False
End Function

' --------- 本模块用到的小工具 ----------
Private Sub SaveBinaryFile(ByVal path As String, ByVal bytes As Variant)
    On Error GoTo EH
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.SaveToFile path, 2
    stm.Close
    Exit Sub
EH:
    Debug.Print "[SaveBinary] 写入失败 -> "; path; " | Err "; Err.Number; ": "; Err.Description
End Sub

Private Sub EnsureParentFolderExists(ByVal filePath As String)
    Dim fso As Object, parentPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    parentPath = fso.GetParentFolderName(filePath)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then fso.CreateFolder parentPath
    End If
End Sub

Private Function TrimEndSlash(ByVal s As String) As String
    If Len(s) > 0 And Right$(s, 1) = "/" Then TrimEndSlash = Left$(s, Len(s) - 1) Else TrimEndSlash = s
End Function

Private Function QuoteJSON(ByVal s As String) As String
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", "\""")
    s = Replace$(s, vbCrLf, "\n")
    s = Replace$(s, vbCr, "\n")
    s = Replace$(s, vbLf, "\n")
    QuoteJSON = """" & s & """"
End Function

Private Function JsonNum(ByVal d As Double) As String
    Dim s As String: s = CStr(d)
    JsonNum = Replace$(s, ",", ".")
End Function


