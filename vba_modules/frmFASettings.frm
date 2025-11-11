VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFASettings 
   Caption         =   "设置"
   ClientHeight    =   3760
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5440
   OleObjectBlob   =   "frmFASettings.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmFASettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 是否点击了“确定”
Public IsOK As Boolean

Private Sub UserForm_Initialize()
    ' 从全局变量带入
    On Error Resume Next
    txtToken.text = FA_TOKEN
    txtModelTitle.text = KEY_MODEL_TITLE
    txtModelID.text = FA_REFERENCE_ID
    chkClone.Value = True
    lblStatus.Caption = ""
    On Error GoTo 0
End Sub

Private Sub btnFetchId_Click()
    Dim token As String, title As String
    Dim url As String, json As String, id As String
    Dim selfQS As String

    lblStatus.Caption = ""
    token = Trim$(txtToken.text)
    title = Trim$(txtModelTitle.text)

    If Len(token) = 0 Then
        MsgBox "请输入 API Token。", vbExclamation
        txtToken.SetFocus
        Exit Sub
    End If
    If Len(title) = 0 Then
        MsgBox "请输入模型 Title。", vbExclamation
        txtModelTitle.SetFocus
        Exit Sub
    End If

    btnFetchId.Enabled = False
    lblStatus.Caption = "正在获取 ID..."
    Me.Repaint
    
    ' ★ 根据是否“克隆声音”决定 self=true/false
    selfQS = IIf(chkClone.Value, "true", "false")

    ' 先试 /model?self=true&title=...
    url = TrimEndSlash(FA_BASE_URL) & "/model?title=" & UrlEncodeUTF8(title) & "&self=" & selfQS
    json = HttpGetJson(url, token)

    ' 失败则回退 /v1/model?...
    If Len(json) = 0 Then
        url = TrimEndSlash(FA_BASE_URL) & "/v1/model?title=" & UrlEncodeUTF8(title) & "&self=" & selfQS
        json = HttpGetJson(url, token)
    End If

    If Len(json) = 0 Then
        lblStatus.Caption = ""
        MsgBox "获取失败：无法从服务端读取模型信息。" & vbCrLf & _
               "可能原因：Token 无效、网络/代理问题、或模型名称不正确。", vbExclamation
        btnFetchId.Enabled = True
        Exit Sub
    End If

    id = ExtractFirstId(json)
    If Len(id) = 0 Then
        lblStatus.Caption = ""
        MsgBox "未找到名称为 """ & title & """ 的模型。", vbExclamation
        btnFetchId.Enabled = True
        Exit Sub
    End If

    txtModelID.text = id
    lblStatus.Caption = "已获取 ID。"
    btnFetchId.Enabled = True
End Sub

Private Sub btnOK_Click()
    ' 把文本框里的值写回全局变量
    If Len(Trim$(txtToken.text)) = 0 Then
        MsgBox "请先输入 API Token。", vbExclamation
        txtToken.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtModelID.text)) = 0 Then
        MsgBox "请先点击“获取ID”并成功获取。", vbExclamation
        Exit Sub
    End If

    FA_TOKEN = Trim$(txtToken.text)
    FA_REFERENCE_ID = Trim$(txtModelID.text)

    IsOK = True
    Unload Me
End Sub

Private Sub btnCancel_Click()
    IsOK = False
    Unload Me
End Sub

' ========== HTTP / 工具（本窗体专用；不做持久化） ==========
Private Function HttpGetJson(ByVal url As String, ByVal token As String) As String
    On Error GoTo EH
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "User-Agent", "VBA-FishAudio/1.0"
    http.Send

    If http.Status = 200 Then
        Dim ctype As String, charset As String
        ctype = LCase$(http.GetResponseHeader("Content-Type"))
        charset = ParseCharset(ctype)
        If Len(charset) = 0 Then charset = "utf-8"  ' 默认按 UTF-8
        HttpGetJson = BytesToText(http.ResponseBody, charset)
    Else
        Debug.Print "[HTTP] "; http.Status; " - "; url
        Debug.Print Left$(BytesToText(http.ResponseBody, "utf-8"), 300)
        HttpGetJson = ""
    End If
    Exit Function
EH:
    Debug.Print "[ERR] HttpGetJson "; Err.Number; " "; Err.Description
    HttpGetJson = ""
End Function

Private Function UrlEncodeUTF8(ByVal s As String) As String
    Dim stm As Object, vBytes As Variant, i As Long, ch As Integer, t As String
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.charset = "utf-8": stm.Open: stm.WriteText s: stm.Position = 0
    stm.Type = 1: vBytes = stm.Read: stm.Close
    For i = 0 To UBound(vBytes)
        ch = vBytes(i)
        If (ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or ch = 45 Or ch = 95 Or ch = 46 Or ch = 126 Then
            t = t & Chr$(ch)
        ElseIf ch = 32 Then
            t = t & "%20"
        Else
            t = t & "%" & Right$("0" & Hex$(ch), 2)
        End If
    Next
    UrlEncodeUTF8 = t
End Function

' 从 Content-Type 里解析 charset=...
Private Function ParseCharset(ByVal ctype As String) As String
    Dim p As Long
    p = InStr(1, ctype, "charset=", vbTextCompare)
    If p > 0 Then
        ParseCharset = LCase$(Trim$(Split(Mid$(ctype, p + 8), ";")(0)))
        If ParseCharset = "gbk" Then ParseCharset = "gb2312" ' ADODB 不识别 gbk，用 gb2312 代替
    End If
End Function

' 把字节数组按指定字符集转成文本（支持 utf-8 / gb2312 等）
Private Function BytesToText(ByVal bytes As Variant, ByVal charset As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2
    On Error Resume Next
    stm.charset = charset
    If Err.Number <> 0 Then
        Err.Clear
        stm.charset = "utf-8"  ' 兜底
    End If
    On Error GoTo 0
    BytesToText = stm.ReadText(-1)
    stm.Close
End Function

' 兼容 "id" 与 "_id"
Private Function ExtractFirstId(ByVal json As String) As String
    On Error GoTo EH
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.Regexp")
    re.Global = True: re.IgnoreCase = True: re.Multiline = True
    re.pattern = """_?id""\s*:\s*""([^""]+)"""
    Set ms = re.Execute(json)
    If ms.Count > 0 Then ExtractFirstId = ms(0).SubMatches(0)
    Exit Function
EH:
    ExtractFirstId = ""
End Function

Private Function TrimEndSlash(ByVal s As String) As String
    If Len(s) > 0 And Right$(s, 1) = "/" Then TrimEndSlash = Left$(s, Len(s) - 1) Else TrimEndSlash = s
End Function




