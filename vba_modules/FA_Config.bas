Attribute VB_Name = "FA_Config"
Option Explicit

' === Fish Audio 基本配置 ===
Public Const FA_BASE_URL     As String = "https://api.fish.audio"
Public Const FA_MODEL        As String = "s1"                               ' "s1" | "speech-1.6" | "speech-1.5"

Public FA_TOKEN        As String                 ' ← API密钥
Public KEY_MODEL_TITLE As String             ' ← 参考音色Title
Public FA_REFERENCE_ID As String             ' ← 参考音色ID

' === 网络配置（需要代理就开）===
Public Const FA_USE_PROXY    As Boolean = True
Public Const FA_PROXY_ADDR   As String = "127.0.0.1:10808"                  ' v2rayN 常见 HTTP 端口

' TLS 位掩码（WinHTTP）
Public Const TLS1  As Long = &H80
Public Const TLS11 As Long = &H200
Public Const TLS12 As Long = &H800

