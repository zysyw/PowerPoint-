Attribute VB_Name = "FA_MCI"
' ===== Module: FA_MCI.bas =====
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function mciSendStringA Lib "winmm.dll" ( _
        ByVal lpstrCommand As String, _
        ByVal lpstrReturnString As String, _
        ByVal uReturnLength As LongPtr, _
        ByVal hwndCallback As LongPtr) As Long
#Else
    Private Declare Function mciSendStringA Lib "winmm.dll" ( _
        ByVal lpstrCommand As String, _
        ByVal lpstrReturnString As String, _
        ByVal uReturnLength As Long, _
        ByVal hwndCallback As Long) As Long
#End If

Public Function FA_GetMp3DurationSec(ByVal path As String) As Double
    On Error GoTo EH
    Dim aliasName As String, rc As Long
    Dim buf As String * 64, outStr As String

    aliasName = "mp3_" & Hex(Timer * 1000)

    rc = mciSendStringA("open """ & path & """ type mpegvideo alias " & aliasName, vbNullString, 0, 0)
    If rc <> 0 Then GoTo fail

    rc = mciSendStringA("set " & aliasName & " time format milliseconds", vbNullString, 0, 0)
    rc = mciSendStringA("status " & aliasName & " length", buf, Len(buf), 0)
    If rc <> 0 Then GoTo fail

    outStr = Trim$(Replace$(buf, Chr$(0), ""))
    If IsNumeric(outStr) Then FA_GetMp3DurationSec = CDbl(outStr) / 1000#
Cleanup:
    mciSendStringA "close " & aliasName, vbNullString, 0, 0
    Exit Function
fail:
    FA_GetMp3DurationSec = 0#
    GoTo Cleanup
EH:
    FA_GetMp3DurationSec = 0#
    GoTo Cleanup
End Function


