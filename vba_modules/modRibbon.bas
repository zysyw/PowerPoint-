Attribute VB_Name = "modRibbon"
Option Explicit
Public g_Ribbon As IRibbonUI

'Callback for customUI.onLoad
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set g_Ribbon = ribbon
End Sub

'Callback for btnCount onAction
Sub Ribbon_CountNotes(control As IRibbonControl)
    Call ShowNotesCharCounts
End Sub

'Callback for btnExport onAction
Sub Ribbon_ExportNotes(control As IRibbonControl)
    Call ExportNotes
End Sub

'Callback for btnClear onAction
Sub Ribbon_ClearNotes(control As IRibbonControl)
    If Not ConfirmAction("清空所有备注", "将删除所有幻灯片的备注内容。") Then Exit Sub
    Call ClearAllNotes
End Sub

'Callback for btnRestore onAction
Sub Ribbon_RestoreNotes(control As IRibbonControl)
    If Not ConfirmAction("恢复所有备注", "将尝试从备份恢复全部备注。") Then Exit Sub
    Call RestoreNotesFromBackup
End Sub

'Callback for btnIns onAction
Sub Ribbon_InsertAudio(control As IRibbonControl)
    If Not ConfirmAction("为所有幻灯片插入音频", "将根据备注内容为每张幻灯片生成并插入音频，可能需要较长时间。") Then Exit Sub
    Call NotesToAudio_AllSlides
End Sub

'Callback for btninsCur onAction
Sub Ribbon_DeleteAudioCurrent(control As IRibbonControl)
    If Not ConfirmAction("删除当前幻灯片音频", "此操作将删除当前幻灯片中已插入的音频文件。") Then Exit Sub
    Call NotesToAudio_CurrentSlide
End Sub

'Callback for btnRestoreAllAudio onAction
Sub Ribbon_RestoreAllAudio(control As IRibbonControl)
    If Not ConfirmAction("恢复所有音频", _
        "将尝试从备份恢复所有幻灯片的音频文件。") Then Exit Sub
    Call RestoreAllAudio
End Sub

'Callback for btnDelAll onAction
Sub Ribbon_DeleteAudioAll(control As IRibbonControl)
    If Not ConfirmAction("删除所有音频", "将删除所有幻灯片中的音频文件。") Then Exit Sub
    Call DeleteAllAudio
End Sub

Public Sub Ribbon_SettingAPI(control As IRibbonControl)
    frmFASettings.Show vbModal
End Sub

Public Sub Ribbon_Help(ctrl As IRibbonControl)
    frmFAHelp.Show vbModal
End Sub
