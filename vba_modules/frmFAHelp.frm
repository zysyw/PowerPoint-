VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFAHelp 
   Caption         =   "UserForm1"
   ClientHeight    =   4980
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7660
   OleObjectBlob   =   "frmFAHelp.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmFAHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "帮助"
    Me.StartUpPosition = 1 ' CenterOwner

    ' 显示你要求的三条说明
    Dim msg As String
    msg = "1、使用前需先进行设置，输入API密钥和模型名称，如使用自建模型需选中“克隆声音”，然后点击获取ID，正确得到ID后确定即可。" & vbCrLf & vbCrLf & _
          "2、默认导出的目录与PPT文件相同，生成的音频文件在PPT文件目录下的audio子目录下。" & vbCrLf & vbCrLf & _
          "3、程序支持OneDrive目录，使用前请检查正确的“OneDrive”环境变量。"
    txtHelp.text = msg

    ' 让文本框随窗体大小自适应（可选）
    txtHelp.Left = 12
    txtHelp.Top = 12
    txtHelp.Width = Me.InsideWidth - 24
    txtHelp.Height = Me.InsideHeight - 48

    btnClose.Caption = "关闭"
    btnClose.Left = Me.InsideWidth - btnClose.Width - 12
    btnClose.Top = Me.InsideHeight - btnClose.Height - 12
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    txtHelp.Width = Me.InsideWidth - 24
    txtHelp.Height = Me.InsideHeight - 48
    btnClose.Left = Me.InsideWidth - btnClose.Width - 12
    btnClose.Top = Me.InsideHeight - btnClose.Height - 12
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


