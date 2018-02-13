VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{FB712C5B-9403-4F42-B022-73ABA0F01A72}#1.0#0"; "VOXExpCtrl.ocx"
Begin VB.Form frmRecordVoice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "录音"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmRecordVoice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1717"
   Begin VOXEXPCTRLLib.VOXExpCtrl RecCtl 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin MSComDlg.CommonDialog dlgSaveFile 
      Left            =   5760
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFilePath 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   6930
      Picture         =   "frmRecordVoice.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "1721"
      Top             =   3180
      Width           =   1005
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   855
      Left            =   6930
      Picture         =   "frmRecordVoice.frx":2136
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "1720"
      Top             =   2340
      Width           =   1005
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   855
      Left            =   6930
      Picture         =   "frmRecordVoice.frx":25B8
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "1719"
      Top             =   1500
      Width           =   1005
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record"
      Height          =   855
      Left            =   6930
      Picture         =   "frmRecordVoice.frx":2A3A
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "1718"
      Top             =   660
      Width           =   1005
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4005
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmRecordVoice.frx":2EBC
      Top             =   0
      Width           =   6915
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "停止"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   6930
      TabIndex        =   5
      Top             =   0
      Width           =   1005
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRecordVoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modify Log
'Last Modified : 2007-11-01
'[1] : 录音界面标签歧义  Oct,24,2K7
'[2] : 文件后缀名判断代码不健壮

Option Explicit
Dim strScriptOut As String
Dim tempPath As String
 
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPause_Click()
    Me.lblStatus.Visible = True
    Me.lblStatus.Caption = "暂停"
    RecCtl.PauseRecord
End Sub

Private Sub cmdRecord_Click()
    Me.lblStatus.Visible = True
    Me.lblStatus.Caption = "录音"
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Start Recording..", "Filepath:" & txtFilePath)
    
    'VOX | WAV format
    'Michael Modified @ 2007-11-01 -> 修改后缀名判别方法
    If LCase(Right(txtFilePath, 3)) = "vox" Then                     'Vox Format
        RecCtl.StartRecord txtFilePath
    'Michael Modified @ 2007-11-01 -> 修改后缀名判别方法
    ElseIf LCase(Right(txtFilePath, 3)) = "wav" Then
            tempPath = Left(txtFilePath, Len(txtFilePath) - 3) & "vox"   'Wav Foramt
            RecCtl.StartRecord tempPath
    Else
        '资源文件类型不被支持或文件路径错误,请更改后重新录制...
        Message "E140"
        'Mike added @ 2008-7-7
        Call WriteLogMessage(0, enu_Warnning, "Exception Occured during Recording...", "Filepath:" & txtFilePath)
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub cmdStop_Click()
    Me.lblStatus.Caption = "停止"
    RecCtl.StopRecord
    
    'Michael Modified @ 2007-11-01 -> 修改后缀名判别方法
    If LCase(Right(txtFilePath, 3)) = "wav" Then       'Wav Foramt
        RecCtl.VOXFile2Wave tempPath, txtFilePath
        Call RemoveFile(tempPath)
    End If
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Finish Recording!", "Filepath:" & txtFilePath)
End Sub

Private Sub Form_Load()
    LoadResStrings Me
    'Modified @ 2007-11-01 -> 修改状态标签
    Me.lblStatus.Visible = False
End Sub
