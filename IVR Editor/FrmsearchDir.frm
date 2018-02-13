VERSION 5.00
Begin VB.Form frmSearchDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找路径"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "FrmsearchDir.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Tag             =   "1452"
   Begin VB.CommandButton cmdButton 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3540
      TabIndex        =   4
      Tag             =   "1372"
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   3540
      TabIndex        =   3
      Tag             =   "1144"
      Top             =   480
      Width           =   855
   End
   Begin VB.DriveListBox drvList 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   1110
      Width           =   1665
   End
   Begin VB.DirListBox dirList 
      ForeColor       =   &H00FF0000&
      Height          =   930
      Left            =   1800
      TabIndex        =   1
      Top             =   90
      Width           =   1665
   End
   Begin VB.FileListBox fleList 
      ForeColor       =   &H00FF0000&
      Height          =   1260
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   1665
   End
End
Attribute VB_Name = "frmSearchDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTag As String
Private Sub cmdButton_Click(Index As Integer)
    Dim lv_StrPath As String
    
    Select Case Index
    Case 0      '确定
        'lv_StrPath = GetRelativePath(gSystem.strPath_SysVox, fleList.Path)
        lv_StrPath = mdlcommon.AddDirSepMark(dirList.List(dirList.ListIndex))
        sTag = lv_StrPath
    Case 1      '取消
        sTag = ""
    End Select
    Unload Me
End Sub

Private Sub dirList_Change()
    fleList.Path = dirList.List(dirList.ListIndex)
    If fleList.ListCount > 0 Then
        cmdButton(0).Enabled = True
    Else
        cmdButton(0).Enabled = False
    End If
End Sub

Private Sub drvList_Change()
    dirList.Path = drvList.Drive
End Sub

Private Sub fleList_DblClick()
    If cmdButton(0).Enabled Then
        cmdButton_Click 0
    End If
End Sub

Private Sub Form_Load()
    
    '将本窗体置于屏幕中心
    Call mdlcommon.FormCenterPosition(frmSearchDir)
    sTag = ""
    
    '初始路径
    On Error Resume Next
    dirList.Path = gSystem.strPath_SysVox
    If Err.Number <> 0 Then
        Err.Clear
        dirList.Path = gSystem.strPath_Working
    End If
    LoadResStrings Me
    On Error GoTo 0
    
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Trim(sTag) <> "" Then
        frmOptions.txtVoxPath = sTag
    End If
End Sub

'Get Relative path from root
'
Private Function GetRelativePath(f_Root As String, f_Current As String) As String
    Dim lv_Str As String
    Dim lv_Loc As Integer

    lv_Str = ""
    lv_Loc = InStr(f_Current, f_Root)
    If lv_Loc > 0 Then
        lv_Str = Mid(f_Current, lv_Loc + Len(f_Root))
        lv_Str = mdlcommon.AddDirSepMark(lv_Str)
    Else
    End If
    
    GetRelativePath = lv_Str
    
End Function
