VERSION 5.00
Begin VB.Form frm_040 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "建立留言"
   ClientHeight    =   5595
   ClientLeft      =   3300
   ClientTop       =   2670
   ClientWidth     =   8010
   Icon            =   "frm_040.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1281"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   4230
      Picture         =   "frm_040.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "1948"
      Top             =   5100
      Width           =   333
   End
   Begin VB.Frame Frame4 
      Caption         =   "变量"
      Height          =   2625
      Left            =   4680
      TabIndex        =   40
      Tag             =   "1358"
      Top             =   60
      Width           =   3255
      Begin VB.ComboBox Cb_VarRecDuration 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   1545
      End
      Begin VB.ComboBox Cb_VarAppField 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1740
         Width           =   1545
      End
      Begin VB.ComboBox Cb_usevar 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1545
      End
      Begin VB.ComboBox Cb_VarFileName 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   1545
      End
      Begin VB.ComboBox Cb_VarAppField 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1380
         Width           =   1545
      End
      Begin VB.ComboBox Cb_VarAppField 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2100
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "记录留言时长"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Tag             =   "2121"
         Top             =   1020
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "指定座席"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Tag             =   "1616"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "记录文件名"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Tag             =   "1617"
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "附加字段1"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   43
         Tag             =   "1618"
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "附加字段2"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   42
         Tag             =   "1619"
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "附加字段3"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   41
         Tag             =   "1620"
         Top             =   2160
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1065
      Left            =   4680
      TabIndex        =   37
      Top             =   4320
      Width           =   3255
      Begin VB.TextBox Txt_Description 
         Height          =   675
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   270
         Width           =   3045
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   4680
      TabIndex        =   33
      Tag             =   "1164"
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2790
         Picture         =   "frm_040.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "1145"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2790
         Picture         =   "frm_040.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "1145"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2790
         Picture         =   "frm_040.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox T_vox_play 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   19
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox T_nd_parent 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   21
         Top             =   600
         Width           =   1185
      End
      Begin VB.TextBox T_nd_child 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   23
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "操作提示语音"
         Height          =   180
         Left            =   180
         TabIndex        =   36
         Tag             =   "1266"
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label20 
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Tag             =   "1169"
         Top             =   690
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "子节点ID"
         Height          =   180
         Left            =   180
         TabIndex        =   34
         Tag             =   "1229"
         Top             =   1050
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   4755
      Left            =   60
      TabIndex        =   27
      Tag             =   "1136"
      Top             =   60
      Width           =   4545
      Begin VB.CheckBox chkNotifyPL 
         Caption         =   "通知PL录音系统"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Tag             =   "2125"
         Top             =   2610
         Width           =   2055
      End
      Begin VB.ComboBox cmbNotifyInterval 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         TabIndex        =   8
         ToolTipText     =   "0 - 65535 秒"
         Top             =   3000
         Width           =   1965
      End
      Begin VB.CheckBox chkToneOn 
         Caption         =   "留言开始提示音"
         Height          =   315
         Left            =   2430
         TabIndex        =   6
         Tag             =   "2122"
         Top             =   2610
         Width           =   2055
      End
      Begin VB.ComboBox cboVMSType 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4110
         Width           =   1965
      End
      Begin VB.ComboBox Cb_log 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   3390
         Width           =   1965
      End
      Begin VB.ComboBox cmbVMSClass 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "请选择"
         Top             =   3750
         Width           =   1965
      End
      Begin VB.ComboBox cmbMinRecLen 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         TabIndex        =   1
         ToolTipText     =   "0 - 255 秒"
         Top             =   1080
         Width           =   1965
      End
      Begin VB.ComboBox cmbMaxSilence 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         TabIndex        =   2
         ToolTipText     =   "0 - 60 秒"
         Top             =   1440
         Width           =   1965
      End
      Begin VB.ComboBox Cb_playclear 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "请选择"
         Top             =   1800
         Width           =   1965
      End
      Begin VB.ComboBox Cb_breakkey 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "请选择"
         Top             =   2160
         Width           =   1965
      End
      Begin VB.ComboBox Cb_timeout 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2430
         TabIndex        =   0
         ToolTipText     =   "0 - 65535 秒"
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox T_n_no 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox T_n_id 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "留言进行提示间隔(秒)"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Tag             =   "2123"
         Top             =   3060
         Width           =   1950
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "留言文件类型"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   50
         Tag             =   "1800"
         Top             =   4200
         Width           =   1080
      End
      Begin VB.Label Lbl_log 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Left            =   180
         TabIndex        =   49
         Tag             =   "1245"
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "留言分类"
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Tag             =   "1690"
         Top             =   3840
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "按键中断"
         Height          =   180
         Left            =   180
         TabIndex        =   47
         Tag             =   "1264"
         Top             =   2250
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "最短录音时长(秒)"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   46
         Tag             =   "1698"
         Top             =   1170
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "最大静音时长(秒)"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   39
         Tag             =   "1657"
         Top             =   1530
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "放音清空标志"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   38
         Tag             =   "1244"
         Top             =   1890
         Width           =   1080
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   2250
         TabIndex        =   32
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "录音时间长(秒)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   28
         Tag             =   "1282"
         Top             =   810
         Width           =   1260
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   2400
      TabIndex        =   26
      Tag             =   "1144"
      Top             =   5130
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   1200
      TabIndex        =   25
      Top             =   5130
      Width           =   1035
   End
End
Attribute VB_Name = "frm_040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_040.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：2007-02-28
'//文件描述：建立流言
'//---------------------------------
'//修改版本 : V1.0
'//修改内容 : 添加留言类型选项
'//修改日期 : 06-29-07
'//修改人   : Michael
'//----------------------------------
'//最后修改版本 : V1.2
'//最后修改内容 : 修改最大留言时长
'//最后修改日期 : 07-04-07
'//最后修改人   : Michael
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_breakkey_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_playclear_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_GotFocus()
    Cb_timeout.SelStart = 0
    Cb_timeout.SelLength = Len(Cb_timeout)
End Sub

Private Sub Cb_timeout_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Cb_usevar_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_VarAppField_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Cb_VarFileName_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_VarRecDuration_Click()
    f_DataChanged = True
End Sub

'Michael Added @ 06-29-07
Private Sub cboVMSType_Click()
    f_DataChanged = True
End Sub

'' Sun added 2012-04-18
Private Sub chkNotifyPL_Click()
    f_DataChanged = True
End Sub

Private Sub chkToneOn_Click()
    f_DataChanged = True
End Sub

Private Sub cmbMaxSilence_Change()
    f_DataChanged = True
End Sub

Private Sub cmbMaxSilence_Click()
    f_DataChanged = True
End Sub

Private Sub cmbMaxSilence_GotFocus()
    cmbMaxSilence.SelStart = 0
    cmbMaxSilence.SelLength = Len(cmbMaxSilence)
End Sub

Private Sub cmbMaxSilence_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub cmbMinRecLen_Change()
    f_DataChanged = True
End Sub

Private Sub cmbMinRecLen_Click()
    f_DataChanged = True
End Sub

Private Sub cmbMinRecLen_GotFocus()
    cmbMinRecLen.SelStart = 0
    cmbMinRecLen.SelLength = Len(cmbMinRecLen)
End Sub

Private Sub cmbMinRecLen_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub cmbNotifyInterval_Change()
    f_DataChanged = True
End Sub

Private Sub cmbNotifyInterval_Click()
    f_DataChanged = True
End Sub

Private Sub cmbNotifyInterval_GotFocus()
    cmbNotifyInterval.SelStart = 0
    cmbNotifyInterval.SelLength = Len(cmbNotifyInterval)
End Sub

Private Sub cmbNotifyInterval_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub cmbVMSClass_Click()
    f_DataChanged = True
End Sub

'Mike added this event @ 2008-1-31
Private Sub cmdNodeTag_Click()
    frmNodeTagEdit.iNodeID = CInt(T_n_id)
    frmNodeTagEdit.byNodeNo = CByte(T_n_no.Text)
    frmNodeTagEdit.Show vbModal
End Sub

Private Sub cmdShowNodeList_Click(Index As Integer)
    
    Select Case Index
    Case 0
        Set gSystem.crlCurItem = T_nd_parent
    Case 1
        Set gSystem.crlCurItem = T_nd_child
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click()

    gSystem.intCurStep = 1
    Set gSystem.crlCurItem = T_vox_play
    frmResourceList.Show vbModal

End Sub

Public Sub CommandSave_Click()
'On Error Resume Next

    If f_DataChanged Then
        
        Dim lv_loop As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 录音时间长
        If Trim(Cb_timeout) = "" Then
            Node40_Data1.rectime = 0
        Else
            '' Michael Modify this for expansion the Recordtime length @ 7-3-07
            ''If Val(Cb_timeout) > 255 Then
            If Val(Cb_timeout) > 65535 Then
                Message ("E084")
                Cb_timeout.SetFocus
                Exit Sub
            ''Else
            Else
                Node40_Data1.rectime = CByte(Val(Cb_timeout.Text) Mod 256)
                Node40_Data2.rectime_ho = CByte(Val(Cb_timeout.Text) \ 256)
            End If
        End If

        '' Sun added 2007-03-20
        ' 最短录音时间长(秒)
        If Trim(cmbMinRecLen) = "" Then
            Node40_Data1.MinRecLength = 0
        Else
            If Val(cmbMinRecLen) > 180 Then
                Message ("E109")
                cmbMinRecLen.SetFocus
                Exit Sub
            Else
                Node40_Data1.MinRecLength = CByte(Val(cmbMinRecLen.Text) Mod 256)
            End If
        End If

        ' 最大静音时长(秒)
        If Trim(cmbMaxSilence) = "" Then
            Node40_Data1.maxsilencetime = 0
        Else
            If Val(cmbMaxSilence) > 60 Then
                Message ("E108")
                cmbMaxSilence.SetFocus
                Exit Sub
            Else
                Node40_Data1.maxsilencetime = CByte(Val(cmbMaxSilence.Text) Mod 256)
            End If
        End If
        
        ' 放音清空
        If Cb_playclear.ListIndex = -1 Then
            Node40_Data1.playclear = 0
        Else
            Node40_Data1.playclear = CByte(Cb_playclear.ItemData(Cb_playclear.ListIndex))
        End If

        ' 按键中断
        If Cb_breakkey.ListIndex < 0 Then
            Node40_Data1.breakkey = 0
        Else
            Node40_Data1.breakkey = CByte(Cb_breakkey.ItemData(Cb_breakkey.ListIndex))
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node40_Data1.log = 0
        Else
            Node40_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node40_Data1.var_agent = 0
        Else
            Node40_Data1.var_agent = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        ' Sun added 2007-02-28
        ' 留言分类
        If cmbVMSClass.ListIndex <= 0 Then
            Node40_Data1.vmsclass = 0
        Else
            Node40_Data1.vmsclass = CByte(cmbVMSClass.ItemData(cmbVMSClass.ListIndex))
        End If
        
        ' 变量：记录文件名
        If Cb_VarFileName.ListIndex <= 0 Then
            Node40_Data1.var_filename = 0
        Else
            Node40_Data1.var_filename = CByte(Cb_VarFileName.ItemData(Cb_VarFileName.ListIndex))
        End If
        
        ' 变量：附加字段
        For lv_loop = 0 To 2
            If Cb_VarAppField(lv_loop).ListIndex <= 0 Then
                Node40_Data1.var_appfield(lv_loop) = 0
            Else
                Node40_Data1.var_appfield(lv_loop) = CByte(Cb_VarAppField(lv_loop).ItemData(Cb_VarAppField(lv_loop).ListIndex))
            End If
        Next
        
        ' 录音提示
        If Trim(T_vox_play) = "" Then
            Node40_Data2.vox_op = 0
        Else
            If CLng(Trim(T_vox_play)) > 32767 Then
                Message ("E092")
                T_vox_play.SetFocus
                Exit Sub
            Else
                Node40_Data2.vox_op = CLng(Val(T_vox_play.Text))
            End If
        End If
    
        '-------------------------------------------------------
        '' Sun added 2009-07-24
        ' 留言开始提示音
        Node40_Data1.toneoff = (chkToneOn.value = vbUnchecked)
        
        ' 留言进行提示间隔(秒)
        If Trim(cmbNotifyInterval) = "" Then
            Node40_Data2.var_notifyintvl = 0
        Else
            If Val(Cb_timeout.Text) > 0 And Val(cmbNotifyInterval.Text) > Val(Cb_timeout.Text) Then
                Message ("E145")
                cmbNotifyInterval.SetFocus
                Exit Sub
            Else
                Node40_Data2.var_notifyintvl = CByte(Val(cmbNotifyInterval.Text) Mod 256)
            End If
        End If
        
        ' 变量：记录留言时长
        If Cb_VarRecDuration.ListIndex <= 0 Then
            Node40_Data2.var_rectime = 0
        Else
            Node40_Data2.var_rectime = CByte(Cb_VarRecDuration.ItemData(Cb_VarRecDuration.ListIndex))
        End If
        '-------------------------------------------------------
        
        '-------------------------------------------------------
        ' Sun added 2012-04-18
        Node40_Data2.NotifyPL = (chkNotifyPL.value = vbChecked)
        '-------------------------------------------------------
        
        Node40_Data2.reserved1(0) = 0
        Node40_Data2.reserved2(0) = 0
        
        '父节点
        If Trim(T_nd_parent) = "" Then
            Node40_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node40_Data2.nd_parent = CInt(Val(T_nd_parent.Text))
            End If
        End If

        '子节点
        If Trim(T_nd_child) = "" Then
            Node40_Data2.nd_child = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                Node40_Data2.nd_child = CInt(Val(T_nd_child.Text))
            End If
        End If
        
        'Michael Added @ 06-29-07
        '留言文件类型
        If cboVMSType.Text = LoadNationalResString(1802) Then Node40_Data2.recfiletype = 0
        If cboVMSType.Text = LoadNationalResString(1803) Then Node40_Data2.recfiletype = 1
        'Add End
    
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
   
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no.Text)
   
        f_DataChanged = False
        
    End If
    
    Unload Me

End Sub

Private Sub CommandExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
Dim lv_nIndex As Integer
Dim lv_strRes As String

    ' 放音清空标志
    With Cb_playclear
        .AddItem LoadNationalResString(1241)
        .ItemData(.ListCount - 1) = 0
        .AddItem LoadNationalResString(1242)
        .ItemData(.ListCount - 1) = 1
    End With
   
    ' 被访日志
    With Cb_log
        .AddItem LoadNationalResString(1178)
        .ItemData(.ListCount - 1) = 0
        For i = 1 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1179)
            .ItemData(.ListCount - 1) = i
        Next
    End With

    ' 输入终止符
    F_FillPhoneKeyList Cb_breakkey, 1

    ' 录音时间长(秒)
    With Cb_timeout
        For i = 0 To 200 Step 10
            .AddItem Trim(Str(i))
        Next
    End With
        
    '' Sun added 2007-03-20
    ' 最短录音时间长(秒)
    With cmbMinRecLen
        For i = 0 To 30 Step 5
            .AddItem Trim(Str(i))
        Next
    End With

    ' 最大静音时长(秒)
    With cmbMaxSilence
        For i = 0 To 60 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
    '' Sun added 2009-07-24
    ' 留言进行提示间隔(秒)
    With cmbNotifyInterval
        For i = 6 To 30 Step 3
            .AddItem Trim(Str(i))
        Next
    End With
    
    ' Sun added 2007-02-28
    ' 留言分类
    Dim strArrCaptions() As String
    lv_strRes = LoadNationalResString(1691)
    strArrCaptions = Split(lv_strRes, ";")
    lv_nIndex = 0
    With cmbVMSClass
        For i = DEF_NODE040_VMSCLASS_UNKNOWN To DEF_NODE040_VMSCLASS_USER
            If lv_nIndex <= UBound(strArrCaptions) Then
                lv_strRes = Trim(strArrCaptions(lv_nIndex))
                If lv_strRes <> "" Then
                    .AddItem lv_strRes
                End If
            End If
            .ItemData(.ListCount - 1) = i
            lv_nIndex = lv_nIndex + 1
        Next
    End With
    
    'Michael Added @ 06-29-07
    '留言文件类型
    cboVMSType.Clear
    cboVMSType.AddItem LoadNationalResString(1802), 0
    cboVMSType.AddItem LoadNationalResString(1803), 1
    
    If Node40_Data2.recfiletype = 0 Then
        cboVMSType.ListIndex = 0
    ElseIf Node40_Data2.recfiletype = 1 Then
        cboVMSType.ListIndex = 1
    Else
        If gSystem.intRecFileType = 0 Then
            cboVMSType.ListIndex = 0
        ElseIf gSystem.intRecFileType = 1 Then
            cboVMSType.ListIndex = 1
        End If
    End If
    'Add End
    
    ' 使用变量ID
    RefreshVariablesList Cb_usevar
    RefreshVariablesList Cb_VarFileName
    RefreshVariablesList Cb_VarRecDuration          '' Sun added 2009-07-24
    RefreshVariablesList Cb_VarAppField(0)
    RefreshVariablesList Cb_VarAppField(1)
    RefreshVariablesList Cb_VarAppField(2)
    
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_play = Node40_Data2.vox_op
    T_nd_parent.Text = Node40_Data2.nd_parent
    T_nd_child.Text = Node40_Data2.nd_child
    
    ''Cb_timeout.Text = Node40_Data1.rectime
    ''Michael Modified for Expansion Record Length
    Cb_timeout.Text = CLng(Node40_Data2.rectime_ho) * 256 + CLng(Node40_Data1.rectime)
   
    cmbMaxSilence.Text = Node40_Data1.maxsilencetime
    Cb_playclear.ListIndex = SearchItemDataIndex(Cb_playclear, CLng(Node40_Data1.playclear), 0)
    Cb_breakkey.ListIndex = SearchItemDataIndex(Cb_breakkey, CLng(Node40_Data1.breakkey), 10)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node40_Data1.log), 0)
    
    ' Sun added 2007-03-20
    cmbMinRecLen.Text = Node40_Data1.MinRecLength
    
    ' Sun added 2007-02-28
    cmbVMSClass.ListIndex = SearchItemDataIndex(cmbVMSClass, CLng(Node40_Data1.vmsclass), 0)
    
    '' Sun added 2009-07-24
    ' 留言开始提示音
    chkToneOn.value = IIf(Node40_Data1.toneoff, vbUnchecked, vbChecked)
    
    '-------------------------------------------------------
    ' Sun added 2012-04-18
    chkNotifyPL.value = IIf(Node40_Data2.NotifyPL, vbChecked, vbUnchecked)
    '-------------------------------------------------------
    
    '' Sun added 2009-07-24
    ' 留言进行提示间隔(秒)
    cmbNotifyInterval = Node40_Data2.var_notifyintvl
    
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node40_Data1.var_agent), 0)
    Cb_VarFileName.ListIndex = SearchItemDataIndex(Cb_VarFileName, CLng(Node40_Data1.var_filename), 0)
    Cb_VarRecDuration.ListIndex = SearchItemDataIndex(Cb_VarRecDuration, CLng(Node40_Data2.var_rectime), 0)
    Cb_VarAppField(0).ListIndex = SearchItemDataIndex(Cb_VarAppField(0), CLng(Node40_Data1.var_appfield(0)), 0)
    Cb_VarAppField(1).ListIndex = SearchItemDataIndex(Cb_VarAppField(1), CLng(Node40_Data1.var_appfield(1)), 0)
    Cb_VarAppField(2).ListIndex = SearchItemDataIndex(Cb_VarAppField(2), CLng(Node40_Data1.var_appfield(2)), 0)
  
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If f_DataChanged Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub T_nd_child_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_child_GotFocus()
    T_nd_child.SelStart = 0
    T_nd_child.SelLength = Len(T_nd_child)
End Sub

Private Sub T_nd_child_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_nd_parent_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_parent_GotFocus()
    T_nd_parent.SelStart = 0
    T_nd_parent.SelLength = Len(T_nd_parent)
End Sub

Private Sub T_nd_parent_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_play_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_play

End Sub

Private Sub T_vox_play_GotFocus()
    T_vox_play.SelStart = 0
    T_vox_play.SelLength = Len(T_vox_play)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_play)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_play_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
