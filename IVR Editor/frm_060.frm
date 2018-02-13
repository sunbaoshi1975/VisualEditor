VERSION 5.00
Begin VB.Form frm_060 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "转接座席"
   ClientHeight    =   4230
   ClientLeft      =   2415
   ClientTop       =   2670
   ClientWidth     =   9840
   Icon            =   "frm_060.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1301"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6795
      Picture         =   "frm_060.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "1948"
      Top             =   3801
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1065
      Index           =   1
      Left            =   3720
      TabIndex        =   60
      Tag             =   "1104"
      Top             =   2640
      Width           =   6045
      Begin VB.TextBox Txt_Description 
         Height          =   735
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   240
         Width           =   5865
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   2535
      Index           =   0
      Left            =   3720
      TabIndex        =   45
      Tag             =   "1164"
      Top             =   60
      Width           =   6045
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   4
         Left            =   5580
         Picture         =   "frm_060.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "1145"
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   5580
         Picture         =   "frm_060.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "1145"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   5580
         Picture         =   "frm_060.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "1145"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   5580
         Picture         =   "frm_060.frx":29DA
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "1145"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   5580
         Picture         =   "frm_060.frx":2D64
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "1145"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   6
         Left            =   5580
         Picture         =   "frm_060.frx":30EE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   5
         Left            =   2610
         Picture         =   "frm_060.frx":31F0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "1146"
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   4
         Left            =   2610
         Picture         =   "frm_060.frx":32F2
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "1146"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   2610
         Picture         =   "frm_060.frx":33F4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "1146"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2610
         Picture         =   "frm_060.frx":34F6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "1146"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2610
         Picture         =   "frm_060.frx":35F8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "1146"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2610
         Picture         =   "frm_060.frx":36FA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtNdNoAnswer 
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
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   29
         Top             =   1680
         Width           =   1005
      End
      Begin VB.TextBox txtVoxNoAnswer 
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
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox T_nd_ok 
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
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox T_nd_busy 
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
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   27
         Top             =   1320
         Width           =   1005
      End
      Begin VB.TextBox T_nd_nobody 
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
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   25
         Top             =   960
         Width           =   1005
      End
      Begin VB.TextBox T_vox_ok 
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
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   21
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox T_vox_sw 
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
         TabIndex        =   17
         Top             =   1680
         Width           =   1005
      End
      Begin VB.TextBox T_vox_busy 
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
         TabIndex        =   15
         Top             =   1320
         Width           =   1005
      End
      Begin VB.TextBox T_vox_wt 
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
         TabIndex        =   13
         Top             =   960
         Width           =   1005
      End
      Begin VB.TextBox T_vox_nobody 
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
         TabIndex        =   11
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox T_vox_op 
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
         TabIndex        =   9
         Top             =   240
         Width           =   1005
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
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   23
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答转节点"
         Height          =   180
         Index           =   11
         Left            =   3150
         TabIndex        =   62
         Tag             =   "1313"
         Top             =   1764
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答提示语音"
         Height          =   180
         Index           =   10
         Left            =   180
         TabIndex        =   61
         Tag             =   "1314"
         Top             =   2130
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "转接成功转节点"
         Height          =   180
         Index           =   9
         Left            =   3150
         TabIndex        =   59
         Tag             =   "1315"
         Top             =   2130
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "座席忙转节点"
         Height          =   180
         Index           =   8
         Left            =   3150
         TabIndex        =   58
         Tag             =   "1311"
         Top             =   1395
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "没有上班转节点"
         Height          =   180
         Index           =   7
         Left            =   3150
         TabIndex        =   57
         Tag             =   "1309"
         Top             =   1035
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "转接成功提示"
         Height          =   180
         Index           =   6
         Left            =   3150
         TabIndex        =   56
         Tag             =   "1306"
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "转接提示语音"
         Height          =   180
         Index           =   5
         Left            =   180
         TabIndex        =   55
         Tag             =   "1312"
         Top             =   1764
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "座席忙提示音"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   54
         Tag             =   "1310"
         Top             =   1398
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "等待音乐"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   53
         Tag             =   "1308"
         Top             =   1032
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "没有上班提示音"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   52
         Tag             =   "1307"
         Top             =   666
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "操作提示语音"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   51
         Tag             =   "1266"
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   180
         Index           =   0
         Left            =   3150
         TabIndex        =   50
         Tag             =   "1169"
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3645
      Left            =   60
      TabIndex        =   37
      Tag             =   "1136"
      Top             =   60
      Width           =   3585
      Begin VB.ComboBox Cb_agtInfoLen 
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
         Left            =   2325
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "0-不限制；1-10位"
         Top             =   3240
         Width           =   1125
      End
      Begin VB.CheckBox chkReadAgentInfo 
         Caption         =   "宣读座席编号"
         Height          =   285
         Left            =   80
         TabIndex        =   7
         Tag             =   "1305"
         Top             =   3315
         Width           =   1470
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2850
         Width           =   1965
      End
      Begin VB.ComboBox Cb_looptimes 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2490
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
         Left            =   1470
         TabIndex        =   0
         Text            =   "Cb_timeout"
         ToolTipText     =   "0 - 255 秒"
         Top             =   690
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
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   47
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Cb_var_key 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2130
         Width           =   1965
      End
      Begin VB.ComboBox Cb_getlength 
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
         Left            =   1470
         TabIndex        =   3
         Text            =   "Cb_getlength"
         Top             =   1770
         Width           =   1965
      End
      Begin VB.ComboBox CB_agentid 
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
         Left            =   1470
         TabIndex        =   2
         Text            =   "CB_agentid"
         Top             =   1410
         Width           =   1965
      End
      Begin VB.ComboBox CB_switchtype 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "编号长度"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   63
         Tag             =   "1974"
         Top             =   3330
         Width           =   720
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1530
         TabIndex        =   49
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "转接方式"
         Height          =   180
         Left            =   345
         TabIndex        =   44
         Tag             =   "1302"
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblGetDigitLen 
         AutoSize        =   -1  'True
         Caption         =   "按键长度"
         Height          =   180
         Left            =   345
         TabIndex        =   43
         Tag             =   "1263"
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label lblKeyVar 
         AutoSize        =   -1  'True
         Caption         =   "按键记录"
         Height          =   180
         Left            =   345
         TabIndex        =   42
         Tag             =   "1265"
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "播放次数"
         Height          =   180
         Left            =   345
         TabIndex        =   41
         Tag             =   "1304"
         Top             =   2580
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Index           =   0
         Left            =   345
         TabIndex        =   40
         Tag             =   "1159"
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "节点超时"
         Height          =   180
         Left            =   345
         TabIndex        =   39
         Tag             =   "1154"
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblAgentID 
         AutoSize        =   -1  'True
         Caption         =   "指定座席"
         Height          =   180
         Left            =   345
         TabIndex        =   38
         Tag             =   "1303"
         Top             =   1500
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   3450
      TabIndex        =   34
      Tag             =   "1007"
      Top             =   3780
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   4680
      TabIndex        =   35
      Tag             =   "1144"
      Top             =   3780
      Width           =   1035
   End
End
Attribute VB_Name = "frm_060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_060.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：转接座席
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub CB_agentid_Change()
    f_DataChanged = True
End Sub

Private Sub CB_agentid_Click()
    f_DataChanged = True
End Sub

Private Sub CB_agentid_GotFocus()
    CB_agentid.SelStart = 0
    CB_agentid.SelLength = Len(CB_agentid)
End Sub

Private Sub Cb_agentid_KeyPress(KeyAscii As Integer)
    If CB_switchtype.ListIndex = 1 Then
        KeyPress KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Cb_getlength_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_getlength_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_getlength_GotFocus()
    Cb_getlength.SelStart = 0
    Cb_getlength.SelLength = Len(Cb_getlength)
End Sub

Private Sub Cb_getlength_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_looptimes_Click()
    f_DataChanged = True
End Sub

Private Sub CB_switchtype_Click()
Dim i As Integer

    f_DataChanged = True
    
    Select Case CB_switchtype.ListIndex
    Case Is <= 0
        '' 自动传接
        CB_switchtype.ListIndex = 0
        lblAgentID.Enabled = False
        CB_agentid.Enabled = False
        lblKeyVar.Enabled = False
        Cb_var_key.Enabled = False
        lblGetDigitLen.Enabled = False
        Cb_getlength.Enabled = False
        
    Case 1
        '' 指定座席
        lblAgentID.Caption = LoadNationalResString(1316)
        CB_agentid.Clear
        lblAgentID.Enabled = True
        CB_agentid.Enabled = True
        lblKeyVar.Enabled = False
        Cb_var_key.Enabled = False
        lblGetDigitLen.Enabled = False
        Cb_getlength.Enabled = False
    
    Case 2
        '' 用户输入
        lblAgentID.Caption = LoadNationalResString(1249)
        ''' 输入终止符
        F_FillPhoneKeyList CB_agentid, 1

        lblAgentID.Enabled = True
        CB_agentid.Enabled = True
        lblKeyVar.Enabled = True
        Cb_var_key.Enabled = True
        lblGetDigitLen.Enabled = True
        Cb_getlength.Enabled = True

    End Select
        
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

Private Sub Cb_var_key_Click()
    f_DataChanged = True
End Sub

Private Sub chkReadAgentInfo_Click()
    f_DataChanged = True
    'Mike Added @2008-5-27
    Call setAgtInfoLenStatus
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
        Set gSystem.crlCurItem = T_nd_nobody
    Case 2
        Set gSystem.crlCurItem = T_nd_busy
    Case 3
        Set gSystem.crlCurItem = txtNdNoAnswer
    Case 4
        Set gSystem.crlCurItem = T_nd_ok
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click(Index As Integer)

    Select Case Index
    Case 0
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_op
    Case 1
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_nobody
    Case 2
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_wt
    Case 3
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_busy
    Case 4
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_sw
    Case 5
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = txtVoxNoAnswer
    Case 6
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_ok
    End Select
    frmResourceList.Show vbModal

End Sub

Public Sub Command1_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node60_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node60_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If

        ' 转接方式
        If CB_switchtype.ListIndex = -1 Then
            Node60_Data1.switchtype = 0
        Else
            Node60_Data1.switchtype = CByte(CB_switchtype.ItemData(CB_switchtype.ListIndex))
        End If
        
        ' 指定坐席/按键中断
        ' 按键长度
        Select Case Node60_Data1.switchtype
        Case Is < 1
            Node60_Data1.agentid = 0
            Node60_Data1.getlength = 0
        Case 1      '' 指定坐席
            Node60_Data1.agentid = Val(CB_agentid.Text) Mod 256
            Node60_Data1.getlength = CByte(Val(CB_agentid.Text) / 256)
        Case 2      '' 按键中断
            Node60_Data1.agentid = Val(Trim(CB_agentid.Text))
            If CLng(Trim(Cb_getlength)) > 255 Then
                Message ("E081")
                Cb_getlength.SetFocus
                Exit Sub
            Else
                Node60_Data1.getlength = CByte(Val(Cb_getlength.Text) Mod 256)
            End If
        End Select
        
        ' 等待播放次数
        'Mike Modified @2008-5-28; Fixed can not save loop times
        'If Cb_looptimes.ListIndex = -1 Then
        If Cb_looptimes.ListIndex < 0 Then
            Node60_Data1.looptimes = 0
        Else
            Node60_Data1.looptimes = CByte(Cb_looptimes.ItemData(Cb_looptimes.ListIndex))
        End If

        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node60_Data1.log = 0
        Else
            Node60_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        ' 按键记录
        If Cb_var_key.ListIndex <= 0 Then
            Node60_Data1.var_key = 0
        Else
            Node60_Data1.var_key = CByte(Cb_var_key.ItemData(Cb_var_key.ListIndex))
        End If
        
        ' 是否宣读座席信息
        Node60_Data1.agentinfo = chkReadAgentInfo.value
        
        'Mike Added @2008-5-27
        ' 座席编号长度
        If CBool(chkReadAgentInfo.value) Then
            If Cb_agtInfoLen.ListIndex <= 0 Then
                Node60_Data2.length_agentinfo = 0
            Else
                Node60_Data2.length_agentinfo = CByte(Cb_agtInfoLen.ItemData(Cb_agtInfoLen.ListIndex))
            End If
        End If
        
        Node60_Data1.reserved1(0) = 0
        
        ' 操作提示语音
        If Trim(T_vox_op) = "" Then
            Node60_Data2.vox_op = 0
        Else
            If CLng(Trim(T_vox_op)) > 32767 Then
                Message ("E092")
                T_vox_op.SetFocus
                Exit Sub
            Else
                Node60_Data2.vox_op = CInt(Trim(T_vox_op.Text))
            End If
        End If
        
        ' 转接提示语音
        If Trim(T_vox_sw) = "" Then
            Node60_Data2.vox_sw = 0
        Else
            If CLng(Trim(T_vox_sw)) > 32767 Then
                Message ("E096")
                T_vox_sw.SetFocus
                Exit Sub
            Else
                Node60_Data2.vox_sw = CInt(Trim(T_vox_sw.Text))
            End If
        End If

        ' 等待循环语音
        If Trim(T_vox_wt) = "" Then
            Node60_Data2.vox_wt = 0
        Else
            If CLng(Trim(T_vox_wt)) > 32767 Then
                Message ("E097")
                T_vox_wt.SetFocus
                Exit Sub
            Else
                Node60_Data2.vox_wt = CInt(Trim(T_vox_wt.Text))
            End If
        End If
        
        ' 没上班提示语音
        If Trim(T_vox_nobody) = "" Then
            Node60_Data2.vox_nobody = 0
        Else
            If CLng(Trim(T_vox_nobody)) > 32767 Then
                Message ("E098")
                T_vox_nobody.SetFocus
                Exit Sub
            Else
                Node60_Data2.vox_nobody = CInt(Trim(T_vox_nobody.Text))
            End If
        End If
        
        ' 座席忙提示音
        If Trim(T_vox_busy) = "" Then
            Node60_Data2.vox_busy = 0
        Else
            If CLng(Trim(T_vox_busy)) > 32767 Then
                Message ("E099")
                T_vox_busy.SetFocus
                Exit Sub
            Else
                Node60_Data2.vox_busy = CInt(Trim(T_vox_busy.Text))
            End If
        End If
        
        ' 座席无应答提示语音
        If Trim(txtVoxNoAnswer) = "" Then
            Node60_Data2.vox_noanswer = 0
        Else
            If CLng(Trim(txtVoxNoAnswer)) > 32767 Then
                Message ("E120")
                txtVoxNoAnswer.SetFocus
                Exit Sub
            Else
                Node60_Data2.vox_noanswer = CInt(Trim(txtVoxNoAnswer))
            End If
        End If
        
        ' 成功转接提示音
        If Trim(T_vox_ok) = "" Then
            Node60_Data2.vox_ok = 0
        Else
            If CLng(Trim(T_vox_ok)) > 32767 Then
                Message ("E119")
                T_vox_ok.SetFocus
                Exit Sub
            Else
                Node60_Data2.vox_ok = CInt(Trim(T_vox_ok.Text))
            End If
        End If
        
        Node60_Data2.reserved1(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node60_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node60_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
    
        ' 没有上班转接节点
        If Trim(T_nd_nobody) = "" Then
            Node60_Data2.nd_nobody = 0
        Else
            If (Val(T_nd_nobody) > 32767 Or Val(T_nd_nobody) < 256) And Val(T_nd_nobody) <> 0 Then
                Message ("E111")
                T_nd_nobody.SetFocus
                Exit Sub
            Else
                Node60_Data2.nd_nobody = CInt(Trim(T_nd_nobody.Text))
            End If
        End If
        
        ' 坐席忙转接点
        If Trim(T_nd_busy) = "" Then
            Node60_Data2.nd_busy = 0
        Else
            If (Val(T_nd_busy) > 32767 Or Val(T_nd_busy) < 256) And Val(T_nd_busy) <> 0 Then
                Message ("E112")
                T_nd_busy.SetFocus
                Exit Sub
            Else
                Node60_Data2.nd_busy = CInt(Trim(T_nd_busy.Text))
            End If
        End If
        
        ' 座席无应答转节点
        If Trim(txtNdNoAnswer) = "" Then
            Node60_Data2.nd_noanswer = 0
        Else
            If (Val(txtNdNoAnswer) > 32767 Or Val(txtNdNoAnswer) < 256) And Val(txtNdNoAnswer) <> 0 Then
                Message ("E121")
                txtNdNoAnswer.SetFocus
                Exit Sub
            Else
                Node60_Data2.nd_noanswer = CInt(Trim(txtNdNoAnswer))
            End If
        End If
        
        ' 成功转接点
        If Trim(T_nd_ok) = "" Then
            Node60_Data2.nd_ok = 0
        Else
            If (Val(T_nd_ok) > 32767 Or Val(T_nd_ok) < 256) And Val(T_nd_ok) <> 0 Then
                Message ("E113")
                T_nd_ok.SetFocus
                Exit Sub
            Else
                Node60_Data2.nd_ok = CInt(Trim(T_nd_ok.Text))
            End If
        End If
        
        Node60_Data2.reserved2(0) = 0
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
   
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no)
        f_DataChanged = False
        
    End If
    
    Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If f_DataChanged Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer

     ' 节点超时(秒)
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
   
    ' 转接方式
    With CB_switchtype
        .AddItem LoadNationalResString(1317)
        .ItemData(.ListCount - 1) = 0
        .AddItem LoadNationalResString(1318)
        .ItemData(.ListCount - 1) = 1
        .AddItem LoadNationalResString(1319)
        .ItemData(.ListCount - 1) = 2
    End With
    
    ' 按键长度
    With Cb_getlength
        For i = 0 To 9
            .AddItem Trim(Str(i))
        Next
    End With
    
    ' 按键记录
    RefreshVariablesList Cb_var_key
    
    ' 播放次数
    With Cb_looptimes
        For i = 0 To 10
            .AddItem Trim(Str(i)) & LoadNationalResString(1180)
            .ItemData(.ListCount - 1) = i
        Next
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
    
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_op.Text = Node60_Data2.vox_op
    T_vox_sw.Text = Node60_Data2.vox_sw
    T_vox_wt.Text = Node60_Data2.vox_wt
    T_vox_nobody.Text = Node60_Data2.vox_nobody
    T_vox_busy.Text = Node60_Data2.vox_busy
    T_vox_ok.Text = Node60_Data2.vox_ok
    
    txtVoxNoAnswer = Node60_Data2.vox_noanswer
    txtNdNoAnswer = Node60_Data2.nd_noanswer
    chkReadAgentInfo.value = Node60_Data1.agentinfo
    'Mike Added @ 2008-5-27 # Set AgentInfoLength combobox
    Call setAgtInfoLenStatus
    
    T_nd_parent.Text = Node60_Data2.nd_parent
    T_nd_nobody.Text = Node60_Data2.nd_nobody
    T_nd_busy.Text = Node60_Data2.nd_busy
    T_nd_ok.Text = Node60_Data2.nd_ok
    
    Cb_timeout.Text = Node60_Data1.timeout
    CB_switchtype.ListIndex = SearchItemDataIndex(CB_switchtype, CLng(Node60_Data1.switchtype), 0)
    CB_switchtype_Click
    Select Case CB_switchtype.ListIndex
    Case 1
        CB_agentid.Text = Trim(Str(Node60_Data1.agentid + Node60_Data1.getlength * 256))
        Cb_getlength.Text = ""
    Case 2
        CB_agentid.Text = Node60_Data1.agentid
        Cb_getlength.Text = Node60_Data1.getlength
    Case Else
        CB_agentid.Text = ""
        Cb_getlength.Text = ""
    End Select
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node60_Data1.log), 0)
    Cb_looptimes.ListIndex = SearchItemDataIndex(Cb_looptimes, CLng(Node60_Data1.looptimes), 3)
    
    Cb_var_key.ListIndex = SearchItemDataIndex(Cb_var_key, CLng(Node60_Data1.var_key), 0)

    ' Data OK
    f_DataChanged = False
    LoadResStrings Me
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub T_nd_busy_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_busy_GotFocus()
    T_nd_busy.SelStart = 0
    T_nd_busy.SelLength = Len(T_nd_busy)
End Sub

Private Sub T_nd_busy_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_nd_nobody_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_nobody_GotFocus()
    T_nd_nobody.SelStart = 0
    T_nd_nobody.SelLength = Len(T_nd_nobody)
End Sub

Private Sub T_nd_nobody_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_nd_ok_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_ok_GotFocus()
    T_nd_ok.SelStart = 0
    T_nd_ok.SelLength = Len(T_nd_ok)
End Sub
Private Sub T_nd_ok_KeyPress(KeyAscii As Integer)
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

Private Sub T_vox_busy_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_busy

End Sub

Private Sub T_vox_busy_GotFocus()
    T_vox_busy.SelStart = 0
    T_vox_busy.SelLength = Len(T_vox_busy)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_busy)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_busy_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_nobody_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_nobody

End Sub

Private Sub T_vox_nobody_GotFocus()
    T_vox_nobody.SelStart = 0
    T_vox_nobody.SelLength = Len(T_vox_nobody)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_nobody)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_nobody_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_ok_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_ok

End Sub

Private Sub T_vox_ok_GotFocus()
    T_vox_ok.SelStart = 0
    T_vox_ok.SelLength = Len(T_vox_ok)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_ok)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_ok_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_op_Change()
    f_DataChanged = True
    
    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_op

End Sub

Private Sub T_vox_op_GotFocus()
    T_vox_op.SelStart = 0
    T_vox_op.SelLength = Len(T_vox_op)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_op)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_op_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_sw_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_sw

End Sub

Private Sub T_vox_sw_GotFocus()
    T_vox_sw.SelStart = 0
    T_vox_sw.SelLength = Len(T_vox_sw)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_sw)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_sw_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_wt_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_wt

End Sub

Private Sub T_vox_wt_GotFocus()
    T_vox_wt.SelStart = 0
    T_vox_wt.SelLength = Len(T_vox_wt)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_wt)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_wt_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtNdNoAnswer_Change()
    f_DataChanged = True
End Sub

Private Sub txtNdNoAnswer_GotFocus()
    txtNdNoAnswer.SelStart = 0
    txtNdNoAnswer.SelLength = Len(txtNdNoAnswer)
End Sub

Private Sub txtNdNoAnswer_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtVoxNoAnswer_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtVoxNoAnswer

End Sub

Private Sub txtVoxNoAnswer_GotFocus()
    txtVoxNoAnswer.SelStart = 0
    txtVoxNoAnswer.SelLength = Len(txtVoxNoAnswer)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(txtVoxNoAnswer)
    Call SoundResourceIDChanged

End Sub

Private Sub txtVoxNoAnswer_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

' Mike Added this function @ 2008-5-27,
' en/disable the combobox by the value of chkAgentInfo
Private Sub setAgtInfoLenStatus()
    Cb_agtInfoLen.Enabled = CBool(chkReadAgentInfo.value)
    ' fill the combo list
    Dim i As Byte
    With Cb_agtInfoLen
        .Clear
        .AddItem LoadNationalResString(1975)
        .ItemData(.ListCount - 1) = 0
        
        For i = 1 To 10
            .AddItem Trim(Str(i)) & LoadNationalResString(1973)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
    If Cb_agtInfoLen.Enabled Then
        Cb_agtInfoLen.ListIndex = SearchItemDataIndex(Cb_agtInfoLen, CLng(Node60_Data2.length_agentinfo), 4)
    Else
        Cb_agtInfoLen.ListIndex = 0
    End If
End Sub

Private Sub Cb_agtInfoLen_Click()
     f_DataChanged = True
End Sub
' Added End
