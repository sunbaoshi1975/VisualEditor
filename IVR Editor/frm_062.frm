VERSION 5.00
Begin VB.Form frm_062 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "电话会议"
   ClientHeight    =   4800
   ClientLeft      =   2415
   ClientTop       =   2670
   ClientWidth     =   9885
   Icon            =   "frm_062.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2126"
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   825
      Index           =   1
      Left            =   3810
      TabIndex        =   35
      Tag             =   "1104"
      Top             =   3360
      Width           =   6015
      Begin VB.TextBox Txt_Description 
         Height          =   495
         Left            =   150
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   3285
      Index           =   0
      Left            =   3780
      TabIndex        =   33
      Tag             =   "1164"
      Top             =   0
      Width           =   6045
      Begin VB.TextBox T_vagency 
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
         Left            =   4125
         MaxLength       =   24
         TabIndex        =   8
         Top             =   360
         Width           =   1785
      End
      Begin VB.TextBox txtPreDial 
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
         TabIndex        =   7
         Top             =   360
         Width           =   1005
      End
      Begin VB.TextBox txtSysError 
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
         Top             =   1140
         Width           =   1005
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   7
         Left            =   5580
         Picture         =   "frm_062.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "1146"
         Top             =   1170
         Width           =   315
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
         TabIndex        =   25
         Top             =   1500
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
         Top             =   780
         Width           =   1005
      End
      Begin VB.TextBox T_vox_noconf 
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
         Top             =   780
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
         TabIndex        =   11
         Top             =   1140
         Width           =   1005
      End
      Begin VB.TextBox T_vox_answer 
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
         Top             =   1860
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
         TabIndex        =   13
         Top             =   1500
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2580
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
         Top             =   1860
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
         Top             =   2580
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
         TabIndex        =   17
         Top             =   2220
         Width           =   1005
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
         Top             =   2220
         Width           =   1005
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2610
         Picture         =   "frm_062.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "1146"
         Top             =   810
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   6
         Left            =   5580
         Picture         =   "frm_062.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "1146"
         Top             =   810
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2610
         Picture         =   "frm_062.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "1146"
         Top             =   1170
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   2610
         Picture         =   "frm_062.frx":10D2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "1146"
         Top             =   1890
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2610
         Picture         =   "frm_062.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "1146"
         Top             =   1530
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   4
         Left            =   2610
         Picture         =   "frm_062.frx":12D6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "1146"
         Top             =   2250
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   5
         Left            =   2610
         Picture         =   "frm_062.frx":13D8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "1146"
         Top             =   2610
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   5580
         Picture         =   "frm_062.frx":14DA
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "1145"
         Top             =   1530
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   5580
         Picture         =   "frm_062.frx":1864
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "1145"
         Top             =   1890
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   5580
         Picture         =   "frm_062.frx":1BEE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "1145"
         Top             =   2250
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   5580
         Picture         =   "frm_062.frx":1F78
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "1145"
         Top             =   2610
         Width           =   315
      End
      Begin VB.Label lblDialNO 
         AutoSize        =   -1  'True
         Caption         =   "电话号码"
         Height          =   195
         Left            =   3150
         TabIndex        =   63
         Tag             =   "1343"
         Top             =   450
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "拨号前缀"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   62
         Tag             =   "1342"
         Top             =   450
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "系统错提示音"
         Height          =   195
         Index           =   13
         Left            =   3150
         TabIndex        =   61
         Tag             =   "2132"
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   180
         Index           =   12
         Left            =   3120
         TabIndex        =   60
         Tag             =   "1169"
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "操作提示语音"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   59
         Tag             =   "1266"
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "磋商失败提示音"
         Height          =   195
         Index           =   2
         Left            =   3150
         TabIndex        =   58
         Tag             =   "2130"
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "等待音乐"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   57
         Tag             =   "1308"
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "成功应答提示音"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   56
         Tag             =   "2131"
         Top             =   1935
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "磋商完成提示音"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   55
         Tag             =   "2128"
         Top             =   1590
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "会议成功提示音"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   54
         Tag             =   "2129"
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "会议失败转节点"
         Height          =   195
         Index           =   8
         Left            =   3150
         TabIndex        =   53
         Tag             =   "2134"
         Top             =   1935
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "会议成功转节点"
         Height          =   195
         Index           =   9
         Left            =   3150
         TabIndex        =   52
         Tag             =   "2133"
         Top             =   2670
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答提示语音"
         Height          =   180
         Index           =   10
         Left            =   180
         TabIndex        =   51
         Tag             =   "1314"
         Top             =   2310
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答转节点"
         Height          =   180
         Index           =   11
         Left            =   3150
         TabIndex        =   50
         Tag             =   "1313"
         Top             =   2310
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   9420
      Picture         =   "frm_062.frx":2302
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "1948"
      Top             =   4320
      Width           =   333
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   4185
      Left            =   60
      TabIndex        =   39
      Tag             =   "1136"
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox Cb_varwaitansto 
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
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1620
         Width           =   1965
      End
      Begin VB.ComboBox Cb_NoAnsTO 
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
         TabIndex        =   1
         Text            =   "No Answer Timeout"
         ToolTipText     =   "0 - 255 秒"
         Top             =   1230
         Width           =   1965
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2430
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
         Top             =   2850
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   3240
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
         TabIndex        =   3
         Top             =   2040
         Width           =   1965
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
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   360
         Width           =   525
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
         Top             =   810
         Width           =   1965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "应答超时变量"
         Height          =   195
         Index           =   7
         Left            =   345
         TabIndex        =   64
         Tag             =   "2140"
         Top             =   1695
         Width           =   1080
      End
      Begin VB.Label lblNoAnsTO 
         AutoSize        =   -1  'True
         Caption         =   "应答超时"
         Height          =   195
         Left            =   345
         TabIndex        =   49
         Tag             =   "2127"
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   48
         Tag             =   "1248"
         Top             =   2505
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Index           =   0
         Left            =   345
         TabIndex        =   47
         Tag             =   "1159"
         Top             =   3330
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "播放次数"
         Height          =   180
         Index           =   1
         Left            =   345
         TabIndex        =   46
         Tag             =   "1304"
         Top             =   2940
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "节点超时"
         Height          =   180
         Index           =   0
         Left            =   345
         TabIndex        =   45
         Tag             =   "1154"
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "转接方式"
         Height          =   180
         Left            =   345
         TabIndex        =   44
         Tag             =   "1302"
         Top             =   2130
         Width           =   720
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Tag             =   "1143"
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1530
         TabIndex        =   42
         Tag             =   "1137"
         Top             =   420
         Width           =   525
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   5220
      TabIndex        =   37
      Tag             =   "1144"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   36
      Tag             =   "1007"
      Top             =   4320
      Width           =   1035
   End
End
Attribute VB_Name = "frm_062"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_062.frm
'//用途：    创建新的节点
'//作者:     Tony
'//创建日期：2012/04/18
'//修改日期：
'//文件描述：建立电话会议
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_looptimes_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_NoAnsTO_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_NoAnsTO_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_NoAnsTO_GotFocus()
    Cb_NoAnsTO.SelStart = 0
    Cb_NoAnsTO.SelLength = Len(Cb_NoAnsTO)
End Sub

Private Sub Cb_NoAnsTO_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub CB_switchtype_Click()
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
    
    If Cb_usevar.ListIndex > 0 Then
        lblDialNO.Enabled = False
        T_vagency.Enabled = False
    Else
        lblDialNO.Enabled = True
        T_vagency.Enabled = True
    End If
    
End Sub

Private Sub Cb_varwaitansto_Click()
    f_DataChanged = True
    
    If Cb_varwaitansto.ListIndex > 0 Then
        lblNoAnsTO.Enabled = False
        Cb_NoAnsTO.Enabled = False
    Else
        lblNoAnsTO.Enabled = True
        Cb_NoAnsTO.Enabled = True
    End If
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
        Set gSystem.crlCurItem = T_nd_busy
    Case 2
        Set gSystem.crlCurItem = txtNdNoAnswer
    Case 3
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
        Set gSystem.crlCurItem = T_vox_wt
    Case 2
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_sw
    Case 3
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_answer
    Case 4
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = txtVoxNoAnswer
    Case 5
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_ok
    Case 6
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_noconf
    Case 7
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = txtSysError
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
            Node62_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node62_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If

        ' 应答超时
        If Trim(Cb_NoAnsTO) = "" Then
            Node62_Data1.waitansto = 0
        Else
            If Val(Cb_NoAnsTO) > 255 Then
                Message ("E051")
                Cb_NoAnsTO.SetFocus
                Exit Sub
            Else
                Node62_Data1.waitansto = CByte(Val(Cb_NoAnsTO.Text) Mod 256)
            End If
        End If
        
        ' 转接方式
        If CB_switchtype.ListIndex = -1 Then
            Node62_Data1.switchtype = 0
        Else
            Node62_Data1.switchtype = CByte(CB_switchtype.ItemData(CB_switchtype.ListIndex))
        End If
        
        ' Sun added 2014-01-29
        ' 应答超时变量
        If Cb_varwaitansto.ListIndex <= 0 Then
            Node62_Data1.var_waitansto = 0
        Else
            Node62_Data1.var_waitansto = CByte(Cb_varwaitansto.ItemData(Cb_varwaitansto.ListIndex))
        End If
        
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node62_Data1.usevar = 0
        Else
            Node62_Data1.usevar = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        ' 等待播放次数
        If Cb_looptimes.ListIndex < 0 Then
            Node62_Data1.looptimes = 0
        Else
            Node62_Data1.looptimes = CByte(Cb_looptimes.ItemData(Cb_looptimes.ListIndex))
        End If

        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node62_Data1.log = 0
        Else
            Node62_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        ' 拨号前缀
        Call StringToByteArray(txtPreDial, Node62_Data2.predial, txtPreDial.MaxLength)
        
        ' 电话号码
        Call StringToByteArray(T_vagency, Node62_Data2.DialNo, T_vagency.MaxLength)
        
        Node62_Data1.reserved1 = 0
        Node62_Data1.reserved2(0) = 0
        
        ' 操作提示语音
        If Trim(T_vox_op) = "" Then
            Node62_Data2.vox_op = 0
        Else
            If CLng(Trim(T_vox_op)) > 32767 Then
                Message ("E092")
                T_vox_op.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_op = CInt(Trim(T_vox_op.Text))
            End If
        End If
        
        ' 磋商完成提示音
        If Trim(T_vox_sw) = "" Then
            Node62_Data2.vox_sw = 0
        Else
            If CLng(Trim(T_vox_sw)) > 32767 Then
                Message ("E149")
                T_vox_sw.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_sw = CInt(Trim(T_vox_sw.Text))
            End If
        End If

        ' 等待循环语音
        If Trim(T_vox_wt) = "" Then
            Node62_Data2.vox_wt = 0
        Else
            If CLng(Trim(T_vox_wt)) > 32767 Then
                Message ("E097")
                T_vox_wt.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_wt = CInt(Trim(T_vox_wt.Text))
            End If
        End If
        
        ' 成功应答提示音
        If Trim(T_vox_answer) = "" Then
            Node62_Data2.vox_ans = 0
        Else
            If CLng(Trim(T_vox_answer)) > 32767 Then
                Message ("E149")
                T_vox_answer.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_ans = CInt(Trim(T_vox_answer.Text))
            End If
        End If
        
        ' 无应答提示语音
        If Trim(txtVoxNoAnswer) = "" Then
            Node62_Data2.vox_noans = 0
        Else
            If CLng(Trim(txtVoxNoAnswer)) > 32767 Then
                Message ("E120")
                txtVoxNoAnswer.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_noans = CInt(Trim(txtVoxNoAnswer))
            End If
        End If
        
        ' 会议成功提示音
        If Trim(T_vox_ok) = "" Then
            Node62_Data2.vox_ok = 0
        Else
            If CLng(Trim(T_vox_ok)) > 32767 Then
                Message ("E119")
                T_vox_ok.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_ok = CInt(Trim(T_vox_ok.Text))
            End If
        End If
        
        ' 磋商失败提示音
        If Trim(T_vox_noconf) = "" Then
            Node62_Data2.vox_noconf = 0
        Else
            If CLng(Trim(T_vox_noconf)) > 32767 Then
                Message ("E149")
                T_vox_noconf.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_noconf = CInt(Trim(T_vox_noconf.Text))
            End If
        End If
       
        ' 系统错提示音
        If Trim(txtSysError) = "" Then
            Node62_Data2.vox_syserror = 0
        Else
            If CLng(Trim(txtSysError)) > 32767 Then
                Message ("E149")
                txtSysError.SetFocus
                Exit Sub
            Else
                Node62_Data2.vox_syserror = CInt(Trim(txtSysError.Text))
            End If
        End If
        
        Node62_Data2.reserved1(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node62_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node62_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
    
        ' 会议失败转节点
        If Trim(T_nd_busy) = "" Then
            Node62_Data2.nd_failed = 0
        Else
            If (Val(T_nd_busy) > 32767 Or Val(T_nd_busy) < 256) And Val(T_nd_busy) <> 0 Then
                Message ("E148")
                T_nd_busy.SetFocus
                Exit Sub
            Else
                Node62_Data2.nd_failed = CInt(Trim(T_nd_busy.Text))
            End If
        End If
        
        ' 无应答转节点
        If Trim(txtNdNoAnswer) = "" Then
            Node62_Data2.nd_noans = 0
        Else
            If (Val(txtNdNoAnswer) > 32767 Or Val(txtNdNoAnswer) < 256) And Val(txtNdNoAnswer) <> 0 Then
                Message ("E148")
                txtNdNoAnswer.SetFocus
                Exit Sub
            Else
                Node62_Data2.nd_noans = CInt(Trim(txtNdNoAnswer))
            End If
        End If
        
        ' 成功转接点
        If Trim(T_nd_ok) = "" Then
            Node62_Data2.nd_ok = 0
        Else
            If (Val(T_nd_ok) > 32767 Or Val(T_nd_ok) < 256) And Val(T_nd_ok) <> 0 Then
                Message ("E113")
                T_nd_ok.SetFocus
                Exit Sub
            Else
                Node62_Data2.nd_ok = CInt(Trim(T_nd_ok.Text))
            End If
        End If
        
        Node62_Data2.reserved2(0) = 0
        
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
   
    ' 应答超时(秒)
    With Cb_NoAnsTO
        For i = 0 To 60 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
    ' Sun added 2014-01-29
    ' 应答超时变量
    RefreshVariablesList Cb_varwaitansto
    
    ' 转接方式
    With CB_switchtype
        .Clear
        .AddItem LoadNationalResString(1335)
        .ItemData(.ListCount - 1) = 0
        .AddItem LoadNationalResString(1336)
        .ItemData(.ListCount - 1) = 1
    End With
    
    ' 使用变量ID
    RefreshVariablesList Cb_usevar
    
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
   
    ' 拨号前缀
    txtPreDial = ByteArrayToString(Node62_Data2.predial, txtPreDial.MaxLength)
   
    ' 电话号码
    T_vagency = ByteArrayToString(Node62_Data2.DialNo, T_vagency.MaxLength)
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node62_Data1.usevar), 0)
    Cb_usevar_Click
   
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_op.Text = Node62_Data2.vox_op
    T_vox_sw.Text = Node62_Data2.vox_sw
    T_vox_wt.Text = Node62_Data2.vox_wt
    txtVoxNoAnswer.Text = Node62_Data2.vox_noans
    T_vox_answer.Text = Node62_Data2.vox_ans
    T_vox_noconf.Text = Node62_Data2.vox_noconf
    txtSysError.Text = Node62_Data2.vox_syserror
    T_vox_ok.Text = Node62_Data2.vox_ok
    
    T_nd_parent.Text = Node62_Data2.nd_parent
    txtNdNoAnswer = Node62_Data2.nd_noans
    T_nd_busy.Text = Node62_Data2.nd_failed
    T_nd_ok.Text = Node62_Data2.nd_ok
    
    Cb_timeout.Text = Node62_Data1.timeout
    Cb_NoAnsTO.Text = Node62_Data1.waitansto
    CB_switchtype.ListIndex = SearchItemDataIndex(CB_switchtype, CLng(Node62_Data1.switchtype), 0)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node62_Data1.log), 0)
    Cb_looptimes.ListIndex = SearchItemDataIndex(Cb_looptimes, CLng(Node62_Data1.looptimes), 3)
    

    ' Sun added 2014-01-29
    ' 应答超时变量
    Cb_varwaitansto.ListIndex = SearchItemDataIndex(Cb_varwaitansto, CLng(Node62_Data1.var_waitansto), 0)
    Cb_varwaitansto_Click

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

Private Sub T_vagency_Change()
    f_DataChanged = True
End Sub

Private Sub T_vagency_GotFocus()
    T_vagency.SelStart = 0
    T_vagency.SelLength = Len(T_vagency)
End Sub

Private Sub T_vagency_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_answer_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_ok

End Sub

Private Sub T_vox_answer_GotFocus()
    T_vox_answer.SelStart = 0
    T_vox_answer.SelLength = Len(T_vox_answer)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_answer)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_answer_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_noconf_Change()
    f_DataChanged = True
    
    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_ok
    
End Sub

Private Sub T_vox_noconf_GotFocus()
    T_vox_noconf.SelStart = 0
    T_vox_noconf.SelLength = Len(T_vox_noconf)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_noconf)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_noconf_KeyPress(KeyAscii As Integer)
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

Private Sub txtPreDial_Change()
    f_DataChanged = True
End Sub

Private Sub txtSysError_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtVoxNoAnswer

End Sub

Private Sub txtSysError_GotFocus()
    txtSysError.SelStart = 0
    txtSysError.SelLength = Len(txtSysError)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(txtSysError)
    Call SoundResourceIDChanged

End Sub

Private Sub txtSysError_KeyPress(KeyAscii As Integer)
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

