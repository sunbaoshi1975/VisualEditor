VERSION 5.00
Begin VB.Form frm_061 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "转接座席组  "
   ClientHeight    =   5880
   ClientLeft      =   2295
   ClientTop       =   2670
   ClientWidth     =   9975
   Icon            =   "frm_061.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1320"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   9552
      Picture         =   "frm_061.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "1948"
      Top             =   5181
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1755
      Index           =   1
      Left            =   3870
      TabIndex        =   65
      Tag             =   "1104"
      Top             =   3060
      Width           =   6015
      Begin VB.TextBox Txt_Description 
         Height          =   1455
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   210
         Width           =   5775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   2895
      Index           =   0
      Left            =   3870
      TabIndex        =   54
      Tag             =   "1164"
      Top             =   60
      Width           =   6015
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   5
         Left            =   5580
         Picture         =   "frm_061.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "1145"
         Top             =   2040
         Width           =   315
      End
      Begin VB.TextBox txtNdWait 
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
         TabIndex        =   40
         Top             =   2040
         Width           =   1005
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   4
         Left            =   5580
         Picture         =   "frm_061.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "1145"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   5580
         Picture         =   "frm_061.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "1145"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   5580
         Picture         =   "frm_061.frx":29DA
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "1145"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   5580
         Picture         =   "frm_061.frx":2D64
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "1145"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   5580
         Picture         =   "frm_061.frx":30EE
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   6
         Left            =   2610
         Picture         =   "frm_061.frx":3478
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "1146"
         Top             =   2430
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   5
         Left            =   2610
         Picture         =   "frm_061.frx":357A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "1146"
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   4
         Left            =   2610
         Picture         =   "frm_061.frx":367C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "1146"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   2610
         Picture         =   "frm_061.frx":377E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "1146"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2610
         Picture         =   "frm_061.frx":3880
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "1146"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2610
         Picture         =   "frm_061.frx":3982
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "1146"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2610
         Picture         =   "frm_061.frx":3A84
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   36
         Top             =   1320
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
         TabIndex        =   26
         Top             =   2040
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
         TabIndex        =   30
         Top             =   240
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
         TabIndex        =   16
         Top             =   240
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
         TabIndex        =   18
         Top             =   600
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
         TabIndex        =   20
         Top             =   960
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
         TabIndex        =   22
         Top             =   1320
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
         TabIndex        =   24
         Top             =   1680
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
         TabIndex        =   28
         Top             =   2400
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
         TabIndex        =   32
         Top             =   600
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
         TabIndex        =   34
         Top             =   960
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
         TabIndex        =   38
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "等待流程入口节点"
         Height          =   180
         Index           =   12
         Left            =   3030
         TabIndex        =   71
         Tag             =   "1811"
         Top             =   2130
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答转节点"
         Height          =   180
         Index           =   11
         Left            =   3030
         TabIndex        =   67
         Tag             =   "1313"
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答提示语音"
         Height          =   180
         Index           =   10
         Left            =   180
         TabIndex        =   66
         Tag             =   "1314"
         Top             =   2130
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   180
         Index           =   0
         Left            =   3030
         TabIndex        =   64
         Tag             =   "1169"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "操作提示语音"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   63
         Tag             =   "1266"
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "没有上班提示音"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   62
         Tag             =   "1307"
         Top             =   666
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "等待音乐"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   61
         Tag             =   "1308"
         Top             =   1032
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "座席忙提示音"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   60
         Tag             =   "1310"
         Top             =   1398
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "转接提示语音"
         Height          =   180
         Index           =   5
         Left            =   180
         TabIndex        =   59
         Tag             =   "1312"
         Top             =   1764
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "转接成功提示"
         Height          =   180
         Index           =   6
         Left            =   180
         TabIndex        =   58
         Tag             =   "1306"
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "没有上班转节点"
         Height          =   180
         Index           =   7
         Left            =   3030
         TabIndex        =   57
         Tag             =   "1309"
         Top             =   675
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "座席忙转节点"
         Height          =   180
         Index           =   8
         Left            =   3030
         TabIndex        =   56
         Tag             =   "1311"
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "转接成功转节点"
         Height          =   180
         Index           =   9
         Left            =   3030
         TabIndex        =   55
         Tag             =   "1315"
         Top             =   1770
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   5715
      Left            =   60
      TabIndex        =   48
      Tag             =   "1136"
      Top             =   60
      Width           =   3735
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
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "0-不限制；1-10位"
         Top             =   5280
         Width           =   1125
      End
      Begin VB.TextBox txtWaitansto 
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
         Left            =   1860
         TabIndex        =   1
         Top             =   1090
         Width           =   1695
      End
      Begin VB.TextBox T_groupid 
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
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2730
         Width           =   1695
      End
      Begin VB.ComboBox cboUserVarID 
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3105
         Width           =   1695
      End
      Begin VB.TextBox txtUserID 
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
         Left            =   1860
         TabIndex        =   12
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txtLoginID 
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
         Left            =   1860
         TabIndex        =   13
         Top             =   4920
         Width           =   1695
      End
      Begin VB.ComboBox cboReadEWT 
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CheckBox chkReadEWT 
         Caption         =   "播报等待时间"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Tag             =   "1814"
         Top             =   4260
         Width           =   1455
      End
      Begin VB.ComboBox cboSwitchMode 
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3465
         Width           =   1695
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1520
         Width           =   1695
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1950
         Width           =   1695
      End
      Begin VB.ComboBox cboWaitMode 
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3840
         Width           =   1695
      End
      Begin VB.OptionButton optTransferTo 
         Caption         =   "转RoutePoint"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Tag             =   "1536"
         Top             =   2400
         Width           =   1635
      End
      Begin VB.OptionButton optTransferTo 
         Caption         =   "转ACD"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Tag             =   "1535"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtWaitTime 
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
         Left            =   1860
         MaxLength       =   5
         TabIndex        =   0
         Top             =   660
         Width           =   1695
      End
      Begin VB.CheckBox chkReadAgentInfo 
         Caption         =   "宣读座席编号"
         Height          =   285
         Left            =   60
         TabIndex        =   14
         Tag             =   "1305"
         Top             =   5318
         Width           =   1425
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "编号长度"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   76
         Tag             =   "1974"
         Top             =   5363
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   75
         Top             =   3188
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "等待座席应答超时(秒)"
         Height          =   195
         Left            =   180
         TabIndex        =   74
         Tag             =   "1818"
         Top             =   1200
         Width           =   1710
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "转接座席工号变量"
         Height          =   195
         Left            =   180
         TabIndex        =   73
         Tag             =   "1813"
         Top             =   5010
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "转接座席ID记录变量"
         Height          =   195
         Left            =   180
         TabIndex        =   72
         Tag             =   "1812"
         Top             =   4650
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "等待方式"
         Height          =   180
         Left            =   180
         TabIndex        =   70
         Tag             =   "1806"
         Top             =   3930
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "转接方式"
         Height          =   180
         Left            =   180
         TabIndex        =   69
         Tag             =   "1805"
         Top             =   3570
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "等待超时(秒)"
         Height          =   195
         Left            =   180
         TabIndex        =   68
         Tag             =   "1534"
         Top             =   743
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   53
         Tag             =   "1159"
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "播放次数"
         Height          =   180
         Left            =   180
         TabIndex        =   52
         Tag             =   "1304"
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   51
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1920
         TabIndex        =   50
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACD DN"
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Tag             =   "1537"
         Top             =   2820
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   7200
      TabIndex        =   44
      Tag             =   "1144"
      Top             =   5160
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   43
      Tag             =   "1007"
      Top             =   5160
      Width           =   1035
   End
End
Attribute VB_Name = "frm_061"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_061.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：转接座席组
'//////////////////////////////////////////////////////////////////
'//修改时间 : 7-4-07
'//修改内容 : 添加4项新属性 (usevar, switchtype, waitmethod, nd_wait)
'//修改人   : Michael
'//修改版本 : V1.1
'/////////////////////////////////////////////////////////////////
'
'//修改时间 : July,9,07
'//修改内容 : SData1中添加3项属性(readEWT,var_userid,var_loginid)
'//修改人   : Michael
'//修改版本 : V1.2
'//////////////////////////////////////////////////////////////////
'//修改时间 : July,10,07
'//修改内容 : SData1中添加1项属性(waitansto)
'//修改人   : Michael
'//修改版本 : V1.2.1
'//////////////////////////////////////////////////////////////////
'//修改时间 : 2007-11-27
'//修改内容 : 修改用户变量列表
'//修改人   : Michael
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

'Michael Added @7-4-07
Private Sub cboSwitchMode_Click()
    f_DataChanged = True
End Sub

'Michael Added @7-4-07
Private Sub cboUserVarID_Click()
    f_DataChanged = True
End Sub

'Michael Added @7-4-07
Private Sub cboWaitMode_Click()
    f_DataChanged = True
End Sub

Private Sub chkReadAgentInfo_Click()
    'Mike Added @2008-5-27
    Call setAgtInfoLenStatus
    f_DataChanged = True
End Sub

'Michael Added @July,9,07
Private Sub chkReadEWT_Click()
    If chkReadEWT.value = 1 Then
        cboReadEWT.Enabled = True
        FillcboReadEWT
    Else
        cboReadEWT.Enabled = False
    End If
    
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
        Set gSystem.crlCurItem = T_nd_nobody
    Case 2
        Set gSystem.crlCurItem = T_nd_busy
    Case 3
        Set gSystem.crlCurItem = txtNdNoAnswer
    Case 4
        Set gSystem.crlCurItem = T_nd_ok
    ' Michael add @ 7-5-07
    Case 5
        Set gSystem.crlCurItem = txtNdWait
    ' Add End
    
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

        Node61_Data1.maxwait = CInt(txtWaitTime)
        
        'Michael Added @Jul,10,07
        '等待坐席应答超时
        Node61_Data1.waitansto = CByte(txtWaitansto Mod 256)
        Node61_Data2.waitansto_hi = CByte(txtWaitansto \ 256)
        
        ' 坐席组ID
        Node61_Data1.toacd = optTransferTo(0).value
        If Node61_Data1.toacd > 0 Then
            Call StringToByteArray(T_groupid, Node61_Data2.acddn, 20)
        Else
            If Len(Trim(T_groupid)) = 0 Then
               Node61_Data2.routepointid = 0
            Else
               If CLng(Trim(T_groupid)) > 32767 Then
                  Message ("E085")
                  T_groupid.SetFocus
                  Exit Sub
               Else
                  Node61_Data2.routepointid = CInt(Trim(T_groupid.Text))
                End If
            End If
        End If
        
        ' 保留
        'Michael commoned @ Jul,10,07
        'Node61_Data1.reserved1 = 0
        
        ' 等待播放次数
        If Cb_looptimes.ListIndex < 0 Then
            Node61_Data1.looptimes = 0
        Else
            Node61_Data1.looptimes = CByte(Cb_looptimes.ItemData(Cb_looptimes.ListIndex))
        End If

        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node61_Data1.log = 0
        Else
            Node61_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
                
        ' 是否宣读座席信息
        Node61_Data1.agentinfo = chkReadAgentInfo.value
        
        'Mike Added @2008-5-27
        ' 座席编号长度
        If CBool(chkReadAgentInfo.value) Then
            If Cb_agtInfoLen.ListIndex <= 0 Then
                Node61_Data2.length_agentinfo = 0
            Else
                Node61_Data2.length_agentinfo = CByte(Cb_agtInfoLen.ItemData(Cb_agtInfoLen.ListIndex))
            End If
        End If
        
        ' 保留
        'Michael Modify @ July,9,07
        'Node61_Data1.reserved2(0) = 0
        
        ' 操作提示语音
        If Trim(T_vox_op) = "" Then
            Node61_Data2.vox_op = 0
        Else
            If CLng(Trim(T_vox_op)) > 32767 Then
                Message ("E092")
                T_vox_op.SetFocus
                Exit Sub
            Else
                Node61_Data2.vox_op = CInt(Trim(T_vox_op.Text))
            End If
        End If
        
        ' 转接提示语音
        If Trim(T_vox_sw) = "" Then
            Node61_Data2.vox_sw = 0
        Else
            If CLng(Trim(T_vox_sw)) > 32767 Then
                Message ("E096")
                T_vox_sw.SetFocus
                Exit Sub
            Else
                Node61_Data2.vox_sw = CInt(Trim(T_vox_sw.Text))
            End If
        End If

        ' 等待循环语音
        If Trim(T_vox_wt) = "" Then
            Node61_Data2.vox_wt = 0
        Else
            If CLng(Trim(T_vox_wt)) > 32767 Then
                Message ("E097")
                T_vox_wt.SetFocus
                Exit Sub
            Else
                Node61_Data2.vox_wt = CInt(Trim(T_vox_wt.Text))
            End If
        End If
        
        ' 没上班提示语音
        If Trim(T_vox_nobody) = "" Then
            Node61_Data2.vox_nobody = 0
        Else
            If CLng(Trim(T_vox_nobody)) > 32767 Then
                Message ("E098")
                T_vox_nobody.SetFocus
                Exit Sub
            Else
                Node61_Data2.vox_nobody = CInt(Trim(T_vox_nobody.Text))
            End If
        End If
        
        ' 座席忙提示音
        If Trim(T_vox_busy) = "" Then
            Node61_Data2.vox_busy = 0
        Else
            If CLng(Trim(T_vox_busy)) > 32767 Then
                Message ("E099")
                T_vox_busy.SetFocus
                Exit Sub
            Else
                Node61_Data2.vox_busy = CInt(Trim(T_vox_busy.Text))
            End If
        End If
        
        ' 座席无应答提示语音
        If Trim(txtVoxNoAnswer) = "" Then
            Node61_Data2.vox_noanswer = 0
        Else
            If CLng(Trim(txtVoxNoAnswer)) > 32767 Then
                Message ("E120")
                txtVoxNoAnswer.SetFocus
                Exit Sub
            Else
                Node61_Data2.vox_noanswer = CInt(Trim(txtVoxNoAnswer))
            End If
        End If
        
        ' 成功转接提示音
        If Trim(T_vox_ok) = "" Then
            Node61_Data2.vox_ok = 0
        Else
            If CLng(Trim(T_vox_ok)) > 32767 Then
                Message ("E119")
                T_vox_ok.SetFocus
                Exit Sub
            Else
                Node61_Data2.vox_ok = CInt(Trim(T_vox_ok.Text))
            End If
        End If
        
        ' 保留
        Node61_Data2.reserved1(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node61_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node61_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
    
        ' 没有上班转接节点
        If Trim(T_nd_nobody) = "" Then
            Node61_Data2.nd_nobody = 0
        Else
            If (Val(T_nd_nobody) > 32767 Or Val(T_nd_nobody) < 256) And Val(T_nd_nobody) <> 0 Then
                Message ("E111")
                T_nd_nobody.SetFocus
                Exit Sub
            Else
                Node61_Data2.nd_nobody = CInt(Trim(T_nd_nobody.Text))
            End If
        End If
        
        ' 坐席忙转接点
        If Trim(T_nd_busy) = "" Then
            Node61_Data2.nd_busy = 0
        Else
            If (Val(T_nd_busy) > 32767 Or Val(T_nd_busy) < 256) And Val(T_nd_busy) <> 0 Then
                Message ("E112")
                T_nd_busy.SetFocus
                Exit Sub
            Else
                Node61_Data2.nd_busy = CInt(Trim(T_nd_busy.Text))
            End If
        End If
        
        ' 座席无应答转节点
        If Trim(txtNdNoAnswer) = "" Then
            Node61_Data2.nd_noanswer = 0
        Else
            If (Val(txtNdNoAnswer) > 32767 Or Val(txtNdNoAnswer) < 256) And Val(txtNdNoAnswer) <> 0 Then
                Message ("E121")
                txtNdNoAnswer.SetFocus
                Exit Sub
            Else
                Node61_Data2.nd_noanswer = CInt(Trim(txtNdNoAnswer))
            End If
        End If
        
        ' 成功转接点
        If Trim(T_nd_ok) = "" Then
            Node61_Data2.nd_ok = 0
        Else
            If (Val(T_nd_ok) > 32767 Or Val(T_nd_ok) < 256) And Val(T_nd_ok) <> 0 Then
                Message ("E113")
                T_nd_ok.SetFocus
                Exit Sub
            Else
                Node61_Data2.nd_ok = CInt(Trim(T_nd_ok.Text))
            End If
        End If
        
        '转接座席ID记录变量
        'Michael Added @ July,9,07
        If Trim(txtUserID) = "" Then
            Node61_Data1.var_userid = 0
        Else
            If (Val(txtUserID) > 255 Or Val(txtUserID) < 1) And Val(txtUserID) <> 0 Then
                Message ("E132")
                txtUserID.SetFocus
                Exit Sub
            Else
                Node61_Data1.var_userid = CByte(Val(txtUserID.Text) Mod 256)
            End If
        End If
        
        '转接座席工号记录变量
        'Michael Added @ July,9,07
        If Trim(txtLoginID) = "" Then
            Node61_Data1.var_loginid = 0
        Else
            If (Val(txtLoginID) > 255 Or Val(txtLoginID) < 1) And Val(txtLoginID) <> 0 Then
                Message ("E133")
                txtLoginID.SetFocus
                Exit Sub
            Else
                Node61_Data1.var_loginid = CByte(Val(txtLoginID.Text) Mod 256)
            End If
        End If
        
        
        '等待流程入口节点ID
        'Michael Add @ 7-5-07
        If Trim(txtNdWait) = "" Then
            Node61_Data2.nd_wait = 0
        Else
            If (Val(txtNdWait) > 32767 Or Val(txtNdWait) < 256) And Val(txtNdWait) <> 0 Then
                Message ("E131")
                txtNdWait.SetFocus
                Exit Sub
            Else
                Node61_Data2.nd_wait = CInt(Trim(txtNdWait.Text))
            End If
        End If
        
        '播报预测等待时间选项
        'Michael Added @ July,9,07
        If chkReadEWT.value = 0 Then
            Node61_Data1.readEWT = 0
        Else
            Node61_Data1.readEWT = CByte(cboReadEWT.ListIndex)
        End If
        'End Add
        
        ' 保留
        Node61_Data2.reserved2(0) = 0
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        'Michael Added @7-4-07
        'Wait Mode
        Node61_Data1.waitmethod = CByte(cboWaitMode.ListIndex)
        
        'Switch Mode
        Node61_Data1.switchtype = CByte(cboSwitchMode.ListIndex)
        
        'User Var ID
        'Modified @ 2007-11-27
        If cboUserVarID.ListIndex <= 0 Then
            Node61_Data1.usevar = 0
        Else
            Node61_Data1.usevar = CByte(cboUserVarID.ItemData(cboUserVarID.ListIndex))
        End If

        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no.Text)
   
        f_DataChanged = False
        
    End If
    
    Unload Me

End Sub

Private Sub Command2_Click()
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
    T_vox_op.Text = Node61_Data2.vox_op
    T_vox_sw.Text = Node61_Data2.vox_sw
    T_vox_wt.Text = Node61_Data2.vox_wt
    T_vox_nobody.Text = Node61_Data2.vox_nobody
    T_vox_busy.Text = Node61_Data2.vox_busy
    T_vox_ok.Text = Node61_Data2.vox_ok
    
    txtVoxNoAnswer = Node61_Data2.vox_noanswer
    txtNdNoAnswer = Node61_Data2.nd_noanswer
    chkReadAgentInfo.value = Node61_Data1.agentinfo
    'Mike Added @ 2008-5-27 # Set AgentInfoLength combobox
    Call setAgtInfoLenStatus
    
    T_nd_parent.Text = Node61_Data2.nd_parent
    T_nd_nobody.Text = Node61_Data2.nd_nobody
    T_nd_busy.Text = Node61_Data2.nd_busy
    T_nd_ok.Text = Node61_Data2.nd_ok
    'Michael Add @7-5-07
    txtNdWait.Text = Node61_Data2.nd_wait
    
    'Michael Added @ July,9,07
    txtUserID.Text = Node61_Data1.var_userid
    txtLoginID.Text = Node61_Data1.var_loginid
    
    'Michael Added @7-4-07
    With cboWaitMode
        .AddItem LoadNationalResString(1809), 0
        .AddItem LoadNationalResString(1810), 1
    End With
    
    With cboSwitchMode
        .AddItem LoadNationalResString(1807), 0
        .AddItem LoadNationalResString(1808), 1
    End With
    '-----  Add End ---------
    
    ' 等待超时（秒）
    txtWaitTime = Node61_Data1.maxwait
    
    'Michael Added @ Jul,10,07
    txtWaitansto = CLng(Node61_Data2.waitansto_hi) * 256 + Node61_Data1.waitansto
    
    ' 转接方式
    If Node61_Data1.toacd = 0 Then
        optTransferTo(1).value = True
        T_groupid.Text = Node61_Data2.routepointid
    Else
        optTransferTo(0).value = True
        T_groupid.Text = ByteArrayToString(Node61_Data2.acddn, 20)
    End If
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node61_Data1.log), 0)
    Cb_looptimes.ListIndex = SearchItemDataIndex(Cb_looptimes, CLng(Node61_Data1.looptimes), 3)
    
    'Michael Added @ 7-4-07
    'User Var ID
    'Michael Modified @ 2007-11-27
    Call RefreshVariablesList(cboUserVarID)
    cboUserVarID.ListIndex = SearchItemDataIndex(cboUserVarID, CLng(Node61_Data1.usevar), 0)
    
    'Michael Added @7-4-07
    'Switch Mode
    If Node61_Data1.switchtype = 0 Then
        cboSwitchMode.ListIndex = 0
    Else
        cboSwitchMode.ListIndex = 1
    End If
    
    'Michael Added @ 7-4-07
    'Wait Mode
    If Node61_Data1.waitmethod = 0 Then
        cboWaitMode.ListIndex = 0
    Else
        cboWaitMode.ListIndex = 1
    End If
    
    'Michael Added @ July,9,07
    If Node61_Data1.readEWT = 0 Then
        chkReadEWT.value = 0
        cboReadEWT.Enabled = False
    Else
        chkReadEWT.value = 1
        cboReadEWT.Enabled = True
        FillcboReadEWT
        cboReadEWT.ListIndex = Node61_Data1.readEWT
    End If
    
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me
End Sub

Private Sub optTransferTo_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub T_groupid_Change()
    f_DataChanged = True
End Sub

Private Sub T_groupid_GotFocus()
    T_groupid.SelStart = 0
    T_groupid.SelLength = Len(T_groupid)
End Sub

Private Sub T_groupid_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
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

'Michael Addd @ July-9-07
Private Sub txtLoginID_Change()
    f_DataChanged = True
End Sub

'Michael Addd @ July-9-07
Private Sub txtLoginID_keypress(KeyAscii As Integer)
    KeyPress KeyAscii
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

'Michael Add @ July-7-07
Private Sub txtNdWait_Change()
    f_DataChanged = True
End Sub

'Michael Add @ July -7 -07
Private Sub txtNdWait_GotFocus()
    txtNdWait.SelStart = 0
    txtNdWait.SelLength = Len(txtNdWait)
End Sub

'Michael Addd @ July-9-07
Private Sub txtndwait_keypress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

'Michael Added July,9,07
Private Sub txtUserID_Change()
    f_DataChanged = True
End Sub

'Michael Added July,9,07
Private Sub txtUserID_keypress(KeyAscii As Integer)
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

Private Sub txtWaitTime_Change()
    f_DataChanged = True
End Sub

Private Sub txtWaitTime_GotFocus()
    txtWaitTime.SelStart = 0
    txtWaitTime.SelLength = Len(txtWaitTime)
End Sub

Private Sub txtWaitTime_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

'Michael Added @ Jul,10,07
Private Sub txtWaitansto_Change()
    f_DataChanged = True
End Sub

'Michael Added @ Jul,10,07
Private Sub txtWaitansto_GotFocus()
    txtWaitansto.SelStart = 0
    txtWaitansto.SelLength = Len(txtWaitansto)
End Sub

'Michael Added @ Jul,10,07
Private Sub txtWaitansto_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

'Michael Added @ July,9,07
'Fill the combo box
Private Sub FillcboReadEWT()
    With cboReadEWT
        .Clear
        .AddItem LoadNationalResString(1592), 0
        .AddItem LoadNationalResString(1593), 1
        .AddItem LoadNationalResString(1594), 2
        .AddItem LoadNationalResString(1815), 3
    End With
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
        Cb_agtInfoLen.ListIndex = SearchItemDataIndex(Cb_agtInfoLen, CLng(Node61_Data2.length_agentinfo), 4)
    Else
        Cb_agtInfoLen.ListIndex = 0
    End If
End Sub

Private Sub Cb_agtInfoLen_Click()
     f_DataChanged = True
End Sub
' Added End
