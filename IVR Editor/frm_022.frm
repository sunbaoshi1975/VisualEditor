VERSION 5.00
Begin VB.Form frm_022 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "放音等待按键"
   ClientHeight    =   4860
   ClientLeft      =   2790
   ClientTop       =   2550
   ClientWidth     =   8790
   ForeColor       =   &H00000000&
   Icon            =   "frm_022.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1260"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   8400
      Picture         =   "frm_022.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "1948"
      Top             =   4431
      Width           =   333
   End
   Begin VB.Frame Frame4 
      Caption         =   "描述"
      Height          =   855
      Left            =   60
      TabIndex        =   48
      Tag             =   "1104"
      Top             =   3870
      Width           =   4035
      Begin VB.TextBox Txt_Description 
         Height          =   525
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   240
         Width           =   3825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   4215
      Left            =   4170
      TabIndex        =   47
      Tag             =   "1164"
      Top             =   60
      Width           =   4545
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   13
         Left            =   3450
         Picture         =   "frm_022.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "1145"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   12
         Left            =   3450
         Picture         =   "frm_022.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "1145"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   11
         Left            =   4080
         Picture         =   "frm_022.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "1145"
         Top             =   3690
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   10
         Left            =   4080
         Picture         =   "frm_022.frx":29DA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "1145"
         Top             =   3330
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   9
         Left            =   4080
         Picture         =   "frm_022.frx":2D64
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "1145"
         Top             =   2970
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   8
         Left            =   4080
         Picture         =   "frm_022.frx":30EE
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "1145"
         Top             =   2610
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   7
         Left            =   4080
         Picture         =   "frm_022.frx":3478
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "1145"
         Top             =   2250
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   6
         Left            =   4080
         Picture         =   "frm_022.frx":3802
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "1145"
         Top             =   1890
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   5
         Left            =   1890
         Picture         =   "frm_022.frx":3B8C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "1145"
         Top             =   3690
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   4
         Left            =   1890
         Picture         =   "frm_022.frx":3F16
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "1145"
         Top             =   3330
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   1890
         Picture         =   "frm_022.frx":42A0
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "1145"
         Top             =   2970
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   1890
         Picture         =   "frm_022.frx":462A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "1145"
         Top             =   2610
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   1890
         Picture         =   "frm_022.frx":49B4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "1145"
         Top             =   2250
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   1890
         Picture         =   "frm_022.frx":4D3E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "1145"
         Top             =   1890
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   3450
         Picture         =   "frm_022.frx":50C8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "1146"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   3450
         Picture         =   "frm_022.frx":51CA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
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
         Index           =   0
         Left            =   900
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1860
         Width           =   975
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
         Index           =   11
         Left            =   3090
         MaxLength       =   6
         TabIndex        =   38
         Top             =   3660
         Width           =   975
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
         Index           =   10
         Left            =   3090
         MaxLength       =   6
         TabIndex        =   36
         Top             =   3300
         Width           =   975
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
         Index           =   9
         Left            =   3090
         MaxLength       =   6
         TabIndex        =   34
         Top             =   2940
         Width           =   975
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
         Index           =   8
         Left            =   3090
         MaxLength       =   6
         TabIndex        =   32
         Top             =   2580
         Width           =   975
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
         Index           =   7
         Left            =   3090
         MaxLength       =   6
         TabIndex        =   30
         Top             =   2220
         Width           =   975
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
         Index           =   6
         Left            =   3090
         MaxLength       =   6
         TabIndex        =   28
         Top             =   1860
         Width           =   975
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
         Index           =   5
         Left            =   900
         MaxLength       =   6
         TabIndex        =   26
         Top             =   3660
         Width           =   975
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
         Index           =   4
         Left            =   900
         MaxLength       =   6
         TabIndex        =   24
         Top             =   3300
         Width           =   975
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
         Index           =   3
         Left            =   900
         MaxLength       =   6
         TabIndex        =   22
         Top             =   2940
         Width           =   975
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
         Index           =   2
         Left            =   900
         MaxLength       =   6
         TabIndex        =   20
         Top             =   2580
         Width           =   975
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
         Index           =   1
         Left            =   900
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2220
         Width           =   975
      End
      Begin VB.TextBox T_nd_Fail 
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   14
         Top             =   1320
         Width           =   1035
      End
      Begin VB.TextBox T_vox_nodefail 
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   10
         Top             =   600
         Width           =   1035
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   8
         Top             =   240
         Width           =   1035
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   12
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "按0》转"
         Height          =   195
         Left            =   180
         TabIndex        =   72
         Tag             =   "1267"
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "按6》转"
         Height          =   180
         Left            =   2370
         TabIndex        =   71
         Tag             =   "1273"
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "按1》转"
         Height          =   180
         Left            =   180
         TabIndex        =   70
         Tag             =   "1268"
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "按2》转"
         Height          =   180
         Left            =   180
         TabIndex        =   69
         Tag             =   "1269"
         Top             =   2670
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "按3》转"
         Height          =   180
         Left            =   180
         TabIndex        =   68
         Tag             =   "1270"
         Top             =   3030
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "按4》转"
         Height          =   180
         Left            =   180
         TabIndex        =   67
         Tag             =   "1271"
         Top             =   3390
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "按5》转"
         Height          =   180
         Left            =   180
         TabIndex        =   66
         Tag             =   "1272"
         Top             =   3750
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "按7》转"
         Height          =   180
         Left            =   2370
         TabIndex        =   65
         Tag             =   "1274"
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "按8》转"
         Height          =   180
         Left            =   2370
         TabIndex        =   64
         Tag             =   "1275"
         Top             =   2670
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "按9》转"
         Height          =   180
         Left            =   2370
         TabIndex        =   63
         Tag             =   "1276"
         Top             =   3030
         Width           =   630
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "按*》转"
         Height          =   180
         Left            =   2370
         TabIndex        =   62
         Tag             =   "1277"
         Top             =   3390
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "按#》转"
         Height          =   180
         Left            =   2370
         TabIndex        =   61
         Tag             =   "1278"
         Top             =   3750
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "失败转入节点ID"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   60
         Tag             =   "1171"
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "操作节点错误播放语音"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   59
         Tag             =   "1552"
         Top             =   690
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "操作提示语音"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   58
         Tag             =   "1266"
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   57
         Tag             =   "1169"
         Top             =   1020
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   60
      TabIndex        =   43
      Tag             =   "1136"
      Top             =   60
      Width           =   4035
      Begin VB.ComboBox Cb_maxtrytime 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "请选择"
         Top             =   1050
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "请选择"
         Top             =   2490
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "请选择"
         Top             =   1770
         Width           =   1965
      End
      Begin VB.ComboBox Cb_MaxInterval 
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
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "0 - 255 秒"
         Top             =   1410
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2850
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
         Left            =   1920
         TabIndex        =   0
         ToolTipText     =   "0 - 255 秒"
         Top             =   690
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   50
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
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
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
         Left            =   1920
         TabIndex        =   4
         ToolTipText     =   "0 - 255 个键"
         Top             =   2130
         Width           =   1965
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3210
         Width           =   1965
      End
      Begin VB.Label Lbl_trytime 
         AutoSize        =   -1  'True
         Caption         =   "最大尝试次数"
         Height          =   195
         Left            =   180
         TabIndex        =   73
         Tag             =   "1158"
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "放音清空标志"
         Height          =   195
         Left            =   180
         TabIndex        =   56
         Tag             =   "1262"
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "按键最大间隔(秒)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   55
         Tag             =   "1261"
         Top             =   1500
         Width           =   1440
      End
      Begin VB.Label Lbl_log 
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   54
         Tag             =   "1245"
         Top             =   2910
         Width           =   915
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "节点超时(秒)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   53
         Tag             =   "1154"
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1950
         TabIndex        =   51
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "按键中断"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Tag             =   "1264"
         Top             =   2580
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "按键记录"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Tag             =   "1265"
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "按键长度"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Tag             =   "1263"
         Top             =   2220
         Width           =   720
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   6540
      TabIndex        =   42
      Tag             =   "1144"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   5220
      TabIndex        =   41
      Tag             =   "1007"
      Top             =   4440
      Width           =   1035
   End
End
Attribute VB_Name = "frm_022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_022.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：放音等待按键
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_breakkey_Click()
    f_DataChanged = True
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

Private Sub Cb_MaxInterval_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_MaxInterval_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_MaxInterval_GotFocus()
    Cb_MaxInterval.SelStart = 0
    Cb_MaxInterval.SelLength = Len(Cb_MaxInterval)
End Sub

Private Sub Cb_MaxInterval_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Cb_maxtrytime_Click()
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

Private Sub Cb_var_key_Click()
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
    Case 12
        Set gSystem.crlCurItem = T_nd_parent
    Case 13
        Set gSystem.crlCurItem = T_nd_Fail
    Case Is < 12
        Set gSystem.crlCurItem = T_nd_child(Index)
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click(Index As Integer)
    
    Select Case Index
    Case 0
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_play
    Case 1
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_nodefail
    End Select
    frmResourceList.Show vbModal

End Sub

Public Sub CommandSave_Click()
On Error Resume Next

Dim i As Integer

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' Sun added 2007-03-20
        '最大尝试次数
        If Cb_maxtrytime.ListIndex < 0 Then
            Node22_Data1.maxtrytime = 3
        Else
            Node22_Data1.maxtrytime = CByte(Cb_maxtrytime.ItemData(Cb_maxtrytime.ListIndex))
        End If
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node22_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node22_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        ' 按键最大间隔
        If Trim(Cb_MaxInterval) = "" Then
            Node22_Data1.maxinterval = 0
        Else
            If Val(Cb_MaxInterval) > 255 Then
                Message ("E082")
                Cb_MaxInterval.SetFocus
                Exit Sub
            Else
                Node22_Data1.maxinterval = CByte(Val(Cb_MaxInterval.Text) Mod 256)
            End If
        End If
    
        ' 放音清空
        If Cb_playclear.ListIndex = -1 Then
            Node22_Data1.playclear = 0
        Else
            Node22_Data1.playclear = CByte(Cb_playclear.ItemData(Cb_playclear.ListIndex))
        End If
        
        ' 按键长度
        If Trim(Cb_getlength) = "" Then
            Node22_Data1.getlength = 0
        Else
            If CLng(Trim(Cb_getlength)) > 255 Then
                Message ("E081")
                Cb_getlength.SetFocus
                Exit Sub
            Else
                Node22_Data1.getlength = CByte(Val(Cb_getlength.Text) Mod 256)
            End If
        End If

        ' 按键中断
        If Cb_breakkey.ListIndex < 0 Then
            Node22_Data1.breakkey = 0
        Else
            Node22_Data1.breakkey = CByte(Cb_breakkey.ItemData(Cb_breakkey.ListIndex))
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node22_Data1.log = 0
        Else
            Node22_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        ' 按键记录
        If Cb_var_key.ListIndex <= 0 Then
            Node22_Data1.var_key = 0
        Else
            Node22_Data1.var_key = CByte(Cb_var_key.ItemData(Cb_var_key.ListIndex))
        End If
        
        Node22_Data1.reserved1(0) = 0
        
        ' 播放语音
        If Trim(T_vox_play) = "" Then
            Node22_Data2.vox_play = 0
        Else
            If CLng(Trim(T_vox_play)) > 32767 Then
                Message ("E076")
                T_vox_play.SetFocus
                Exit Sub
            Else
                Node22_Data2.vox_play = CInt(Trim(T_vox_play.Text))
            End If
        End If
        
        ' 操作节点错误播放语音
        If Trim(T_vox_nodefail) = "" Then
            Node22_Data2.vox_nodefail = 0
        Else
            If CLng(Trim(T_vox_nodefail)) > 32767 Then
                Message ("E076")
                T_vox_nodefail.SetFocus
                Exit Sub
            Else
                Node22_Data2.vox_nodefail = CInt(Trim(T_vox_nodefail.Text))
            End If
        End If
        
        Node22_Data2.reserved1(0) = 0
                
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node22_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node22_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
        
        ' 失败转入节点ID
        If Trim(T_nd_Fail) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_Fail) > 32767 Or Val(T_nd_Fail) < 256) And Val(T_nd_Fail) <> 0 Then
                Message ("E070")
                T_nd_Fail.SetFocus
                Exit Sub
            Else
                lv_nNewNode = CInt(Trim(T_nd_Fail.Text))
            End If
        End If
        
        '' Sun added 2007-03-25
        If Node22_Data2.nd_nodefail <> lv_nNewNode Then
            Node22_Data2.nd_nodefail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 255)
        End If
        
        ' 按键子节点
        For i = T_nd_child.LBound To T_nd_child.UBound
            
            If Trim(T_nd_child(i).Text) = "" Then
                lv_nNewNode = 0
            Else
                If (Val(T_nd_child(i)) > 32767 Or Val(T_nd_child(i)) < 256) And Val(T_nd_child(i)) <> 0 Then
                    Message ("E100")
                    T_nd_child(i).SetFocus
                    Exit Sub
                Else
                    lv_nNewNode = CInt(Trim(T_nd_child(i)))
                End If
            End If

            '' Sun added 2007-03-25
            If Node22_Data2.nd_key(i) <> lv_nNewNode Then
                Node22_Data2.nd_key(i) = lv_nNewNode
                Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, i)
            End If
        
        Next
        
        Node22_Data2.reserved2(0) = 0
        
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

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer

    ' Sun added 2007-03-20
    '最大尝试次数
    With Cb_maxtrytime
        For i = 0 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1180)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
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

    '输入终止符
    F_FillPhoneKeyList Cb_breakkey, 1

    ' 节点超时(秒)
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
    ' 按键最大间隔(秒)
    With Cb_MaxInterval
        For i = 0 To 9
            .AddItem Trim(Str(i))
        Next
    End With
    
    ' 按键长度
    With Cb_getlength
        For i = 0 To 9
            .AddItem Trim(Str(i))
        Next
    End With
  
    ' 按键记录
    RefreshVariablesList Cb_var_key
       
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_play.Text = Node22_Data2.vox_play
    T_vox_nodefail = Node22_Data2.vox_nodefail
    T_nd_parent.Text = Node22_Data2.nd_parent
    T_nd_Fail = Node22_Data2.nd_nodefail
    For i = T_nd_child.LBound To T_nd_child.UBound
        T_nd_child(i) = Node22_Data2.nd_key(i)
    Next
    
    Cb_timeout.Text = Node22_Data1.timeout
    Cb_MaxInterval.Text = Node22_Data1.maxinterval
    Cb_playclear.ListIndex = SearchItemDataIndex(Cb_playclear, CLng(Node22_Data1.playclear), 0)
    Cb_getlength.Text = Node22_Data1.getlength
    Cb_breakkey.ListIndex = SearchItemDataIndex(Cb_breakkey, CLng(Node22_Data1.breakkey), 11)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node22_Data1.log), 0)
    
    ' Sun added 2007-03-20
    Cb_maxtrytime.ListIndex = SearchItemDataIndex(Cb_maxtrytime, _
                                    CLng(Node22_Data1.maxtrytime), 3)
    
    Cb_var_key.ListIndex = SearchItemDataIndex(Cb_var_key, CLng(Node22_Data1.var_key), 0)
    
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

Private Sub CommandExit_Click()
   Unload Me
End Sub

Private Sub T_nd_child_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub T_nd_child_GotFocus(Index As Integer)
    T_nd_child(Index).SelStart = 0
    T_nd_child(Index).SelLength = Len(T_nd_child(Index))
End Sub

Private Sub T_nd_child_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_nd_fail_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_fail_GotFocus()
    T_nd_Fail.SelStart = 0
    T_nd_Fail.SelLength = Len(T_nd_Fail)
End Sub

Private Sub T_nd_fail_KeyPress(KeyAscii As Integer)
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

Private Sub T_vox_nodefail_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_nodefail

End Sub

Private Sub T_vox_nodefail_GotFocus()
    T_vox_nodefail.SelStart = 0
    T_vox_nodefail.SelLength = Len(T_vox_nodefail)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_nodefail)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_nodefail_KeyPress(KeyAscii As Integer)
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
