VERSION 5.00
Begin VB.Form frm_009 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间分支"
   ClientHeight    =   8595
   ClientLeft      =   2955
   ClientTop       =   2610
   ClientWidth     =   6960
   Icon            =   "frm_009.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1183"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6552
      Picture         =   "frm_009.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   116
      ToolTipText     =   "1948"
      Top             =   8211
      Width           =   333
   End
   Begin VB.Frame Frame5 
      Caption         =   "工作日时间段安排"
      Height          =   675
      Left            =   60
      TabIndex        =   115
      Tag             =   "1193"
      Top             =   2040
      Width           =   6825
      Begin VB.CheckBox chkTimeSec 
         Caption         =   "时间段6"
         Height          =   345
         Index           =   5
         Left            =   5700
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkTimeSec 
         Caption         =   "时间段5"
         Height          =   345
         Index           =   4
         Left            =   4596
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkTimeSec 
         Caption         =   "时间段4"
         Height          =   345
         Index           =   3
         Left            =   3492
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkTimeSec 
         Caption         =   "时间段3"
         Height          =   345
         Index           =   2
         Left            =   2388
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkTimeSec 
         Caption         =   "时间段2"
         Height          =   345
         Index           =   1
         Left            =   1284
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkTimeSec 
         Caption         =   "时间段1"
         Height          =   345
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "工作日安排"
      Height          =   1155
      Left            =   60
      TabIndex        =   113
      Tag             =   "1184"
      Top             =   840
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   6390
         Picture         =   "frm_009.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "1145"
         Top             =   690
         Width           =   315
      End
      Begin VB.TextBox T_nd_sparetime 
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
         Left            =   5490
         MaxLength       =   6
         TabIndex        =   11
         Top             =   660
         Width           =   885
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "星期日"
         Height          =   315
         Index           =   0
         Left            =   1485
         TabIndex        =   10
         Tag             =   "1191"
         Top             =   683
         Width           =   1185
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "星期一"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Tag             =   "1185"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "星期二"
         Height          =   315
         Index           =   2
         Left            =   1485
         TabIndex        =   5
         Tag             =   "1186"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "星期三"
         Height          =   315
         Index           =   3
         Left            =   2820
         TabIndex        =   6
         Tag             =   "1187"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "星期四"
         Height          =   315
         Index           =   4
         Left            =   4155
         TabIndex        =   7
         Tag             =   "1188"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "星期五"
         Height          =   315
         Index           =   5
         Left            =   5490
         TabIndex        =   8
         Tag             =   "1189"
         Top             =   300
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox chkWorkDay 
         Caption         =   "星期六"
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   9
         Tag             =   "1190"
         Top             =   683
         Width           =   1185
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "休息日转节点ID"
         Height          =   195
         Left            =   4140
         TabIndex        =   114
         Tag             =   "1192"
         Top             =   750
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "时间段6"
      Enabled         =   0   'False
      Height          =   705
      Index           =   5
      Left            =   60
      TabIndex        =   96
      Top             =   6510
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   7
         Left            =   6360
         Picture         =   "frm_009.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.ComboBox Cmbts_h1 
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
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tts_m1 
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   98
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Cmbte_h1 
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
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tte_m1 
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
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   100
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox T_nd_time1 
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
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   101
         Text            =   "0000"
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "起始时间"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   112
         Tag             =   "1195"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   5
         Left            =   2220
         TabIndex        =   111
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label7 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   1620
         TabIndex        =   110
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "终止时间"
         Height          =   195
         Index           =   5
         Left            =   2550
         TabIndex        =   109
         Tag             =   "1196"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   3990
         TabIndex        =   107
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label11 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   5
         Left            =   4560
         TabIndex        =   105
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Index           =   5
         Left            =   4890
         TabIndex        =   103
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   2430
         X2              =   2430
         Y1              =   210
         Y2              =   600
      End
      Begin VB.Line Line2 
         Index           =   5
         X1              =   4770
         X2              =   4770
         Y1              =   210
         Y2              =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "时间段5"
      Enabled         =   0   'False
      Height          =   705
      Index           =   4
      Left            =   60
      TabIndex        =   82
      Top             =   5760
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   6
         Left            =   6360
         Picture         =   "frm_009.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.ComboBox Cmbts_h1 
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
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tts_m1 
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   84
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Cmbte_h1 
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
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tte_m1 
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
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   86
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox T_nd_time1 
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
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   87
         Text            =   "0000"
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "起始时间"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   95
         Tag             =   "1195"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   4
         Left            =   2220
         TabIndex        =   94
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label7 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   1620
         TabIndex        =   93
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "终止时间"
         Height          =   195
         Index           =   4
         Left            =   2550
         TabIndex        =   92
         Tag             =   "1196"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   3990
         TabIndex        =   91
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label11 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   4
         Left            =   4560
         TabIndex        =   90
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Index           =   4
         Left            =   4890
         TabIndex        =   89
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   2430
         X2              =   2430
         Y1              =   210
         Y2              =   600
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   4770
         X2              =   4770
         Y1              =   210
         Y2              =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "时间段4"
      Enabled         =   0   'False
      Height          =   705
      Index           =   3
      Left            =   60
      TabIndex        =   68
      Top             =   5010
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   5
         Left            =   6360
         Picture         =   "frm_009.frx":29DA
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.ComboBox Cmbts_h1 
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
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tts_m1 
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   70
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Cmbte_h1 
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
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tte_m1 
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
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   72
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox T_nd_time1 
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
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   73
         Text            =   "0000"
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "起始时间"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   81
         Tag             =   "1195"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   2220
         TabIndex        =   80
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label7 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   1620
         TabIndex        =   79
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "终止时间"
         Height          =   195
         Index           =   3
         Left            =   2550
         TabIndex        =   78
         Tag             =   "1196"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   3990
         TabIndex        =   77
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label11 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   76
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Index           =   3
         Left            =   4890
         TabIndex        =   75
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2430
         X2              =   2430
         Y1              =   210
         Y2              =   600
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   4770
         X2              =   4770
         Y1              =   210
         Y2              =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "时间段3"
      Enabled         =   0   'False
      Height          =   705
      Index           =   2
      Left            =   60
      TabIndex        =   54
      Top             =   4260
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   4
         Left            =   6360
         Picture         =   "frm_009.frx":2D64
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.ComboBox Cmbts_h1 
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
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tts_m1 
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   56
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Cmbte_h1 
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
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tte_m1 
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
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   58
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox T_nd_time1 
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
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   59
         Text            =   "0000"
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "起始时间"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   67
         Tag             =   "1195"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   2
         Left            =   2220
         TabIndex        =   66
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label7 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   1620
         TabIndex        =   65
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "终止时间"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   64
         Tag             =   "1196"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   3990
         TabIndex        =   63
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label11 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   2
         Left            =   4560
         TabIndex        =   62
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Index           =   2
         Left            =   4890
         TabIndex        =   61
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2430
         X2              =   2430
         Y1              =   210
         Y2              =   600
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   4770
         X2              =   4770
         Y1              =   210
         Y2              =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "时间段2"
      Enabled         =   0   'False
      Height          =   705
      Index           =   1
      Left            =   60
      TabIndex        =   40
      Top             =   3510
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   6360
         Picture         =   "frm_009.frx":30EE
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.ComboBox Cmbts_h1 
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
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tts_m1 
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   42
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Cmbte_h1 
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
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tte_m1 
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
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   44
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox T_nd_time1 
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
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   45
         Text            =   "0000"
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "起始时间"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   53
         Tag             =   "1195"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   1
         Left            =   2220
         TabIndex        =   52
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label7 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   51
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "终止时间"
         Height          =   195
         Index           =   1
         Left            =   2550
         TabIndex        =   50
         Tag             =   "1196"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   3990
         TabIndex        =   49
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label11 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   48
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Index           =   1
         Left            =   4890
         TabIndex        =   47
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2430
         X2              =   2430
         Y1              =   210
         Y2              =   600
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   4770
         X2              =   4770
         Y1              =   210
         Y2              =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   855
      Left            =   60
      TabIndex        =   35
      Tag             =   "1104"
      Top             =   7260
      Width           =   6825
      Begin VB.TextBox Txt_Description 
         Height          =   495
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   104
         Top             =   270
         Width           =   6645
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "时间段1"
      Enabled         =   0   'False
      Height          =   705
      Index           =   0
      Left            =   60
      TabIndex        =   21
      Top             =   2760
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   6360
         Picture         =   "frm_009.frx":3478
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox T_nd_time1 
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
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   33
         Text            =   "0000"
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox Tte_m1 
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
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Cmbte_h1 
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
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox Tts_m1 
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Cmbts_h1 
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
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   645
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   4770
         X2              =   4770
         Y1              =   210
         Y2              =   600
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2430
         X2              =   2430
         Y1              =   210
         Y2              =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Index           =   0
         Left            =   4890
         TabIndex        =   32
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label11 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   4560
         TabIndex        =   31
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label10 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   3990
         TabIndex        =   29
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "终止时间"
         Height          =   195
         Index           =   0
         Left            =   2550
         TabIndex        =   27
         Tag             =   "1196"
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   1620
         TabIndex        =   26
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   2220
         TabIndex        =   24
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "起始时间"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Tag             =   "1195"
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   2100
      TabIndex        =   106
      Tag             =   "1007"
      Top             =   8220
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   3630
      TabIndex        =   108
      Tag             =   "1144"
      Top             =   8220
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Tag             =   "1136"
      Top             =   60
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   6390
         Picture         =   "frm_009.frx":3802
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
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
         TabIndex        =   37
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
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
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
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   240
         Width           =   1095
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
         Left            =   5580
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1260
         TabIndex        =   39
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   2790
         TabIndex        =   20
         Tag             =   "1159"
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "父节点"
         Height          =   180
         Left            =   4980
         TabIndex        =   19
         Tag             =   "1150"
         Top             =   300
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_009.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：时间分支
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub chkTimeSec_Click(Index As Integer)
    f_DataChanged = True
    Frame3(Index).Enabled = (chkTimeSec(Index).value = vbChecked)
End Sub

Private Sub chkWorkDay_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Cmbte_h1_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Cmbts_h1_Click(Index As Integer)
    f_DataChanged = True
End Sub

'Mike added this event @ 2008-1-30
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
        Set gSystem.crlCurItem = T_nd_sparetime
    Case Else
        Set gSystem.crlCurItem = T_nd_time1(Index - 2)
    End Select
    frmNodeList.Show vbModal

End Sub

Public Sub Command1_Click()
On Error Resume Next

Dim i As Integer

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        '保留
        Node9_Data1.reserved1(0) = 0
        
        '被访问日志
        If Cb_log.ListIndex < 0 Then
           Node9_Data1.log = 0
        Else
           Node9_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        '保留
        Node9_Data1.reserved2(0) = 0
        
        '工作日安排
        Node9_Data2.workday = 0
        For i = chkWorkDay.LBound To chkWorkDay.UBound
            If chkWorkDay(i).value = vbChecked Then
                Call Set_Bit_Value(Node9_Data2.workday, CByte(i), 1)
            Else
                Call Set_Bit_Value(Node9_Data2.workday, CByte(i), 0)
            End If
        Next

        '工作日时间段安排
        Node9_Data2.worktime = 0
        For i = chkTimeSec.LBound To chkTimeSec.UBound
            If chkTimeSec(i).value = vbChecked Then
                Call Set_Bit_Value(Node9_Data2.worktime, CByte(i), 1)
            Else
                Call Set_Bit_Value(Node9_Data2.worktime, CByte(i), 0)
            End If
        Next
        
        '时间段
        For i = Frame3.LBound To Frame3.UBound
            
            If Frame3(i).Enabled = True Then
            
                '' 起始时间
                ''' HH
                If Cmbts_h1(i).ListIndex < 0 Then
                    Node9_Data2.timesec(i * 4) = 0
                Else
                    Node9_Data2.timesec(i * 4) = CByte(Cmbts_h1(i).ItemData(Cmbts_h1(i).ListIndex))
                End If
                ''' MM
                If Trim(Tts_m1(i)) = "" Then
                    Node9_Data2.timesec(i * 4 + 1) = 0
                Else
                    If CLng(Trim(Tts_m1(i))) > 59 Then
                        Message ("E057")
                        Tts_m1(i).SetFocus
                        Exit Sub
                    Else
                        Node9_Data2.timesec(i * 4 + 1) = CByte(Trim(Tts_m1(i)))
                    End If
                End If
                
                '' 终止时间
                ''' HH
                If Cmbte_h1(i).ListIndex < 0 Then
                    Node9_Data2.timesec(i * 4 + 2) = 0
                Else
                    Node9_Data2.timesec(i * 4 + 2) = CByte(Cmbte_h1(i).ItemData(Cmbte_h1(i).ListIndex))
                End If
                ''' MM
                If Trim(Tte_m1(i)) = "" Then
                    Node9_Data2.timesec(i * 4 + 3) = 0
                Else
                    If CLng(Trim(Tte_m1(i))) > 59 Then
                        Message ("E057")
                        Tte_m1(i).SetFocus
                        Exit Sub
                    Else
                        Node9_Data2.timesec(i * 4 + 3) = CByte(Trim(Tte_m1(i)))
                    End If
                End If
                
                '' 时间段判断
                If CInt(Cmbte_h1(i)) < CInt(Cmbts_h1(i)) Then
                    Message ("E066")
                    Cmbts_h1(i).SetFocus
                    Exit Sub
                ElseIf CInt(Cmbte_h1(i)) = CInt(Cmbts_h1(i)) Then
                    If CInt(Tte_m1(i)) <= CInt(Tts_m1(i)) Then
                        Message ("E066")
                        Tts_m1(i).SetFocus
                        Exit Sub
                    End If
                End If
                
                '' 跳转节点
                If Trim(T_nd_time1(i)) = "" Then
                    lv_nNewNode = 0
                Else
                    If (Val(T_nd_time1(i)) > 32767 Or Val(T_nd_time1(i)) < 256) And Val(T_nd_time1(i)) <> 0 Then
                        Message ("E059")
                        T_nd_time1(i).SetFocus
                        Exit Sub
                    Else
                        lv_nNewNode = CInt(Trim(T_nd_time1(i)))
                    End If
                End If
                
                '' Sun added 2007-03-25
                If Node9_Data2.nd_timesec(i) <> lv_nNewNode Then
                    Node9_Data2.nd_timesec(i) = lv_nNewNode
                    Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, i + 1)
                End If
                
            End If
            
        Next
        
        '保留
        Node9_Data2.reserved1(0) = 0
        
        '父节点
        If Trim(T_nd_parent) = "" Then
            Node9_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node9_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
                
        '休息日转节点
        If Trim(T_nd_sparetime) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_sparetime) > 32767 Or Val(T_nd_sparetime) < 256) And Val(T_nd_sparetime) <> 0 Then
                Message ("E065")
                T_nd_sparetime.SetFocus
                Exit Sub
            Else
                lv_nNewNode = CInt(Trim(T_nd_sparetime.Text))
            End If
        End If
        
        '' Sun added 2007-03-25
        If Node9_Data2.nd_sparetime <> lv_nNewNode Then
            Node9_Data2.nd_sparetime = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
        End If
        
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

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_Load()
On Error Resume Next
chkTimeSec(0).Caption = LoadNationalResString(1194) & "1"
chkTimeSec(1).Caption = LoadNationalResString(1194) & "2"
chkTimeSec(2).Caption = LoadNationalResString(1194) & "3"
chkTimeSec(3).Caption = LoadNationalResString(1194) & "4"
chkTimeSec(4).Caption = LoadNationalResString(1194) & "5"
chkTimeSec(5).Caption = LoadNationalResString(1194) & "6"
Frame3(0).Caption = LoadNationalResString(1194) & "1"
Frame3(1).Caption = LoadNationalResString(1194) & "2"
Frame3(2).Caption = LoadNationalResString(1194) & "3"
Frame3(3).Caption = LoadNationalResString(1194) & "4"
Frame3(4).Caption = LoadNationalResString(1194) & "5"
Frame3(5).Caption = LoadNationalResString(1194) & "6"
SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
Dim j As Integer

    '被访日志
    With Cb_log
        .AddItem LoadNationalResString(1178)
        .ItemData(.ListCount - 1) = 0
        For i = 1 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1179)
            .ItemData(.ListCount - 1) = i
        Next
    End With

    For i = Frame3.LBound To Frame3.UBound
        For j = 0 To 23
            Cmbts_h1(i).AddItem Trim(Str(j))
            Cmbts_h1(i).ItemData(Cmbts_h1(i).ListCount - 1) = j
            Cmbte_h1(i).AddItem Trim(Str(j))
            Cmbte_h1(i).ItemData(Cmbte_h1(i).ListCount - 1) = j
        Next
        Cmbts_h1(i).ListIndex = 0
        Cmbte_h1(i).ListIndex = 0
        Tts_m1(i) = "0"
        Tte_m1(i) = "0"
        T_nd_time1(i) = "0"
    Next
        
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node9_Data1.log), 0)
    T_nd_parent.Text = Node9_Data2.nd_parent
    T_nd_sparetime.Text = Node9_Data2.nd_sparetime
    
    For i = chkWorkDay.LBound To chkWorkDay.UBound
        If Get_Bit_Value(Node9_Data2.workday, CByte(i)) = 1 Then
            chkWorkDay(i).value = vbChecked
        Else
            chkWorkDay(i).value = vbUnchecked
        End If
    Next
    
    For i = chkTimeSec.LBound To chkTimeSec.UBound
        If Get_Bit_Value(Node9_Data2.worktime, CByte(i)) = 1 Then
            chkTimeSec(i).value = vbChecked
            Frame3(i).Enabled = True
        Else
            chkTimeSec(i).value = vbUnchecked
            Frame3(i).Enabled = False
        End If
    Next
    
    For i = Frame3.LBound To Frame3.UBound
        Cmbts_h1(i).ListIndex = SearchItemDataIndex(Cmbts_h1(i), CLng(Node9_Data2.timesec(i * 4)), 0)
        Tts_m1(i) = Trim(Str(Node9_Data2.timesec(i * 4 + 1)))
        Cmbte_h1(i).ListIndex = SearchItemDataIndex(Cmbte_h1(i), CLng(Node9_Data2.timesec(i * 4 + 2)), 0)
        Tte_m1(i) = Trim(Str(Node9_Data2.timesec(i * 4 + 3)))
        T_nd_time1(i) = Node9_Data2.nd_timesec(i)
    Next
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
        
   
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

Private Sub T_nd_sparetime_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_sparetime_GotFocus()
    T_nd_sparetime.SelStart = 0
    T_nd_sparetime.SelLength = Len(T_nd_sparetime)
End Sub

Private Sub T_nd_sparetime_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

Private Sub T_nd_time1_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub T_nd_time1_GotFocus(Index As Integer)
    T_nd_time1(Index).SelStart = 0
    T_nd_time1(Index).SelLength = Len(T_nd_time1(Index))
End Sub

Private Sub T_nd_time1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Tte_m1_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Tte_m1_GotFocus(Index As Integer)
    Tte_m1(Index).SelStart = 0
    Tte_m1(Index).SelLength = Len(Tte_m1(Index))
End Sub

Private Sub Tte_m1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Tts_m1_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Tts_m1_GotFocus(Index As Integer)
    Tts_m1(Index).SelStart = 0
    Tts_m1(Index).SelLength = Len(Tts_m1(Index))
End Sub

Private Sub Tts_m1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
