VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_096 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "异步通信"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "frm_096.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1568"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   10242
      Picture         =   "frm_096.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "1948"
      Top             =   4605
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   4425
      Index           =   0
      Left            =   4080
      TabIndex        =   53
      Tag             =   "1164"
      Top             =   60
      Width           =   6495
      Begin VB.CheckBox chkCarryOnPlay 
         Caption         =   "通信完成后是否继续放音"
         Height          =   255
         Left            =   1500
         TabIndex        =   41
         ToolTipText     =   "2139"
         Top             =   4050
         Width           =   4065
      End
      Begin VB.PictureBox picVars 
         BorderStyle     =   0  'None
         Height          =   2175
         Index           =   1
         Left            =   180
         ScaleHeight     =   2175
         ScaleWidth      =   6015
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   6015
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   60
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   60
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   480
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   480
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   900
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   900
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1320
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1320
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1740
            Width           =   2025
         End
         Begin VB.ComboBox Cb_rcvvar 
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
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1740
            Width           =   2025
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 1"
            Height          =   195
            Index           =   19
            Left            =   60
            TabIndex        =   78
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 2"
            Height          =   195
            Index           =   18
            Left            =   3300
            TabIndex        =   77
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 3"
            Height          =   195
            Index           =   17
            Left            =   60
            TabIndex        =   76
            Top             =   540
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 4"
            Height          =   195
            Index           =   16
            Left            =   3300
            TabIndex        =   75
            Top             =   540
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 5"
            Height          =   195
            Index           =   15
            Left            =   60
            TabIndex        =   74
            Top             =   960
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 6"
            Height          =   195
            Index           =   14
            Left            =   3300
            TabIndex        =   73
            Top             =   960
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 7"
            Height          =   195
            Index           =   13
            Left            =   60
            TabIndex        =   72
            Top             =   1380
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 8"
            Height          =   195
            Index           =   12
            Left            =   3300
            TabIndex        =   71
            Top             =   1380
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 9"
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   70
            Top             =   1800
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 10"
            Height          =   195
            Index           =   10
            Left            =   3300
            TabIndex        =   69
            Top             =   1800
            Width           =   390
         End
      End
      Begin VB.PictureBox picVars 
         BorderStyle     =   0  'None
         Height          =   2175
         Index           =   0
         Left            =   180
         ScaleHeight     =   2175
         ScaleWidth      =   6015
         TabIndex        =   21
         Top             =   720
         Width           =   6015
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
            Index           =   9
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1740
            Width           =   2025
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
            Index           =   8
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1740
            Width           =   2025
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
            Index           =   7
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1320
            Width           =   2025
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
            Index           =   6
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1320
            Width           =   2025
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
            Index           =   5
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   900
            Width           =   2025
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
            Index           =   4
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   900
            Width           =   2025
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
            Index           =   3
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   480
            Width           =   2025
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
            Index           =   2
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Width           =   2025
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
            Index           =   1
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   60
            Width           =   2025
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
            Index           =   0
            Left            =   540
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   60
            Width           =   2025
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 10"
            Height          =   195
            Index           =   9
            Left            =   3300
            TabIndex        =   67
            Top             =   1800
            Width           =   390
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 9"
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   66
            Top             =   1800
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 8"
            Height          =   195
            Index           =   7
            Left            =   3300
            TabIndex        =   65
            Top             =   1380
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 7"
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   64
            Top             =   1380
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 6"
            Height          =   195
            Index           =   5
            Left            =   3300
            TabIndex        =   63
            Top             =   960
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 5"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   62
            Top             =   960
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 4"
            Height          =   195
            Index           =   3
            Left            =   3300
            TabIndex        =   61
            Top             =   540
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 3"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   60
            Top             =   540
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 2"
            Height          =   195
            Index           =   1
            Left            =   3300
            TabIndex        =   59
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblVar 
            AutoSize        =   -1  'True
            Caption         =   "ID 1"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   58
            Top             =   120
            Width           =   300
         End
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
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   39
         Top             =   3600
         Width           =   1185
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   6000
         Picture         =   "frm_096.frx":15FC
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "1145"
         Top             =   3630
         Width           =   315
      End
      Begin MSComctlLib.TabStrip tabVars 
         Height          =   2715
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4789
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "发送变量"
               Key             =   "keySendVars"
               Object.Tag             =   "1573"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "接收变量"
               Key             =   "keyReceiveVars"
               Object.Tag             =   "1574"
               ImageVarType    =   2
            EndProperty
         EndProperty
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
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   35
         Top             =   3180
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
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   37
         Top             =   3600
         Width           =   1185
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
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   33
         Top             =   3180
         Width           =   1185
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2700
         Picture         =   "frm_096.frx":1986
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "1146"
         Top             =   3210
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   6000
         Picture         =   "frm_096.frx":1A88
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "1145"
         Top             =   3210
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2700
         Picture         =   "frm_096.frx":1E12
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "1145"
         Top             =   3630
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答转节点"
         Height          =   180
         Index           =   11
         Left            =   3480
         TabIndex        =   57
         Tag             =   "1313"
         Top             =   3660
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "等待音乐"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   56
         Tag             =   "1308"
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "父节点ID"
         Height          =   225
         Left            =   3480
         TabIndex        =   55
         Tag             =   "1169"
         Top             =   3255
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "子节点ID"
         Height          =   225
         Left            =   180
         TabIndex        =   54
         Tag             =   "1252"
         Top             =   3660
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   5700
      TabIndex        =   42
      Tag             =   "1007"
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   7500
      TabIndex        =   43
      Tag             =   "1144"
      Top             =   4620
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   975
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Tag             =   "1104"
      Top             =   4020
      Width           =   3915
      Begin VB.TextBox Txt_Description 
         Height          =   585
         Left            =   150
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   60
      TabIndex        =   7
      Tag             =   "1136"
      Top             =   60
      Width           =   3915
      Begin VB.TextBox txtPrefix 
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2430
         Width           =   585
      End
      Begin VB.ComboBox Cb_var_result 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2880
         Width           =   1965
      End
      Begin VB.TextBox txtCommand 
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
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1560
         Width           =   585
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
         Left            =   1800
         TabIndex        =   0
         ToolTipText     =   "0 - 255 秒"
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txtSeperator 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1980
         Width           =   585
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
         TabIndex        =   45
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   3300
         Width           =   1965
      End
      Begin VB.ComboBox cb_ExtDataType 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1140
         Width           =   1965
      End
      Begin VB.Label lblPrefix 
         AutoSize        =   -1  'True
         Caption         =   "文件名前缀"
         Height          =   195
         Left            =   180
         TabIndex        =   79
         Tag             =   "1582"
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "扩展数据记录"
         Height          =   195
         Left            =   180
         TabIndex        =   68
         Tag             =   "1575"
         Top             =   2940
         Width           =   1080
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "命令代码"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   52
         Tag             =   "1572"
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "通信超时(秒)"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   51
         Tag             =   "1570"
         Top             =   810
         Width           =   990
      End
      Begin VB.Label Lbl_tcpip 
         AutoSize        =   -1  'True
         Caption         =   "分隔符"
         Height          =   195
         Left            =   180
         TabIndex        =   50
         Tag             =   "1539"
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "被访问日志"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   49
         Tag             =   "1159"
         Top             =   3390
         Width           =   945
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1920
         TabIndex        =   48
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "扩展数据处理"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   46
         Tag             =   "1571"
         Top             =   1200
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frm_096"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'// Node information
'// 文件名：  Frm_096.frm
'// 用途：    异步通信节点
'// 作者:     Tony Sun
'// 创建日期：2005-03-15
'// 修改日期：2005-03-16
'// 文件描述：异步通信节点
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub cb_ExtDataType_Click()
    f_DataChanged = True
    
    txtPrefix.Enabled = False
    Cb_var_result.Enabled = False
    Select Case cb_ExtDataType.ListIndex
    Case 0
    Case 1
        txtPrefix.Enabled = True
    Case 2
        Cb_var_result.Enabled = True
    Case 3
    Case 4
    Case 5
    End Select
    
End Sub

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_rcvvar_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_usevar_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Cb_Var_Result_Click()
    f_DataChanged = True
End Sub

' Sun added 2012-11-23
'
Private Sub chkCarryOnPlay_Click()
    f_DataChanged = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Mike added this event @ 2008-1-31
Private Sub cmdNodeTag_Click()
    frmNodeTagEdit.iNodeID = CInt(T_n_id)
    frmNodeTagEdit.byNodeNo = CByte(T_n_no.Text)
    frmNodeTagEdit.Show vbModal
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    
Dim i As Integer

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 通信超时(秒)
        If Len(Trim(Cb_timeout)) < 1 Or Trim(Cb_timeout) = "" Then
            Node96_Data1.timeout = 0
        Else
            If CLng(Trim(Cb_timeout)) > 180 Then
                Message ("E123")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node96_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        ' 扩展数据处理
        If cb_ExtDataType.ListIndex < 0 Then
           Node96_Data1.extdata = 0
        Else
           Node96_Data1.extdata = CByte(cb_ExtDataType.ItemData(cb_ExtDataType.ListIndex))
        End If
        
        ' 分隔符
        If Trim(txtSeperator) = "" Then
            Node96_Data1.seperator = 0
        Else
            Node96_Data1.seperator = Asc(Left(txtSeperator, 1))
        End If
        
        ' 扩展数据记录
        If Cb_var_result.ListIndex <= 0 Then
            Node96_Data1.extvar = 0
        Else
            Node96_Data1.extvar = CByte(Cb_var_result.ItemData(Cb_var_result.ListIndex))
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
           Node96_Data1.log = 0
        Else
           Node96_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        ' 扩展数据记录
        If Cb_var_result.ListIndex <= 0 Then
            Node96_Data1.extvar = 0
        Else
            Node96_Data1.extvar = CByte(Cb_var_result.ItemData(Cb_var_result.ListIndex))
        End If
        
        Node96_Data1.reserved1 = 0
        Node96_Data1.reserved2(0) = 0
        
        ' 命令代码
        Node96_Data2.command = CInt(txtCommand)
        
        ' 文件名前缀
        Call StringToByteArray(txtPrefix, Node96_Data2.fileprefix, 2)
    
        ' 收发变量
        For i = 0 To 9
            If Cb_usevar(i).ListIndex <= 0 Then
                Node96_Data2.var_send(i) = 0
            Else
                Node96_Data2.var_send(i) = CByte(Cb_usevar(i).ItemData(Cb_usevar(i).ListIndex))
            End If
            
            If Cb_rcvvar(i).ListIndex <= 0 Then
                Node96_Data2.var_receive(i) = 0
            Else
                Node96_Data2.var_receive(i) = CByte(Cb_rcvvar(i).ItemData(Cb_rcvvar(i).ListIndex))
            End If
        Next
    
        ' 等待音乐
        If Trim(T_vox_op) = "" Then
            Node96_Data2.vox_wt = 0
        Else
            If CLng(Trim(T_vox_op)) > 32767 Then
                Message ("E097")
                T_vox_op.SetFocus
                Exit Sub
            Else
                Node96_Data2.vox_wt = CInt(Trim(T_vox_op))
            End If
        End If
    
        ' 保留
        Node96_Data2.reserved1(0) = 0
        Node96_Data2.reserved2(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node96_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node96_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
        
        ' 子节点
        If Trim(T_nd_child) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                lv_nNewNode = CInt(Trim(T_nd_child.Text))
            End If
        End If
        
        '' Sun added 2007-03-25
        If Node96_Data2.nd_child <> lv_nNewNode Then
            Node96_Data2.nd_child = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
        End If
        
        ' 无应答转节点
        If Trim(txtNdNoAnswer) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(txtNdNoAnswer) > 32767 Or Val(txtNdNoAnswer) < 256) And Val(txtNdNoAnswer) <> 0 Then
                Message ("E121")
                txtNdNoAnswer.SetFocus
                Exit Sub
            Else
                lv_nNewNode = CInt(Trim(txtNdNoAnswer))
            End If
        End If
    
        '' Sun added 2007-03-25
        If Node96_Data2.nd_timeout <> lv_nNewNode Then
            Node96_Data2.nd_timeout = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 255)
        End If
    
        '' Sun added 2012-11-23
        '' 是否收到通信响应是否继续异步放音
        If Node96_Data1.carryonasynplay <> chkCarryOnPlay.value Then
            Node96_Data1.carryonasynplay = chkCarryOnPlay.value
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

Private Sub cmdShowNodeList_Click(Index As Integer)
    Select Case Index
    Case 0
        Set gSystem.crlCurItem = T_nd_parent
    Case 1
        Set gSystem.crlCurItem = T_nd_child
    Case 2
        Set gSystem.crlCurItem = txtNdNoAnswer
    End Select
    frmNodeList.Show vbModal
End Sub

Private Sub cmdShowRes_Click()
    gSystem.intCurStep = 1
    Set gSystem.crlCurItem = T_vox_op
    frmResourceList.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
  
'初始化
' 通信超时(秒)
    With Cb_timeout
        .Clear
        For i = 0 To 120 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
        
' 扩展数据处理
    With cb_ExtDataType
        .Clear
        For i = 0 To 5
            .AddItem Trim(Str(i)) & " - " & LoadNationalResString(1576 + i)
            .ItemData(.ListCount - 1) = i
        Next
    End With

' 被访日志
    With Cb_log
        .Clear
        .AddItem LoadNationalResString(1178)
        .ItemData(.ListCount - 1) = 0
        For i = 1 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1179)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
' 扩展数据记录
    RefreshVariablesList Cb_var_result
  
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo

' 命令代码
    txtCommand = Node96_Data2.command

' 分隔符
    If Node96_Data1.seperator > Asc(" ") Then
        txtSeperator = Chr$(Node96_Data1.seperator)
    Else
        txtSeperator = ""
    End If
    
' 文件名前缀
    If Node96_Data2.fileprefix(0) > Asc(" ") Then
        txtPrefix = Chr$(Node96_Data2.fileprefix(0))
        If Node96_Data2.fileprefix(1) > Asc(" ") Then
            txtPrefix = txtPrefix & Chr$(Node96_Data2.fileprefix(1))
        End If
    Else
        txtPrefix = ""
    End If
    
' 收发变量
    For i = 0 To 9
        Cb_usevar(i).Clear
        RefreshVariablesList Cb_usevar(i)
        Cb_usevar(i).ListIndex = SearchItemDataIndex(Cb_usevar(i), CLng(Node96_Data2.var_send(i)), 0)
    
        Cb_rcvvar(i).Clear
        RefreshVariablesList Cb_rcvvar(i)
        Cb_rcvvar(i).ListIndex = SearchItemDataIndex(Cb_rcvvar(i), CLng(Node96_Data2.var_receive(i)), 0)
    Next
    
' 等待音乐
    T_vox_op = Node96_Data2.vox_wt
    
    '' Sun added 2012-11-23
    '' 是否收到通信响应是否继续异步放音
    chkCarryOnPlay.value = Node96_Data1.carryonasynplay
    
    T_nd_parent.Text = Node96_Data2.nd_parent
    T_nd_child.Text = Node96_Data2.nd_child
    txtNdNoAnswer.Text = Node96_Data2.nd_timeout
    
    cb_ExtDataType.ListIndex = SearchItemDataIndex(cb_ExtDataType, CLng(Node96_Data1.extdata), 0)
    Cb_timeout.Text = Node96_Data1.timeout
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node96_Data1.log), 0)
    Cb_var_result.ListIndex = SearchItemDataIndex(Cb_var_result, CLng(Node96_Data1.extvar), 0)
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    
    Call cb_ExtDataType_Click
    
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

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
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

Private Sub tabVars_Click()
    If tabVars.SelectedItem.Index = 2 Then
        picVars(0).Visible = False
        picVars(1).Visible = True
    Else
        picVars(0).Visible = True
        picVars(1).Visible = False
    End If
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtCommand_Change()
    f_DataChanged = True
End Sub

Private Sub txtCommand_GotFocus()
    txtCommand.SelStart = 0
    txtCommand.SelLength = Len(txtCommand)
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
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

Private Sub txtPrefix_Change()
    f_DataChanged = True
End Sub

Private Sub txtPrefix_GotFocus()
    txtPrefix.SelStart = 0
    txtPrefix.SelLength = Len(txtPrefix)
End Sub

Private Sub txtSeperator_Change()
    f_DataChanged = True
End Sub

Private Sub txtSeperator_GotFocus()
    txtSeperator.SelStart = 0
    txtSeperator.SelLength = Len(txtSeperator)
End Sub
