VERSION 5.00
Begin VB.Form SystemDefault 
   Caption         =   "系统参数设置"
   ClientHeight    =   2655
   ClientLeft      =   3810
   ClientTop       =   3345
   ClientWidth     =   2670
   Icon            =   "SystemDefault.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   2670
   Tag             =   "1469"
   Begin VB.TextBox Txt_uservars 
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
      Left            =   1530
      TabIndex        =   12
      ToolTipText     =   "0-255"
      Top             =   1890
      Width           =   795
   End
   Begin VB.TextBox Txt_ndparent 
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
      Left            =   1530
      TabIndex        =   6
      Top             =   1170
      Width           =   1035
   End
   Begin VB.TextBox Txt_ndroot 
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
      Left            =   1530
      TabIndex        =   5
      Top             =   1530
      Width           =   1035
   End
   Begin VB.ComboBox Cmb_repeat 
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
      Left            =   1530
      TabIndex        =   4
      Top             =   60
      Width           =   1035
   End
   Begin VB.ComboBox Cmb_return 
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
      Left            =   1530
      TabIndex        =   3
      Top             =   420
      Width           =   1035
   End
   Begin VB.ComboBox Cmb_root 
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
      Left            =   1530
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Tag             =   "1007"
      Top             =   2280
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Tag             =   "1144"
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Label Lbl_uservars 
      AutoSize        =   -1  'True
      Caption         =   "用户定义变量数"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Tag             =   "1474"
      Top             =   1950
      Width           =   1260
   End
   Begin VB.Label Lbl_nd_parent 
      AutoSize        =   -1  'True
      Caption         =   "父节点ID"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Tag             =   "1169"
      Top             =   1230
      Width           =   705
   End
   Begin VB.Label Lbl_key_root 
      AutoSize        =   -1  'True
      Caption         =   "回到主菜单"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Tag             =   "1472"
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Lbl_key_return 
      AutoSize        =   -1  'True
      Caption         =   "回上一级节点按键"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Tag             =   "1471"
      Top             =   510
      Width           =   1440
   End
   Begin VB.Label Lbl_key_repeat 
      AutoSize        =   -1  'True
      Caption         =   "重复当前节点按键"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Tag             =   "1470"
      Top             =   150
      Width           =   1440
   End
   Begin VB.Label Lbl_nd_root 
      AutoSize        =   -1  'True
      Caption         =   "根节点ID"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Tag             =   "1473"
      Top             =   1590
      Width           =   705
   End
End
Attribute VB_Name = "SystemDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmb_repeat_KeyPress(KeyAscii As Integer)
       KeyPress 0
End Sub
Private Sub Cmb_return_KeyPress(KeyAscii As Integer)
       KeyPress 0
End Sub
Private Sub Cmb_root_KeyPress(KeyAscii As Integer)
       KeyPress 0
End Sub

Private Sub CommandExit_Click()
Unload Me
End Sub

Private Sub CommandSave_Click()
  If Cmb_repeat = "" Then
     Exit Sub
  Else
     If Cmb_return = "" Then
        Exit Sub
     Else
        If Cmb_root = "" Then
           Exit Sub
        Else
           If Txt_ndparent = "" Or CLng(Txt_ndparent) < 0 Or CLng(Txt_ndparent) > 32767 Then
              Message ("E035")
              Exit Sub
           Else
              If Txt_ndroot = "" Or CLng(Txt_ndroot) < 0 Or CLng(Txt_ndroot) > 32767 Then
                 Message ("E034")
                 Exit Sub
              Else
                 If Txt_uservars = "" Or CLng(Txt_uservars) < 0 Or CLng(Txt_uservars) > 256 Then
                    Message ("E036")
                    Exit Sub
                 Else
  'Cmb_repeat's data is written in EasyRS.ini file
                    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_repeat, Cmb_repeat, gSystem.strINI_File)
  'Cmb_return's data is written in EasyRS.ini file
                    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_return, Cmb_return, gSystem.strINI_File)
  'Cmb_root's data is written in EasyRS.ini file
                    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_root, Cmb_root, gSystem.strINI_File)
  'txt_ndparent's data is written in EasyRS.ini file
                    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Nd_parent, Txt_ndparent, gSystem.strINI_File)
  'Txt_ndroot's data is written in EasyRS.ini file
                    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Nd_root, Txt_ndroot, gSystem.strINI_File)
  'Txt_uservars's data is written in EasyRS.ini file
                    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Uservar, Txt_uservars, gSystem.strINI_File)
                    Default.key_repeat = Asc(Cmb_repeat)
                    Default.key_return = Asc(Cmb_return)
                    Default.key_root = Asc(Cmb_root)
                    Default.nd_parent = Txt_ndparent
                    Default.nd_root = Txt_ndroot
                    Defaultuservar.uservars = Txt_uservars
                    Unload Me
                 End If
              End If
           End If
        End If
     End If
  End If
End Sub

Private Sub Form_Load()
   '初始化
   With Cmb_repeat
     .AddItem "0"
     .AddItem "1"
     .AddItem "2"
     .AddItem "3"
     .AddItem "4"
     .AddItem "5"
     .AddItem "6"
     .AddItem "7"
     .AddItem "8"
     .AddItem "9"
     .AddItem "A"
     .AddItem "N"
     .AddItem "*"
     .AddItem "#"
   End With
   With Cmb_return
     .AddItem "0"
     .AddItem "1"
     .AddItem "2"
     .AddItem "3"
     .AddItem "4"
     .AddItem "5"
     .AddItem "6"
     .AddItem "7"
     .AddItem "8"
     .AddItem "9"
     .AddItem "A"
     .AddItem "N"
     .AddItem "*"
     .AddItem "#"
   End With
   With Cmb_root
     .AddItem "0"
     .AddItem "1"
     .AddItem "2"
     .AddItem "3"
     .AddItem "4"
     .AddItem "5"
     .AddItem "6"
     .AddItem "7"
     .AddItem "8"
     .AddItem "9"
     .AddItem "A"
     .AddItem "N"
     .AddItem "*"
     .AddItem "#"
  End With
  '重复上一节点按键
  Cmb_repeat = Chr(Default.key_repeat)
  '回到上一菜单按键
  Cmb_return = Chr(Default.key_return)
  '回到主菜单按键
  Cmb_root = Chr(Default.key_root)
  '父节点
  Txt_ndparent = Default.nd_parent
  '根节点
  Txt_ndroot = Default.nd_root
  '用户定义变量数
  Txt_uservars = Defaultuservar.uservars
    LoadResStrings Me
End Sub
Private Sub Txt_ndparent_GotFocus()
    Txt_ndparent.SelStart = 0
    Txt_ndparent.SelLength = Len(Txt_ndparent)
End Sub

Private Sub Txt_ndroot_KeyPress(KeyAscii As Integer)
    Txt_ndroot.SelStart = 0
    Txt_ndroot.SelLength = Len(Txt_ndroot)
End Sub

Private Sub Txt_uservars_Change()
    Txt_uservars.SelStart = 0
    Txt_uservars.SelLength = Len(Txt_uservars)
End Sub
