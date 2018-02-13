VERSION 5.00
Begin VB.Form frm_041 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "察看留言"
   ClientHeight    =   4575
   ClientLeft      =   3675
   ClientTop       =   2550
   ClientWidth     =   9600
   Icon            =   "frm_041.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1051"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   9192
      Picture         =   "frm_041.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "1948"
      Top             =   4041
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   795
      Left            =   60
      TabIndex        =   55
      Tag             =   "1104"
      Top             =   3630
      Width           =   3885
      Begin VB.TextBox Txt_Description 
         Height          =   465
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   3525
      Index           =   1
      Left            =   4020
      TabIndex        =   41
      Tag             =   "1164"
      Top             =   60
      Width           =   5505
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "请选择"
         Top             =   2160
         Width           =   1125
      End
      Begin VB.CheckBox chhCloseWhenCheck 
         Caption         =   "浏览后自动存档"
         Height          =   285
         Left            =   1470
         TabIndex        =   23
         Tag             =   "1672"
         Top             =   2610
         Width           =   2355
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   4950
         Picture         =   "frm_041.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "1145"
         Top             =   3030
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2310
         Picture         =   "frm_041.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "1145"
         Top             =   3030
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   4950
         Picture         =   "frm_041.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "1146"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2310
         Picture         =   "frm_041.frx":2752
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "1146"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   4950
         Picture         =   "frm_041.frx":2854
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2310
         Picture         =   "frm_041.frx":2956
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   4110
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "请选择"
         Top             =   2160
         Width           =   1125
      End
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   4110
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "请选择"
         Top             =   1770
         Width           =   1125
      End
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "请选择"
         Top             =   1770
         Width           =   1125
      End
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   4110
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "请选择"
         Top             =   1410
         Width           =   1125
      End
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "请选择"
         Top             =   1410
         Width           =   1125
      End
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   4110
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "请选择"
         Top             =   1050
         Width           =   1125
      End
      Begin VB.ComboBox Cb_Key_Op 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "请选择"
         Top             =   1050
         Width           =   1125
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
         Index           =   3
         Left            =   4110
         MaxLength       =   6
         TabIndex        =   13
         Top             =   600
         Width           =   825
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
         Index           =   2
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   11
         Top             =   600
         Width           =   825
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
         Index           =   1
         Left            =   4110
         MaxLength       =   6
         TabIndex        =   9
         Top             =   240
         Width           =   825
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
         Left            =   4110
         MaxLength       =   6
         TabIndex        =   26
         Top             =   3000
         Width           =   825
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
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   24
         Top             =   3000
         Width           =   825
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
         Index           =   0
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblKeyConvert 
         AutoSize        =   -1  'True
         Caption         =   "转换类型按键"
         Height          =   195
         Left            =   180
         TabIndex        =   56
         Tag             =   "1671"
         Top             =   2250
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "退出节点按键"
         Height          =   180
         Index           =   10
         Left            =   2850
         TabIndex        =   54
         Tag             =   "1293"
         Top             =   2250
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "删除留言按键"
         Height          =   180
         Index           =   9
         Left            =   2850
         TabIndex        =   53
         Tag             =   "1292"
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "重听留言按键"
         Height          =   180
         Index           =   8
         Left            =   180
         TabIndex        =   52
         Tag             =   "1291"
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "听末一条按键"
         Height          =   180
         Index           =   7
         Left            =   2850
         TabIndex        =   51
         Tag             =   "1290"
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "听下一条按键"
         Height          =   180
         Index           =   6
         Left            =   180
         TabIndex        =   50
         Tag             =   "1289"
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "听前一条按键"
         Height          =   180
         Index           =   5
         Left            =   2850
         TabIndex        =   49
         Tag             =   "1288"
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "听第一条按键"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   48
         Tag             =   "1287"
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "操作提示语音"
         Height          =   180
         Index           =   3
         Left            =   2850
         TabIndex        =   47
         Tag             =   "1266"
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "浏览提示语"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   46
         Tag             =   "1286"
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "留言报告后续"
         Height          =   180
         Index           =   1
         Left            =   2850
         TabIndex        =   45
         Tag             =   "1285"
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "子节点ID"
         Height          =   180
         Left            =   2850
         TabIndex        =   44
         Tag             =   "1252"
         Top             =   3060
         Width           =   720
      End
      Begin VB.Label Label20 
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Tag             =   "1169"
         Top             =   3060
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "留言报告前导"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   42
         Tag             =   "1284"
         Top             =   330
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3525
      Left            =   60
      TabIndex        =   31
      Tag             =   "1136"
      Top             =   60
      Width           =   3885
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
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "请选择"
         Top             =   2550
         Width           =   1965
      End
      Begin VB.ComboBox cmbVMSType 
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
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "请选择"
         Top             =   2910
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
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   33
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   1770
         TabIndex        =   0
         ToolTipText     =   "0 - 255 秒"
         Top             =   750
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
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1830
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
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "请选择"
         Top             =   1470
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
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "请选择"
         Top             =   1110
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
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2190
         Width           =   1965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "留言分类"
         Height          =   195
         Left            =   180
         TabIndex        =   58
         Tag             =   "1690"
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "察看留言类型"
         Height          =   195
         Left            =   180
         TabIndex        =   57
         Tag             =   "1670"
         Top             =   2970
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "节点超时(秒)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   40
         Tag             =   "1154"
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1860
         TabIndex        =   38
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_log 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Left            =   180
         TabIndex        =   37
         Tag             =   "1245"
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "放音清空标志"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   36
         Tag             =   "1244"
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "按键中断"
         Height          =   180
         Left            =   180
         TabIndex        =   35
         Tag             =   "1264"
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Tag             =   "1283"
         Top             =   2250
         Width           =   885
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   7290
      TabIndex        =   30
      Top             =   4020
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   5310
      TabIndex        =   29
      Tag             =   "1007"
      Top             =   4020
      Width           =   1035
   End
End
Attribute VB_Name = "frm_041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_041.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：察看流言
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_breakkey_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Key_Op_Click(Index As Integer)
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

Private Sub chhCloseWhenCheck_Click()
    f_DataChanged = True
End Sub

Private Sub cmbVMSClass_Click()
    f_DataChanged = True
End Sub

Private Sub cmbVMSType_Click()
    f_DataChanged = True
    
    If cmbVMSType.ItemData(cmbVMSType.ListIndex) = DEF_NODE041_VMSTYPE_NEW Then
        chhCloseWhenCheck.Enabled = True
    Else
        chhCloseWhenCheck.value = vbUnchecked
        chhCloseWhenCheck.Enabled = False
    End If
    
    If cmbVMSType.ItemData(cmbVMSType.ListIndex) = DEF_NODE041_VMSTYPE_CLOSED Then
        Cb_Key_Op(7).Enabled = False
        lblKeyConvert.Enabled = False
    Else
        Cb_Key_Op(7).Enabled = True
        lblKeyConvert.Enabled = True
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
        Set gSystem.crlCurItem = T_nd_child
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click(Index As Integer)

    gSystem.intCurStep = 1
    Set gSystem.crlCurItem = T_vox_play(Index)
    frmResourceList.Show vbModal

End Sub

Public Sub CommandSave_Click()
On Error Resume Next

Dim i As Integer
Dim j As Integer

    If f_DataChanged Then

        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node41_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node41_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If

        ' 放音清空
        If Cb_playclear.ListIndex = -1 Then
            Node41_Data1.playclear = 0
        Else
            Node41_Data1.playclear = CByte(Cb_playclear.ItemData(Cb_playclear.ListIndex))
        End If

        ' 按键中断
        If Cb_breakkey.ListIndex < 0 Then
            Node41_Data1.breakkey = 0
        Else
            Node41_Data1.breakkey = CByte(Cb_breakkey.ItemData(Cb_breakkey.ListIndex))
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node41_Data1.log = 0
        Else
            Node41_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        ' Sun added 2007-02-28
        ' 留言分类
        If cmbVMSClass.ListIndex <= 0 Then
            Node41_Data1.vmsclass = 0
        Else
            Node41_Data1.vmsclass = CByte(cmbVMSClass.ItemData(cmbVMSClass.ListIndex))
        End If
        
        '' Sun added 2006-02-06
        ' 察看留言类型
        If cmbVMSType.ListIndex < 0 Then
            Node41_Data1.vmstype = 0
        Else
            Node41_Data1.vmstype = CByte(cmbVMSType.ItemData(cmbVMSType.ListIndex))
        End If
        
        '' Sun added 2006-02-06
        ' 浏览后自动存档
        Node41_Data1.closewhencheck = chhCloseWhenCheck.value
        
        Node41_Data1.reserved1 = 0
        Node41_Data1.reserved2(0) = 0
                
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node41_Data1.var_agent = 0
        Else
            Node41_Data1.var_agent = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If

        ' 录音提示
        For i = T_vox_play.LBound To T_vox_play.UBound
            If Trim(T_vox_play(i)) = "" Then
                Node41_Data2.vox_play(i) = 0
            Else
                If CLng(Trim(T_vox_play(i))) > 32767 Then
                    Message ("E092")
                    T_vox_play(i).SetFocus
                    Exit Sub
                Else
                    Node41_Data2.vox_play(i) = CInt(Trim(T_vox_play(i)))
                End If
            End If
        Next

        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node41_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node41_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If

        ' 子节点
        If Trim(T_nd_child) = "" Then
            Node41_Data2.nd_child = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                Node41_Data2.nd_child = CInt(Trim(T_nd_child.Text))
            End If
        End If
    
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        '' Sun added 2006-02-06
        ' 按键设置正确性检查
        For i = 0 To 2
            For j = i + 1 To 3
                If Cb_Key_Op(i).ListIndex > 0 And Cb_Key_Op(j).ListIndex > 0 Then
                    If Cb_Key_Op(i).ListIndex = Cb_Key_Op(j).ListIndex Then
                        Message ("E125")
                        Cb_Key_Op(j).SetFocus
                        Exit Sub
                    End If
                End If
            Next j
        Next i
        For i = 4 To 6
            For j = i + 1 To 7
                If Cb_Key_Op(i).ListIndex > 0 And Cb_Key_Op(j).ListIndex > 0 Then
                    If Cb_Key_Op(i).ListIndex = Cb_Key_Op(j).ListIndex Then
                        Message ("E125")
                        Cb_Key_Op(j).SetFocus
                        Exit Sub
                    End If
                End If
            Next j
        Next i
        
        For i = Cb_Key_Op.LBound To Cb_Key_Op.UBound
            If Cb_Key_Op(i).ListIndex < 0 Then
                Node41_Data2.key_op(i) = 0
            Else
                Node41_Data2.key_op(i) = Cb_Key_Op(i).ItemData(Cb_Key_Op(i).ListIndex)
            End If
        Next

        Node41_Data2.reserved1(0) = 0
        Node41_Data2.reserved2(0) = 0
        
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

    ' 节点超时(秒)
    With Cb_timeout
        For i = 0 To 15
            .AddItem Trim(Str(i))
        Next
    End With

    ' 使用变量ID
    RefreshVariablesList Cb_usevar
    
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
    
    '' Sun added 2006-02-06
    ' 察看留言类型
    With cmbVMSType
        For i = DEF_NODE041_VMSTYPE_NEW To DEF_NODE041_VMSTYPE_DELETED
            .AddItem Trim(Str(i)) & " - " & LoadNationalResString(1673 + i)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
    For i = Cb_Key_Op.LBound To Cb_Key_Op.UBound
        With Cb_Key_Op(i)
            .AddItem "NULL"
            .ItemData(.ListCount - 1) = 0
            .AddItem "0"
            .ItemData(.ListCount - 1) = Asc("0")
            .AddItem "1"
            .ItemData(.ListCount - 1) = Asc("1")
            .AddItem "2"
            .ItemData(.ListCount - 1) = Asc("2")
            .AddItem "3"
            .ItemData(.ListCount - 1) = Asc("3")
            .AddItem "4"
            .ItemData(.ListCount - 1) = Asc("4")
            .AddItem "5"
            .ItemData(.ListCount - 1) = Asc("5")
            .AddItem "6"
            .ItemData(.ListCount - 1) = Asc("6")
            .AddItem "7"
            .ItemData(.ListCount - 1) = Asc("7")
            .AddItem "8"
            .ItemData(.ListCount - 1) = Asc("8")
            .AddItem "9"
            .ItemData(.ListCount - 1) = Asc("9")
            .AddItem "*"
            .ItemData(.ListCount - 1) = Asc("*")
            .AddItem "#"
            .ItemData(.ListCount - 1) = Asc("#")
        End With
    Next

    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    For i = T_vox_play.LBound To T_vox_play.UBound
        T_vox_play(i) = Node41_Data2.vox_play(i)
    Next
    T_nd_parent.Text = Node41_Data2.nd_parent
    T_nd_child.Text = Node41_Data2.nd_child
    
    Cb_timeout.Text = Node41_Data1.timeout
    Cb_playclear.ListIndex = SearchItemDataIndex(Cb_playclear, CLng(Node41_Data1.playclear), 0)
    Cb_breakkey.ListIndex = SearchItemDataIndex(Cb_breakkey, CLng(Node41_Data1.breakkey), 10)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node41_Data1.log), 0)
    
    ' Sun added 2007-02-28
    cmbVMSClass.ListIndex = SearchItemDataIndex(cmbVMSClass, CLng(Node41_Data1.vmsclass), 0)
 
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node41_Data1.var_agent), 0)
    
    For i = Cb_Key_Op.LBound To Cb_Key_Op.UBound
        Cb_Key_Op(i).ListIndex = SearchItemDataIndex(Cb_Key_Op(i), CLng(Node41_Data2.key_op(i)), i + 1)
    Next
    
    '' Sun added 2006-02-06
    cmbVMSType.ListIndex = SearchItemDataIndex(cmbVMSType, CLng(Node41_Data1.vmstype), 0)
    chhCloseWhenCheck.value = Node41_Data1.closewhencheck
        
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

Private Sub T_vox_play_Change(Index As Integer)
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_play(Index)

End Sub

Private Sub T_vox_play_GotFocus(Index As Integer)
    T_vox_play(Index).SelStart = 0
    T_vox_play(Index).SelLength = Len(T_vox_play(Index))

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_play(Index))
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_play_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
