VERSION 5.00
Begin VB.Form frm_063 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "增强转接座席组  "
   ClientHeight    =   3660
   ClientLeft      =   2295
   ClientTop       =   2670
   ClientWidth     =   7365
   Icon            =   "frm_063.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1533"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   2280
      Picture         =   "frm_063.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "1948"
      Top             =   3261
      Width           =   333
   End
   Begin VB.Frame Frame12 
      Caption         =   "描述"
      Height          =   1185
      Left            =   2730
      TabIndex        =   33
      Tag             =   "1104"
      Top             =   2400
      Width           =   4545
      Begin VB.TextBox Txt_Description 
         Height          =   945
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   180
         Width           =   4425
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "座席组转接成功提示语音"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   30
      TabIndex        =   29
      Tag             =   "1324"
      Top             =   2580
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   30
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   0
      TabIndex        =   20
      Tag             =   "1136"
      Top             =   0
      Width           =   2415
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
         Left            =   1050
         TabIndex        =   32
         Top             =   1170
         Width           =   1335
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
         Left            =   1050
         TabIndex        =   31
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1530
         Width           =   1335
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
         Left            =   1050
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox T_n_id 
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
         Left            =   1050
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox T_n_no 
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
         Left            =   1050
         TabIndex        =   21
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "座席组ID"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Tag             =   "1330"
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "编号"
         Height          =   225
         Left            =   150
         TabIndex        =   27
         Tag             =   "1143"
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "节点ID"
         Height          =   225
         Left            =   150
         TabIndex        =   26
         Tag             =   "1137"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Tag             =   "1159"
         Top             =   1620
         Width           =   945
      End
      Begin VB.Label Label11 
         Caption         =   "播放次数"
         Height          =   225
         Left            =   150
         TabIndex        =   24
         Tag             =   "1304"
         Top             =   1260
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "父节点ID"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   2490
      TabIndex        =   18
      Tag             =   "1169"
      Top             =   0
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   19
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "没有上班转节点ID"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   2490
      TabIndex        =   16
      Tag             =   "1325"
      Top             =   600
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   17
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "座席组忙转节点ID"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   2490
      TabIndex        =   14
      Tag             =   "1331"
      Top             =   1200
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   15
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "转接成功转节点ID"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   2490
      TabIndex        =   12
      Tag             =   "1327"
      Top             =   1800
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   13
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "转接提示语音"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   30
      TabIndex        =   10
      Tag             =   "1323"
      Top             =   1950
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   11
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "操作提示语音"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   4890
      TabIndex        =   8
      Tag             =   "1266"
      Top             =   0
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   9
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "等待循环播放语音"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   4920
      TabIndex        =   6
      Tag             =   "1328"
      Top             =   1200
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "没有上班提示音"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   4920
      TabIndex        =   4
      Tag             =   "1307"
      Top             =   600
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "座席组忙提示音"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   4920
      TabIndex        =   2
      Tag             =   "1332"
      Top             =   1800
      Width           =   2355
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
         Height          =   315
         Left            =   1230
         TabIndex        =   3
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   1110
      TabIndex        =   1
      Tag             =   "1144"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Tag             =   "1007"
      Top             =   3240
      Width           =   1035
   End
End
Attribute VB_Name = "frm_063"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_063.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：转接座席组
'//////////////////////////////////////////////////////////////////
Option Explicit
Private Sub Cb_log_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

Private Sub Cb_looptimes_GotFocus()
    Cb_looptimes.SelStart = 0
    Cb_looptimes.SelLength = Len(Cb_looptimes)
End Sub

Private Sub Cb_looptimes_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

'Mike added this event @ 2008-1-31
Private Sub cmdNodeTag_Click()
    frmNodeTagEdit.iNodeID = CInt(T_n_id)
    frmNodeTagEdit.byNodeNo = CByte(T_n_no.Text)
    frmNodeTagEdit.Show vbModal
End Sub

Public Sub Command1_Click()
'保留
Node63_Data1.reserved1(0) = 0
   Unload Me
End Sub

Private Sub Command2_Click()
If Message("Q005") = vbYes Then
   Unload Me
Else
   Exit Sub
End If
End Sub
Private Sub Cb_log_GotFocus()
    Cb_log.SelStart = 0
    Cb_log.SelLength = Len(Cb_log)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_Load()

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
  LoadResStrings Me
End Sub

Private Sub T_groupid_GotFocus()
    T_groupid.SelStart = 0
    T_groupid.SelLength = Len(T_groupid)
End Sub

Private Sub T_groupid_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub
Private Sub T_looptimes_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

Private Sub T_nd_busy_GotFocus()
    T_nd_busy.SelStart = 0
    T_nd_busy.SelLength = Len(T_nd_busy)
End Sub
Private Sub T_nd_busy_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

Private Sub T_nd_nobody_GotFocus()
    T_nd_nobody.SelStart = 0
    T_nd_nobody.SelLength = Len(T_nd_nobody)
End Sub
Private Sub T_nd_nobody_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

Private Sub T_nd_ok_GotFocus()
    T_nd_ok.SelStart = 0
    T_nd_ok.SelLength = Len(T_nd_ok)
End Sub
Private Sub T_nd_ok_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

Private Sub T_nd_parent_GotFocus()
    T_nd_parent.SelStart = 0
    T_nd_parent.SelLength = Len(T_nd_parent)
End Sub
Private Sub T_nd_parent_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub
Private Sub T_vox_busy_GotFocus()
    T_vox_busy.SelStart = 0
    T_vox_busy.SelLength = Len(T_vox_busy)
End Sub
Private Sub T_vox_busy_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub

Private Sub T_vox_nobody_GotFocus()
    T_vox_nobody.SelStart = 0
    T_vox_nobody.SelLength = Len(T_vox_nobody)
End Sub
Private Sub T_vox_nobody_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub
Private Sub T_vox_ok_GotFocus()
    T_vox_ok.SelStart = 0
    T_vox_ok.SelLength = Len(T_vox_ok)
End Sub
Private Sub T_vox_ok_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
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
Private Sub T_vox_sw_GotFocus()
    T_vox_sw.SelStart = 0
    T_vox_sw.SelLength = Len(T_vox_sw)
End Sub
Private Sub T_vox_sw_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub
Private Sub T_vox_wt_GotFocus()
    T_vox_wt.SelStart = 0
    T_vox_wt.SelLength = Len(T_vox_wt)
End Sub
Private Sub T_vox_wt_KeyPress(KeyAscii As Integer)
        KeyPress KeyAscii
End Sub
