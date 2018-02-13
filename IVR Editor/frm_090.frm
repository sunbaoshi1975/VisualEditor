VERSION 5.00
Begin VB.Form frm_090 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "呼叫外线号码"
   ClientHeight    =   5190
   ClientLeft      =   4050
   ClientTop       =   2790
   ClientWidth     =   7725
   Icon            =   "frm_090.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1337"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   7332
      Picture         =   "frm_090.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "1948"
      Top             =   4665
      Width           =   333
   End
   Begin VB.Frame Frame3 
      Caption         =   "描述"
      Height          =   1035
      Index           =   1
      Left            =   4140
      TabIndex        =   47
      Tag             =   "1104"
      Top             =   3480
      Width           =   3525
      Begin VB.TextBox Txt_Description 
         Height          =   615
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   270
         Width           =   3345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   3345
      Left            =   4140
      TabIndex        =   40
      Tag             =   "1164"
      Top             =   60
      Width           =   3525
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   3030
         Picture         =   "frm_090.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "1146"
         Top             =   1170
         Width           =   315
      End
      Begin VB.TextBox txtSetANI 
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
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   13
         Top             =   1110
         Width           =   885
      End
      Begin VB.TextBox txtDialNo 
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
         Left            =   1680
         MaxLength       =   32
         TabIndex        =   12
         Top             =   660
         Width           =   1635
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
         Left            =   1680
         MaxLength       =   14
         TabIndex        =   11
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox T_com_iid 
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
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   15
         ToolTipText     =   "输入COM资源ID"
         Top             =   1560
         Width           =   885
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
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   17
         Top             =   1980
         Width           =   885
      End
      Begin VB.TextBox T_nd_succeed 
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
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2400
         Width           =   885
      End
      Begin VB.TextBox T_nd_fail 
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
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2820
         Width           =   885
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   3030
         Picture         =   "frm_090.frx":203E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "1146"
         Top             =   1590
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3030
         Picture         =   "frm_090.frx":2140
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "1145"
         Top             =   2010
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3030
         Picture         =   "frm_090.frx":24CA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "1145"
         Top             =   2430
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   3030
         Picture         =   "frm_090.frx":2854
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "1145"
         Top             =   2850
         Width           =   315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "设定主叫"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   49
         Tag             =   "2138"
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "呼叫号码"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   46
         Tag             =   "1343"
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "拨号前缀"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   45
         Tag             =   "1342"
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Lbl_com_iid 
         AutoSize        =   -1  'True
         Caption         =   "COM接口ID"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Tag             =   "1168"
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label Lbl_nd_parent 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Tag             =   "1169"
         Top             =   2070
         Width           =   705
      End
      Begin VB.Label Lbl_nd_succeed 
         AutoSize        =   -1  'True
         Caption         =   "成功转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Tag             =   "1170"
         Top             =   2490
         Width           =   1065
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "失败转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Tag             =   "1171"
         Top             =   2910
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   5055
      Index           =   0
      Left            =   60
      TabIndex        =   26
      Tag             =   "1136"
      Top             =   60
      Width           =   4005
      Begin VB.CheckBox chkExplictOffhook 
         Caption         =   "强制摘机"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Tag             =   "2124"
         Top             =   4620
         Width           =   1395
      End
      Begin VB.CheckBox chkResultNotify 
         Caption         =   "外拨结果通知"
         Height          =   195
         Left            =   2160
         TabIndex        =   10
         Tag             =   "1345"
         Top             =   4620
         Width           =   1695
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4140
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3720
         Width           =   1965
      End
      Begin VB.ComboBox Cb_ExtDelay 
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
         TabIndex        =   5
         ToolTipText     =   "0 - 255 秒"
         Top             =   2820
         Width           =   1965
      End
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "请选择"
         Top             =   1980
         Width           =   1965
      End
      Begin VB.ComboBox CB_Connect 
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
         ItemData        =   "frm_090.frx":2BDE
         Left            =   1860
         List            =   "frm_090.frx":2BE0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1560
         Width           =   1965
      End
      Begin VB.ComboBox cb_SwitchType 
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
         ItemData        =   "frm_090.frx":2BE2
         Left            =   1860
         List            =   "frm_090.frx":2BE4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1140
         Width           =   1965
      End
      Begin VB.ComboBox CB_calltype 
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
         ItemData        =   "frm_090.frx":2BE6
         Left            =   1860
         List            =   "frm_090.frx":2BE8
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
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
         Left            =   1860
         TabIndex        =   4
         ToolTipText     =   "0 - 255 秒"
         Top             =   2400
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
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   29
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
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
         TabIndex        =   6
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   3270
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "结果记录变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Tag             =   "1344"
         Top             =   4200
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Tag             =   "1283"
         Top             =   3780
         Width           =   885
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "拨分机前延时(秒)"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   37
         Tag             =   "1341"
         Top             =   2910
         Width           =   1350
      End
      Begin VB.Label Lbl_trytime 
         AutoSize        =   -1  'True
         Caption         =   "最大尝试次数"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Tag             =   "1158"
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "拨号方式"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Tag             =   "1302"
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "号码类型"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   34
         Tag             =   "1338"
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "外拨超时(秒)"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   33
         Tag             =   "1340"
         Top             =   2490
         Width           =   990
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1950
         TabIndex        =   31
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_log 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Left            =   180
         TabIndex        =   30
         Tag             =   "1245"
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "接通判别"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   28
         Tag             =   "1339"
         Top             =   1620
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   5940
      TabIndex        =   25
      Tag             =   "1144"
      Top             =   4650
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   4170
      TabIndex        =   24
      Tag             =   "1007"
      Top             =   4650
      Width           =   1035
   End
End
Attribute VB_Name = "frm_090"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'// Node information
'// 文件名：  Frm_090.frm
'// 用途：    呼叫外线号码节点
'// 作者:     Scott
'// 创建日期：2001-09-13
'// 修改日期：2005-3-12
'// 文件描述：呼叫外线号码
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub CB_calltype_Click()
    f_DataChanged = True
    
    txtDialNo.Enabled = False
    Cb_usevar.Enabled = False
    T_com_iid.Enabled = False
    cmdShowRes(0).Enabled = False
    Select Case CB_calltype.ListIndex
    Case 0
        txtDialNo.Enabled = True
    Case 1
        Cb_usevar.Enabled = True
    Case 2
    Case 3
        T_com_iid.Enabled = True
        cmdShowRes(0).Enabled = True
    End Select
End Sub

Private Sub CB_Connect_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_ExtDelay_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_ExtDelay_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_ExtDelay_GotFocus()
    Cb_ExtDelay.SelStart = 0
    Cb_ExtDelay.SelLength = Len(Cb_ExtDelay)
End Sub

Private Sub Cb_ExtDelay_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_maxtrytime_Click()
    f_DataChanged = True
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
End Sub

Private Sub Cb_Var_Result_Click()
    f_DataChanged = True
End Sub

Private Sub chkExplictOffhook_Click()
    f_DataChanged = True
End Sub

Private Sub chkResultNotify_Click()
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

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 号码类型
        If CB_calltype.ListIndex < 0 Then
           Node90_Data1.numbertype = 0
        Else
           Node90_Data1.numbertype = CByte(CB_calltype.ItemData(CB_calltype.ListIndex))
        End If
        
        ' 拨号方式
        If cb_SwitchType.ListIndex < 0 Then
           Node90_Data1.dialtype = 0
        Else
           Node90_Data1.dialtype = CByte(cb_SwitchType.ItemData(cb_SwitchType.ListIndex))
        End If
        
        ' 接通判别
        If CB_Connect.ListIndex < 0 Then
           Node90_Data1.connecttype = 0
        Else
           Node90_Data1.connecttype = CByte(CB_Connect.ItemData(CB_Connect.ListIndex))
        End If
        
        ' 最大尝试次数
        If Cb_maxtrytime.ListIndex < 0 Then
           Node90_Data1.trytimes = 3
        Else
           Node90_Data1.trytimes = CByte(Cb_maxtrytime.ItemData(Cb_maxtrytime.ListIndex))
        End If
        
        ' 外拨超时(秒)
        If Len(Trim(Cb_timeout)) < 1 Or Trim(Cb_timeout) = "" Then
            Node90_Data1.timeout = 0
        Else
            If CLng(Trim(Cb_timeout)) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node90_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        ' 拨分机前延时(秒)
        If Len(Trim(Cb_ExtDelay)) < 1 Or Trim(Cb_ExtDelay) = "" Then
            Node90_Data1.extdelay = 0
        Else
            If CLng(Trim(Cb_ExtDelay)) > 255 Then
                Message ("E122")
                Cb_ExtDelay.SetFocus
                Exit Sub
            Else
                Node90_Data1.extdelay = CByte(Val(Cb_ExtDelay.Text) Mod 256)
            End If
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
           Node90_Data1.log = 0
        Else
           Node90_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        If chkResultNotify = vbChecked Then
            Node90_Data1.resultinform = 1
        Else
            Node90_Data1.resultinform = 0
        End If
        
        '' Sun added 2012-01-17
        If chkExplictOffhook = vbChecked Then
            Node90_Data1.explictoffhook = 1
        Else
            Node90_Data1.explictoffhook = 0
        End If
        
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node90_Data1.usevar = 0
        Else
            Node90_Data1.usevar = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        ' 结果记录变量ID
        If Cb_var_result.ListIndex <= 0 Then
            Node90_Data1.resultvar = 0
        Else
            Node90_Data1.resultvar = CByte(Cb_var_result.ItemData(Cb_var_result.ListIndex))
        End If
        
        ' 拨号前缀
        Call StringToByteArray(txtPreDial, Node90_Data2.predial, 14)
        
        ' 呼叫号码
        Call StringToByteArray(txtDialNo, Node90_Data2.phoneno, 32)
        
        Node90_Data1.reserved1(0) = 0
    
        ' 保留
        Node90_Data2.reserved1(0) = 0
        
        ' COM接口ID
        If Trim(T_com_iid.Text) = "" Then
            Node90_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E040")
                T_com_iid.SetFocus
                Exit Sub
            Else
                Node90_Data2.com_iid = CInt(Trim(T_com_iid.Text))
            End If
        End If
        
        '' Sun added 2012-06-26
        ' 显式设置主叫号码（资源ID）
        If Trim(txtSetANI.Text) = "" Then
            Node90_Data2.setANI = 0
        Else
            If CLng(Trim(txtSetANI)) > 32767 Then
                Message ("E150")
                txtSetANI.SetFocus
                Exit Sub
            Else
                Node90_Data2.setANI = CInt(Trim(txtSetANI.Text))
            End If
        End If
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node90_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node90_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
        
        ' Succeed node id
        If Trim(T_nd_succeed) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_succeed) > 32767 Or Val(T_nd_succeed) < 256) And Val(T_nd_succeed) <> 0 Then
                Message ("E041")
                T_nd_succeed.SetFocus
                Exit Sub
            Else
                lv_nNewNode = CInt(Trim(T_nd_succeed.Text))
            End If
        End If
        
        '' Sun added 2007-03-25
        If Node90_Data2.nd_succ <> lv_nNewNode Then
            Node90_Data2.nd_succ = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 1)
        End If
        
        ' failed node id
        If Trim(T_nd_fail) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_fail) > 32767 Or Val(T_nd_fail) < 256) And Val(T_nd_fail) <> 0 Then
                Message ("E042")
                T_nd_fail.SetFocus
                Exit Sub
            Else
                lv_nNewNode = CInt(Trim(T_nd_fail.Text))
            End If
        End If
           
        '' Sun added 2007-03-25
        If Node90_Data2.nd_fail <> lv_nNewNode Then
            Node90_Data2.nd_fail = lv_nNewNode
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
    
    '' Sun added 2007-03-25
    If Node90_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
        Message "M144"
    End If
    
    Unload Me

End Sub

Private Sub cmdShowNodeList_Click(Index As Integer)
    
    Select Case Index
    Case 0
        Set gSystem.crlCurItem = T_nd_parent
    Case 1
        Set gSystem.crlCurItem = T_nd_succeed
    Case 2
        Set gSystem.crlCurItem = T_nd_fail
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click(Index As Integer)
    Select Case Index
    Case 0
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_iid
    Case 1
        gSystem.intCurStep = 4
        Set gSystem.crlCurItem = txtSetANI
    End Select
    frmResourceList.Show vbModal
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
  
'初始化
' 号码类型
    With CB_calltype
        .Clear
        For i = 0 To 3
            .AddItem Trim(Str(i)) & " - " & LoadNationalResString(1556 + i)
            .ItemData(.ListCount - 1) = i
        Next
    End With
        
' 拨号方式
    With cb_SwitchType
        .Clear
        .AddItem "0 - " & LoadNationalResString(1560)
        .ItemData(.ListCount - 1) = 0
        .AddItem "1 - " & LoadNationalResString(1561)
        .ItemData(.ListCount - 1) = 1
    End With

' 接通判别
    With CB_Connect
        .Clear
        .AddItem "0 - " & LoadNationalResString(1562)
        .ItemData(.ListCount - 1) = 0
        .AddItem "1 - " & LoadNationalResString(1563)
        .ItemData(.ListCount - 1) = 1
    End With

' 最大尝试次数
    With Cb_maxtrytime
        .Clear
        For i = 0 To 9
            .AddItem Trim(Str(i)) & LoadNationalResString(1180)
            .ItemData(.ListCount - 1) = i
        Next
    End With

' 外拨超时(秒)
    With Cb_timeout
        .Clear
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
' 拨分机前延时(秒)
    With Cb_ExtDelay
        .Clear
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
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
    
' 使用变量ID
    RefreshVariablesList Cb_usevar

' 结果记录变量ID
    RefreshVariablesList Cb_var_result
  
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
' 拨号前缀
    txtPreDial = ByteArrayToString(Node90_Data2.predial, 14)
    
' 呼叫号码
    txtDialNo = ByteArrayToString(Node90_Data2.phoneno, 32)
    
    '' Sun added 2012-06-26
' 显式设置主叫号码（资源ID）
    txtSetANI.Text = Node90_Data2.setANI
    
    T_com_iid.Text = Node90_Data2.com_iid
    T_nd_parent.Text = Node90_Data2.nd_parent
    T_nd_succeed.Text = Node90_Data2.nd_succ
    T_nd_fail.Text = Node90_Data2.nd_fail
    
    CB_calltype.ListIndex = SearchItemDataIndex(CB_calltype, CLng(Node90_Data1.numbertype), 0)
    cb_SwitchType.ListIndex = SearchItemDataIndex(cb_SwitchType, CLng(Node90_Data1.dialtype), 0)
    CB_Connect.ListIndex = SearchItemDataIndex(CB_Connect, CLng(Node90_Data1.connecttype), 0)
    Cb_maxtrytime.ListIndex = SearchItemDataIndex(Cb_maxtrytime, CLng(Node90_Data1.trytimes), 3)
    Cb_timeout.Text = Node90_Data1.timeout
    Cb_ExtDelay.Text = Node90_Data1.extdelay
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node90_Data1.log), 0)
    
    chkResultNotify = IIf(Node90_Data1.resultinform > 0, vbChecked, vbUnchecked)
    
    '' Sun adde 2012-01-17
    chkExplictOffhook = IIf(Node90_Data1.explictoffhook > 0, vbChecked, vbUnchecked)
    
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node90_Data1.usevar), 0)
    Cb_var_result.ListIndex = SearchItemDataIndex(Cb_var_result, CLng(Node90_Data1.resultvar), 0)
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    
    Call CB_calltype_Click
    
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me

End Sub

Private Sub T_com_iid_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_com_iid
    
End Sub

Private Sub T_com_iid_GotFocus()
    T_com_iid.SelStart = 0
    T_com_iid.SelLength = Len(T_com_iid)
End Sub

Private Sub T_com_iid_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_nd_fail_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_fail_GotFocus()
    T_nd_fail.SelStart = 0
    T_nd_fail.SelLength = Len(T_nd_fail)
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

Private Sub T_nd_succeed_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_succeed_GotFocus()
    T_nd_succeed.SelStart = 0
    T_nd_succeed.SelLength = Len(T_nd_succeed)
End Sub

Private Sub T_nd_succeed_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtDialNo_Change()
    f_DataChanged = True
End Sub

Private Sub txtDialNo_GotFocus()
    txtDialNo.SelStart = 0
    txtDialNo.SelLength = Len(txtDialNo)
End Sub

Private Sub txtPreDial_Change()
    f_DataChanged = True
End Sub

Private Sub txtPreDial_GotFocus()
    txtPreDial.SelStart = 0
    txtPreDial.SelLength = Len(txtPreDial)
End Sub

Private Sub txtSetANI_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtSetANI
    
End Sub

Private Sub txtSetANI_GotFocus()
    txtSetANI.SelStart = 0
    txtSetANI.SelLength = Len(txtSetANI)
End Sub

Private Sub txtSetANI_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
