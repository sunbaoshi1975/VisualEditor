VERSION 5.00
Begin VB.Form frm_007 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   4575
   ClientLeft      =   2925
   ClientTop       =   2430
   ClientWidth     =   7035
   Icon            =   "frm_007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1153"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6642
      Picture         =   "frm_007.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "1948"
      Top             =   4182
      Width           =   333
   End
   Begin VB.Frame Frame3 
      Caption         =   "描述"
      Height          =   1095
      Left            =   3900
      TabIndex        =   48
      Tag             =   "1104"
      Top             =   3000
      Width           =   3075
      Begin VB.TextBox Txt_Description 
         Height          =   735
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   270
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   3900
      TabIndex        =   37
      Tag             =   "1164"
      Top             =   60
      Width           =   3075
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   2640
         Picture         =   "frm_007.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "1145"
         Top             =   2430
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2640
         Picture         =   "frm_007.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "1145"
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2640
         Picture         =   "frm_007.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "1145"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   2640
         Picture         =   "frm_007.frx":29DA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "1146"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2640
         Picture         =   "frm_007.frx":2ADC
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "1145"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2640
         Picture         =   "frm_007.frx":2BDE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "1145"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2640
         Picture         =   "frm_007.frx":2CE0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox T_vox_userid 
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   11
         ToolTipText     =   "输入语音资源ID"
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox T_vox_tryagain 
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   15
         ToolTipText     =   "输入语音资源ID"
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox T_vox_password 
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   13
         ToolTipText     =   "输入语音资源ID"
         Top             =   600
         Width           =   795
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   17
         ToolTipText     =   "输入COM资源ID"
         Top             =   1320
         Width           =   795
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   19
         Top             =   1680
         Width           =   795
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2040
         Width           =   795
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   23
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label Lbl_vox_userid 
         Caption         =   "提示用户输入号码"
         Height          =   225
         Left            =   180
         TabIndex        =   46
         Tag             =   "1165"
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Lbl_vox_tryagain 
         Caption         =   "提示用户重新输入"
         Height          =   225
         Left            =   180
         TabIndex        =   43
         Tag             =   "1167"
         Top             =   1020
         Width           =   1545
      End
      Begin VB.Label Lbl_vox_password 
         Caption         =   "提示用户输入口令"
         Height          =   225
         Left            =   180
         TabIndex        =   42
         Tag             =   "1166"
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Lbl_com_iid 
         Caption         =   "COM接口ID"
         Height          =   225
         Left            =   180
         TabIndex        =   41
         Tag             =   "1168"
         Top             =   1350
         Width           =   1515
      End
      Begin VB.Label Lbl_nd_parent 
         Caption         =   "父节点ID"
         Height          =   225
         Left            =   180
         TabIndex        =   40
         Tag             =   "1169"
         Top             =   1740
         Width           =   1545
      End
      Begin VB.Label Lbl_nd_succeed 
         Caption         =   "成功转节点ID"
         Height          =   225
         Left            =   180
         TabIndex        =   39
         Tag             =   "1170"
         Top             =   2100
         Width           =   1545
      End
      Begin VB.Label Lbl_nd_fail 
         Caption         =   "失败转节点ID"
         Height          =   225
         Left            =   180
         TabIndex        =   38
         Tag             =   "1171"
         Top             =   2460
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   4455
      Left            =   60
      TabIndex        =   28
      Tag             =   "1136"
      Top             =   60
      Width           =   3795
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
         TabIndex        =   0
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
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Cb_var_userid 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3960
         Width           =   1965
      End
      Begin VB.ComboBox Cb_var_password 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3600
         Width           =   1965
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   1965
      End
      Begin VB.ComboBox Cb_var_trytime 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   2880
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2520
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "请选择"
         Top             =   2160
         Width           =   1965
      End
      Begin VB.ComboBox Cb_key_term 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "请选择"
         Top             =   1800
         Width           =   1965
      End
      Begin VB.ComboBox Cb_maxpassword 
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
         Left            =   1710
         TabIndex        =   4
         ToolTipText     =   "0 - 255 个按键"
         Top             =   1440
         Width           =   1965
      End
      Begin VB.ComboBox Cb_maxuserid 
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
         Left            =   1710
         TabIndex        =   3
         ToolTipText     =   "0 - 255 个按键"
         Top             =   1080
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
         Left            =   1710
         TabIndex        =   2
         ToolTipText     =   "0 - 255 秒"
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   2010
         TabIndex        =   50
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_var_userid 
         Caption         =   "用户号码记录"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Tag             =   "1163"
         Top             =   4020
         Width           =   1125
      End
      Begin VB.Label Lbl_maxuserid 
         Caption         =   "号码最大长度"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Tag             =   "1155"
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Lbl_timeout 
         Caption         =   "节点超时(秒)"
         Height          =   225
         Left            =   180
         TabIndex        =   36
         Tag             =   "1154"
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Lbl_maxpassword 
         Caption         =   "口令最大长度"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Tag             =   "1156"
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label Lbl_key_term 
         Caption         =   "输入终止按键"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Tag             =   "1157"
         Top             =   1860
         Width           =   1305
      End
      Begin VB.Label Lbl_maxtrytime 
         Caption         =   "最大尝试次数"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Tag             =   "1158"
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Lbl_log 
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Tag             =   "1159"
         Top             =   2580
         Width           =   915
      End
      Begin VB.Label Lbl_var_trytime 
         Caption         =   "验证次数记录"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Tag             =   "1160"
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label Lbl_var_result 
         Caption         =   "验证结果记录"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1161"
         Top             =   3300
         Width           =   1185
      End
      Begin VB.Label Lbl_var_password 
         Caption         =   "用户口令记录"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Tag             =   "1162"
         Top             =   3660
         Width           =   1125
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   5130
      TabIndex        =   27
      Tag             =   "1144"
      Top             =   4200
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   26
      Tag             =   "1007"
      Top             =   4200
      Width           =   1035
   End
End
Attribute VB_Name = "frm_007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_007.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：身份验证
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_key_term_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_maxpassword_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_maxpassword_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_maxpassword_GotFocus()
    Cb_maxpassword.SelStart = 0
    Cb_maxpassword.SelLength = Len(Cb_maxpassword)
End Sub

Private Sub Cb_maxpassword_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Cb_maxtrytime_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_maxuserid_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_maxuserid_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_maxuserid_GotFocus()
    Cb_maxuserid.SelStart = 0
    Cb_maxuserid.SelLength = Len(Cb_maxuserid)
End Sub

Private Sub Cb_maxuserid_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
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

Private Sub Cb_var_password_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Var_Result_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_var_trytime_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_var_userid_Click()
    f_DataChanged = True
End Sub

'Mike added this event @ 2008-1-0
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
        Set gSystem.crlCurItem = T_nd_succeed
    Case 2
        Set gSystem.crlCurItem = T_nd_fail
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click(Index As Integer)
    
    Select Case Index
    Case 0
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_userid
    Case 1
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_password
    Case 2
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_tryagain
    Case 3
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_iid
    End Select
    frmResourceList.Show vbModal
    
End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
    
        '' Sun added 2007-03-20
        If Len(T_com_iid) > 0 And 1 Then
        End If
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        '节点超时
        If Len(Trim(Cb_timeout)) < 1 Or Trim(Cb_timeout) = "" Then
           Node7_Data1.timeout = 0
        Else
           If CLng(Trim(Cb_timeout)) > 255 Then
              Message ("E051")
              Cb_timeout.SetFocus
              Exit Sub
           Else
              Node7_Data1.timeout = CByte(Val(Cb_timeout) Mod 256)
           End If
        End If
        
        '用户号码最大长度
        If Len(Trim(Cb_maxuserid)) < 1 Or Trim(Cb_maxuserid) = "" Then
           Node7_Data1.maxuserid = 0
        Else
           If CLng(Trim(Cb_maxuserid)) > 255 Then
              Message ("E050")
              Cb_maxuserid.SetFocus
              Exit Sub
           Else
              Node7_Data1.maxuserid = CByte(Val(Cb_maxuserid) Mod 256)
           End If
        End If
        
        '用户口令最大长度
        If Len(Trim(Cb_maxpassword)) < 1 Or Trim(Cb_maxpassword) = "" Then
           Node7_Data1.maxpassword = 0
        Else
           If CLng(Trim(Cb_maxpassword)) > 255 Then
              Message ("E049")
              Cb_maxpassword.SetFocus
              Exit Sub
           Else
              Node7_Data1.maxpassword = CByte(Val(Cb_maxpassword) Mod 256)
           End If
        End If
        
        '按键中断
        If Cb_key_term.ListIndex < 0 Then
           Node7_Data1.key_term = 0
        Else
           Node7_Data1.key_term = CByte(Cb_key_term.ItemData(Cb_key_term.ListIndex))
        End If
        
        '最大尝试次数
        If Cb_maxtrytime.ListIndex < 0 Then
           Node7_Data1.maxtrytime = 3
        Else
           Node7_Data1.maxtrytime = CByte(Cb_maxtrytime.ItemData(Cb_maxtrytime.ListIndex))
        End If
        
        '被访问日志
        If Cb_log.ListIndex < 0 Then
           Node7_Data1.log = 0
        Else
           Node7_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        '验证次数日志
        If Cb_var_trytime.ListIndex <= 0 Then
            Node7_Data1.var_trytime = 0
        Else
            Node7_Data1.var_trytime = CByte(Cb_var_trytime.ItemData(Cb_var_trytime.ListIndex))
        End If
        
        '验证结果日志
        If Cb_var_result.ListIndex <= 0 Then
            Node7_Data1.var_result = 0
        Else
            Node7_Data1.var_result = CByte(Cb_var_result.ItemData(Cb_var_result.ListIndex))
        End If
        
        '用户号码记录
        If Cb_var_userid.ListIndex <= 0 Then
            Node7_Data1.var_userid = 0
        Else
            Node7_Data1.var_userid = CByte(Cb_var_userid.ItemData(Cb_var_userid.ListIndex))
        End If
        
        '用户口令记录
        If Cb_var_password.ListIndex <= 0 Then
            Node7_Data1.var_password = 0
        Else
            Node7_Data1.var_password = CByte(Cb_var_password.ItemData(Cb_var_password.ListIndex))
        End If
        
        Node7_Data1.reserved1(0) = 0
        
        '提示用户输入号码
        If Trim(T_vox_userid.Text) = "" Then
           Node7_Data2.vox_userid = 0
        Else
           If Val(T_vox_userid) > 32767 Then
              Message ("E037")
              T_vox_userid.SetFocus
              Exit Sub
           Else
              Node7_Data2.vox_userid = Val(T_vox_userid.Text)
           End If
        End If
        
        '提示用户输入口令
        If Trim(T_vox_password.Text) = "" Then
           Node7_Data2.vox_password = 0
        Else
           If Val(T_vox_password) > 32767 Then
              Message ("E038")
              T_vox_password.SetFocus
              Exit Sub
           Else
              Node7_Data2.vox_password = Val(T_vox_password.Text)
           End If
        End If
        
        '提示用户重新输入
        If Trim(T_vox_tryagain.Text) = "" Then
           Node7_Data2.vox_tryagain = 0
        Else
           If Val(T_vox_tryagain) > 32767 Then
              Message ("E039")
              T_vox_tryagain.SetFocus
              Exit Sub
           Else
              Node7_Data2.vox_tryagain = Val(T_vox_tryagain.Text)
           End If
        End If
        
        Node7_Data2.reserved1(0) = 0
        
        'COM接口ID
        If Trim(T_com_iid.Text) = "" Then
           Node7_Data2.com_iid = 0
        Else
           If Val(T_com_iid) > 32767 Then
              Message ("E040")
              T_com_iid.SetFocus
              Exit Sub
           Else
              Node7_Data2.com_iid = Val(T_com_iid.Text)
           End If
        End If
        
        Node7_Data2.reserved2(0) = 0
        
        '父节点
        If Trim(T_nd_parent.Text) = "" Then
           Node7_Data2.nd_parent = 0
        Else
           If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
           Else
                Node7_Data2.nd_parent = Val(T_nd_parent.Text)
           End If
        End If
        
        'Succeed node id
        If Trim(T_nd_succeed.Text) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_succeed) > 32767 Or Val(T_nd_succeed) < 256) And Val(T_nd_succeed) <> 0 Then
                Message ("E041")
                T_nd_succeed.SetFocus
                Exit Sub
            Else
                lv_nNewNode = Val(T_nd_succeed.Text)
            End If
        End If
        
        '' Sun added 2007-03-25
        If Node7_Data2.nd_succeed <> lv_nNewNode Then
            Node7_Data2.nd_succeed = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 1)
        End If
        
        'failed node id
        If Trim(T_nd_fail.Text) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_fail) > 32767 Or Val(T_nd_fail) < 256) And Val(T_nd_fail) <> 0 Then
                Message ("E042")
                T_nd_fail.SetFocus
                Exit Sub
           Else
              lv_nNewNode = Val(T_nd_fail.Text)
           End If
        End If
        
        '' Sun added 2007-03-25
        If Node7_Data2.nd_fail <> lv_nNewNode Then
            Node7_Data2.nd_fail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
        End If
        
        Node7_Data2.reserved3(0) = 0
        
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
    If Node7_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
        Message "M144"
    End If
    
    Unload Me
    
End Sub

Private Sub CommandExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer

'初始化
    F_FillPhoneKeyList Cb_key_term, 1
    
'被访日志
    With Cb_log
        .AddItem LoadNationalResString(1178)
        .ItemData(.ListCount - 1) = 0
        For i = 1 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1179)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
'最大尝试次数
    With Cb_maxtrytime
        For i = 0 To 9
            .AddItem Trim(Str(i)) & LoadNationalResString(1180)
            .ItemData(.ListCount - 1) = i
        Next
    End With

'用户口令最大长度
    With Cb_maxpassword
        For i = 0 To 9
            .AddItem Trim(Str(i))
        Next
    End With
    
'用户号码最大长度
    With Cb_maxuserid
        For i = 0 To 9
            .AddItem Trim(Str(i))
        Next
    End With
    
'节点超时
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
'用户口令记录
    RefreshVariablesList Cb_var_password
    
'验证结果记录
    RefreshVariablesList Cb_var_result
    
'用户号码记录
    RefreshVariablesList Cb_var_userid
    
'验证次数记录
    RefreshVariablesList Cb_var_trytime

'节点ID
   T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
'节点编号
   T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
 
    '默认提示用户输入号码语音资源为0
    T_vox_userid.Text = Node7_Data2.vox_userid
    '默认提示用户输入口令语音资源为0
    T_vox_password.Text = Node7_Data2.vox_password
    '默认提示用户重新输入语音资源为0
    T_vox_tryagain.Text = Node7_Data2.vox_tryagain
    '默认失败转节点ID为0
    T_nd_fail.Text = Node7_Data2.nd_fail
    '默认父节点ID为0
    T_nd_parent.Text = Node7_Data2.nd_parent
    '默认COM接口ID为0
    T_com_iid.Text = Node7_Data2.com_iid
    '默认成功转节点ID为0
    T_nd_succeed.Text = Node7_Data2.nd_succeed
    
    Cb_timeout.Text = Node7_Data1.timeout
    Cb_maxpassword.Text = Node7_Data1.maxpassword
    Cb_maxuserid.Text = Node7_Data1.maxuserid
    
    Cb_key_term.ListIndex = SearchItemDataIndex(Cb_key_term, CLng(Node7_Data1.key_term), 11)
    Cb_maxtrytime.ListIndex = SearchItemDataIndex(Cb_maxtrytime, CLng(Node7_Data1.maxtrytime), 3)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node7_Data1.log), 0)
    
    Cb_var_trytime.ListIndex = SearchItemDataIndex(Cb_var_trytime, CLng(Node7_Data1.var_trytime), 0)
    Cb_var_result.ListIndex = SearchItemDataIndex(Cb_var_result, CLng(Node7_Data1.var_result), 0)
    Cb_var_password.ListIndex = SearchItemDataIndex(Cb_var_password, CLng(Node7_Data1.var_password), 0)
    Cb_var_userid.ListIndex = SearchItemDataIndex(Cb_var_userid, CLng(Node7_Data1.var_userid), 0)
    
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

Private Sub T_vox_password_Change()
    f_DataChanged = True
    
    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_password
    
End Sub

Private Sub T_vox_password_GotFocus()
    T_vox_password.SelStart = 0
    T_vox_password.SelLength = Len(T_vox_password)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_password)
    Call SoundResourceIDChanged
    
End Sub

Private Sub T_vox_password_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_tryagain_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_tryagain

End Sub

Private Sub T_vox_tryagain_GotFocus()
    T_vox_tryagain.SelStart = 0
    T_vox_tryagain.SelLength = Len(T_vox_tryagain)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_tryagain)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_tryagain_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_userid_Change()
    f_DataChanged = True
    
    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_userid

End Sub

Private Sub T_vox_userid_GotFocus()
    T_vox_userid.SelStart = 0
    T_vox_userid.SelLength = Len(T_vox_userid)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_userid)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_userid_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
