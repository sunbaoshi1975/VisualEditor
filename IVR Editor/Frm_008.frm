VERSION 5.00
Begin VB.Form Frm_008 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改口令"
   ClientHeight    =   4410
   ClientLeft      =   2790
   ClientTop       =   2550
   ClientWidth     =   7290
   Icon            =   "Frm_008.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1039"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   3522
      Picture         =   "Frm_008.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "1948"
      Top             =   3981
      Width           =   333
   End
   Begin VB.Frame Frame3 
      Caption         =   "描述"
      Height          =   975
      Left            =   3930
      TabIndex        =   49
      Tag             =   "1104"
      Top             =   3360
      Width           =   3285
      Begin VB.TextBox Txt_Description 
         Height          =   615
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   270
         Width           =   3105
      End
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   24
      Tag             =   "1007"
      Top             =   3990
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   1290
      TabIndex        =   25
      Tag             =   "1144"
      Top             =   3990
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3795
      Left            =   60
      TabIndex        =   29
      Tag             =   "1136"
      Top             =   60
      Width           =   3795
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   46
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
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3270
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2910
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2550
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   44
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2190
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "请选择"
         Top             =   1830
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
         ItemData        =   "Frm_008.frx":1F3C
         Left            =   1680
         List            =   "Frm_008.frx":1F3E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "请选择"
         Top             =   1470
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
         Left            =   1680
         TabIndex        =   1
         ToolTipText     =   "0 - 255 个按键"
         Top             =   1110
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
         Left            =   1680
         TabIndex        =   0
         ToolTipText     =   "0 - 255 秒"
         Top             =   750
         Width           =   1965
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1680
         TabIndex        =   47
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_var_password 
         AutoSize        =   -1  'True
         Caption         =   "用户口令记录"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Tag             =   "1162"
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label Lbl_var_result 
         AutoSize        =   -1  'True
         Caption         =   "修改结果记录"
         Height          =   180
         Left            =   180
         TabIndex        =   36
         Tag             =   "1161"
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label Lbl_var_trytime 
         AutoSize        =   -1  'True
         Caption         =   "尝试次数记录"
         Height          =   180
         Left            =   180
         TabIndex        =   35
         Tag             =   "1160"
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label Lbl_log 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Tag             =   "1159"
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Lbl_trytime 
         AutoSize        =   -1  'True
         Caption         =   "最大尝试次数"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Tag             =   "1158"
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label Lbl_key_term 
         AutoSize        =   -1  'True
         Caption         =   "输入终止符"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Tag             =   "1157"
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Lbl_maxpassword 
         AutoSize        =   -1  'True
         Caption         =   "口令最大长度"
         Height          =   180
         Left            =   180
         TabIndex        =   31
         Tag             =   "1156"
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "节点超时(秒)"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1154"
         Top             =   810
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   3930
      TabIndex        =   26
      Tag             =   "1164"
      Top             =   60
      Width           =   3285
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   2850
         Picture         =   "Frm_008.frx":1F40
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "1145"
         Top             =   2790
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2850
         Picture         =   "Frm_008.frx":22CA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "1145"
         Top             =   2430
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2850
         Picture         =   "Frm_008.frx":2654
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "1145"
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   4
         Left            =   2850
         Picture         =   "Frm_008.frx":29DE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "1146"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   2850
         Picture         =   "Frm_008.frx":2AE0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "1146"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2850
         Picture         =   "Frm_008.frx":2BE2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "1146"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2850
         Picture         =   "Frm_008.frx":2CE4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "1146"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2850
         Picture         =   "Frm_008.frx":2DE6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2760
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2400
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2040
         Width           =   885
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   15
         ToolTipText     =   "输入COM资源ID"
         Top             =   1680
         Width           =   885
      End
      Begin VB.TextBox T_vox_succeed 
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   13
         ToolTipText     =   "输入语音资源ID"
         Top             =   1320
         Width           =   885
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   11
         ToolTipText     =   "输入语音资源ID"
         Top             =   960
         Width           =   885
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   7
         ToolTipText     =   "输入语音资源ID"
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox T_vox_confirm 
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
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   9
         ToolTipText     =   "输入语音资源ID"
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "失败转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Tag             =   "1171"
         Top             =   2850
         Width           =   1065
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
      Begin VB.Label Lbl_nd_parent 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Tag             =   "1169"
         Top             =   2130
         Width           =   705
      End
      Begin VB.Label Lbl_com_iid 
         AutoSize        =   -1  'True
         Caption         =   "COM接口ID"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Tag             =   "1168"
         Top             =   1740
         Width           =   885
      End
      Begin VB.Label Lbl_vox_succeed 
         AutoSize        =   -1  'True
         Caption         =   "提示用户修改成功"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Tag             =   "1551"
         Top             =   1380
         Width           =   1440
      End
      Begin VB.Label Lbl_vox_tryagain 
         AutoSize        =   -1  'True
         Caption         =   "两次不一至重新输入"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Tag             =   "1550"
         Top             =   1020
         Width           =   1620
      End
      Begin VB.Label Lbl_vox_password 
         AutoSize        =   -1  'True
         Caption         =   "提示用户输入新口令"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Tag             =   "1548"
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label Lbl_vox_confirm 
         AutoSize        =   -1  'True
         Caption         =   "提示用户再次确认"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Tag             =   "1549"
         Top             =   660
         Width           =   1440
      End
   End
End
Attribute VB_Name = "Frm_008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_008.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：修改口令
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
        Set gSystem.crlCurItem = T_vox_password
    Case 1
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_confirm
    Case 2
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_tryagain
    Case 3
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_succeed
    Case 4
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_iid
    End Select
    frmResourceList.Show vbModal

End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        '节点超时
        If Len(Trim(Cb_timeout)) < 1 Or Trim(Cb_timeout) = "" Then
            Node8_Data1.timeout = 0
        Else
            If CLng(Trim(Cb_timeout)) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node8_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        '用户口令最大长度
        If Len(Trim(Cb_maxpassword)) < 1 Or Trim(Cb_maxpassword) = "" Then
            Node8_Data1.maxpassword = 0
        Else
            If CLng(Trim(Cb_maxpassword)) > 255 Then
                Message ("E049")
                Cb_maxpassword.SetFocus
                Exit Sub
            Else
                Node8_Data1.maxpassword = CByte(Val(Cb_maxpassword) Mod 256)
            End If
        End If
        
        Node8_Data1.reserved1 = 0
    
        '按键中断
        If Cb_key_term.ListIndex < 0 Then
            Node8_Data1.key_term = 0
        Else
            Node8_Data1.key_term = CByte(Cb_key_term.ItemData(Cb_key_term.ListIndex))
        End If
        
        '最大尝试次数
        If Cb_maxtrytime.ListIndex < 0 Then
           Node8_Data1.maxtrytime = 3
        Else
           Node8_Data1.maxtrytime = CByte(Cb_maxtrytime.ItemData(Cb_maxtrytime.ListIndex))
        End If
        
        '被访问日志
        If Cb_log.ListIndex < 0 Then
           Node8_Data1.log = 0
        Else
           Node8_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        '尝试次数日志
        If Cb_var_trytime.ListIndex <= 0 Then
            Node8_Data1.var_trytime = 0
        Else
            Node8_Data1.var_trytime = CByte(Cb_var_trytime.ItemData(Cb_var_trytime.ListIndex))
        End If
        
        '修改结果日志
        If Cb_var_result.ListIndex <= 0 Then
            Node8_Data1.var_result = 0
        Else
            Node8_Data1.var_result = CByte(Cb_var_result.ItemData(Cb_var_result.ListIndex))
        End If
    
        Node8_Data1.reserved2 = 0
        
        '用户口令记录
        If Cb_var_password.ListIndex <= 0 Then
            Node8_Data1.var_password = 0
        Else
            Node8_Data1.var_password = CByte(Cb_var_password.ItemData(Cb_var_password.ListIndex))
        End If
        
        Node8_Data1.reserved3(0) = 0
        
        '提示用户输入新口令
        If Trim(T_vox_password) = "" Then
            Node8_Data2.vox_password = 0
        Else
            If CLng(Trim(T_vox_password)) > 32767 Then
                Message ("E053")
                T_vox_password.SetFocus
                Exit Sub
            Else
                Node8_Data2.vox_password = CInt(Trim(T_vox_password.Text))
            End If
        End If
    
        '提示用户再次确认
        If Trim(T_vox_confirm) = "" Then
            Node8_Data2.vox_confirm = 0
        Else
            If CLng(Trim(T_vox_confirm)) > 32767 Then
                Message ("E054")
                T_vox_confirm.SetFocus
                Exit Sub
            Else
                Node8_Data2.vox_confirm = CInt(Trim(T_vox_confirm.Text))
            End If
        End If
        
        '两次不一致重新输入
        If Trim(T_vox_tryagain) = "" Then
           Node8_Data2.vox_tryagain = 0
        Else
            If CLng(Trim(T_vox_tryagain)) > 32767 Then
                Message ("E055")
                T_vox_tryagain.SetFocus
                Exit Sub
            Else
                Node8_Data2.vox_tryagain = CInt(Trim(T_vox_tryagain.Text))
            End If
        End If
        
        '提示用户修改成功
        If Trim(T_vox_succeed) = "" Then
            Node8_Data2.vox_succeed = 0
        Else
            If CLng(Trim(T_vox_succeed)) > 32767 Then
                Message ("E056")
                T_vox_succeed.SetFocus
                Exit Sub
            Else
                Node8_Data2.vox_succeed = CInt(Trim(T_vox_succeed.Text))
            End If
        End If
        
        '保留
        Node8_Data2.reserved1(0) = 0
        
        'COM接口ID
        If Trim(T_com_iid.Text) = "" Then
            Node8_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E040")
                T_com_iid.SetFocus
                Exit Sub
            Else
                Node8_Data2.com_iid = CInt(Trim(T_com_iid.Text))
            End If
        End If
        
        '保留
        Node8_Data2.reserved2(0) = 0
        
        '父节点
        If Trim(T_nd_parent) = "" Then
            Node8_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node8_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
        
        'Succeed node id
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
        If Node8_Data2.nd_succeed <> lv_nNewNode Then
            Node8_Data2.nd_succeed = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 1)
        End If
        
        'failed node id
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
        If Node8_Data2.nd_fail <> lv_nNewNode Then
            Node8_Data2.nd_fail = lv_nNewNode
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
    If Node8_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
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
'输入终止符
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
      
      
'节点超时
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
'用户口令记录
    RefreshVariablesList Cb_var_password

'修改结果记录
    RefreshVariablesList Cb_var_result

'尝试次数记录
    RefreshVariablesList Cb_var_trytime
  
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
    T_vox_confirm.Text = Node8_Data2.vox_confirm
    T_vox_password.Text = Node8_Data2.vox_password
    T_vox_tryagain.Text = Node8_Data2.vox_tryagain
    T_vox_succeed.Text = Node8_Data2.vox_succeed
    T_nd_fail.Text = Node8_Data2.nd_fail
    T_nd_parent.Text = Node8_Data2.nd_parent
    T_com_iid.Text = Node8_Data2.com_iid
    T_nd_succeed.Text = Node8_Data2.nd_succeed
    
    Cb_timeout.Text = Node8_Data1.timeout
    Cb_maxpassword.Text = Node8_Data1.maxpassword
    
    Cb_key_term.ListIndex = SearchItemDataIndex(Cb_key_term, CLng(Node8_Data1.key_term), 11)
    Cb_maxtrytime.ListIndex = SearchItemDataIndex(Cb_maxtrytime, CLng(Node8_Data1.maxtrytime), 3)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node8_Data1.log), 0)
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    
    Cb_var_trytime.ListIndex = SearchItemDataIndex(Cb_var_trytime, CLng(Node8_Data1.var_trytime), 0)
    Cb_var_result.ListIndex = SearchItemDataIndex(Cb_var_result, CLng(Node8_Data1.var_result), 0)
    Cb_var_password.ListIndex = SearchItemDataIndex(Cb_var_password, CLng(Node8_Data1.var_password), 0)
    
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

Private Sub T_vox_confirm_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_confirm

End Sub

Private Sub T_vox_confirm_GotFocus()
    T_vox_confirm.SelStart = 0
    T_vox_confirm.SelLength = Len(T_vox_confirm)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_confirm)
    Call SoundResourceIDChanged

End Sub
Private Sub T_vox_confirm_KeyPress(KeyAscii As Integer)
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

Private Sub T_vox_succeed_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_succeed

End Sub

Private Sub T_vox_succeed_GotFocus()
    T_vox_succeed.SelStart = 0
    T_vox_succeed.SelLength = Len(T_vox_succeed)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_succeed)
    Call SoundResourceIDChanged

End Sub
Private Sub T_vox_succeed_KeyPress(KeyAscii As Integer)
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

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
