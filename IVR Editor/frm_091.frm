VERSION 5.00
Begin VB.Form frm_091 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calling Card"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frm_091.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1584"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   7602
      Picture         =   "frm_091.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "1948"
      Top             =   4431
      Width           =   333
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   2580
      TabIndex        =   41
      Tag             =   "1007"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   3900
      TabIndex        =   42
      Tag             =   "1144"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   4200
      TabIndex        =   19
      Tag             =   "1164"
      Top             =   60
      Width           =   3735
      Begin VB.TextBox txtOutofMoney 
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
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   26
         Top             =   960
         Width           =   1035
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   4
         Left            =   3300
         Picture         =   "frm_091.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "显示资源列表"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   3300
         Picture         =   "frm_091.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "显示资源列表"
         Top             =   630
         Width           =   315
      End
      Begin VB.TextBox txtRemindVox 
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
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   24
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox T_com_fee 
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
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   31
         Top             =   1680
         Width           =   1035
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   3300
         Picture         =   "frm_091.frx":1780
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "显示资源列表"
         Top             =   1710
         Width           =   315
      End
      Begin VB.TextBox T_vox_pred 
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
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   21
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
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   34
         Top             =   2040
         Width           =   1035
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
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   28
         Top             =   1320
         Width           =   1035
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
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   37
         Top             =   2400
         Width           =   1035
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   3300
         Picture         =   "frm_091.frx":1882
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "显示资源列表"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   3300
         Picture         =   "frm_091.frx":1984
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "显示资源列表"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3300
         Picture         =   "frm_091.frx":1A86
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "1145"
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3300
         Picture         =   "frm_091.frx":1E10
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "1145"
         Top             =   2430
         Width           =   315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "余额不足提示语音"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Tag             =   "1598"
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "超时提醒播放语音"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Tag             =   "1597"
         Top             =   690
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "计费COM"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1590"
         Top             =   1770
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "前导播放语音"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Tag             =   "1250"
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Tag             =   "1169"
         Top             =   2130
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "通话时长COM"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Tag             =   "1589"
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "子节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Tag             =   "1252"
         Top             =   2460
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1185
      Left            =   4200
      TabIndex        =   39
      Tag             =   "1104"
      Top             =   3030
      Width           =   3735
      Begin VB.TextBox Txt_Description 
         Height          =   855
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   4155
      Left            =   60
      TabIndex        =   0
      Tag             =   "1136"
      Top             =   60
      Width           =   4035
      Begin VB.TextBox txtOBGroup 
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
         MaxLength       =   3
         TabIndex        =   17
         Top             =   3300
         Width           =   1935
      End
      Begin VB.TextBox txtRemindTime 
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
         MaxLength       =   2
         TabIndex        =   18
         Top             =   3660
         Width           =   1935
      End
      Begin VB.ComboBox cmbVar_TelNo 
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
         TabIndex        =   12
         Top             =   2040
         Width           =   1965
      End
      Begin VB.ComboBox cmbVar_Card 
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
         TabIndex        =   10
         Top             =   1620
         Width           =   1965
      End
      Begin VB.ComboBox Cb_var_calllength 
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
         TabIndex        =   16
         Top             =   2880
         Width           =   1965
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
         TabIndex        =   2
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
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
         TabIndex        =   6
         ToolTipText     =   "0 - 255 秒"
         Top             =   780
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
         TabIndex        =   14
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2460
         Width           =   1965
      End
      Begin VB.ComboBox cmbTimeFormat 
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
         TabIndex        =   8
         ToolTipText     =   "请选择"
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "外拨组号"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Tag             =   "1595"
         Top             =   3420
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "超时提醒(分钟)"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Tag             =   "1596"
         Top             =   3780
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "号码使用变量"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Tag             =   "1587"
         Top             =   2130
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卡号使用变量"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Tag             =   "1586"
         Top             =   1710
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "通话时长记录"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Tag             =   "1588"
         Top             =   2970
         Width           =   1080
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1950
         TabIndex        =   3
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "节点超时(秒)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Tag             =   "1154"
         Top             =   870
         Width           =   1080
      End
      Begin VB.Label Lbl_log 
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Tag             =   "1245"
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "时间播报方式"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Tag             =   "1585"
         Top             =   1290
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frm_091"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'// 文件名：  Frm_091.frm
'// 用途：    Calling Card
'// 作者:     Tony Sun
'// 创建日期：2005/05/26
'// 修改日期：
'// 文件描述：Calling Card
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
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

Private Sub Cb_var_calllength_Click()
    f_DataChanged = True
End Sub

Private Sub cmbTimeFormat_Click()
    f_DataChanged = True

    If cmbTimeFormat.ListIndex = 0 Then
        
    End If
End Sub

Private Sub cmbVar_Card_Click()
    f_DataChanged = True
End Sub

Private Sub cmbVar_TelNo_Click()
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
        Set gSystem.crlCurItem = T_nd_child
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click(Index As Integer)

    Select Case Index
    Case 0
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_pred
    Case 1
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_iid
    Case 2
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_fee
    Case 3
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = txtRemindVox
    End Select
    frmResourceList.Show vbModal

End Sub

Private Sub CommandExit_Click()
   Unload Me
End Sub

Private Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node91_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node91_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        Node91_Data1.reserved1 = 0
        
        ' 可通话时长播报方式
        If cmbTimeFormat.ListIndex = -1 Then
            Node91_Data1.talklentype = 0
        Else
            Node91_Data1.talklentype = CByte(cmbTimeFormat.ItemData(cmbTimeFormat.ListIndex))
        End If
        
        ' 外拨组号
        If Trim(txtOBGroup) = "" Then
            Node91_Data1.obgroup = 0
        Else
            If Val(txtOBGroup) > 255 Then
                Message ("E106")
                txtOBGroup.SetFocus
                Exit Sub
            Else
                Node91_Data1.obgroup = CByte(Val(txtOBGroup) Mod 256)
            End If
        End If
        
        ' 超时提醒(分钟)
        If Trim(txtRemindTime) = "" Then
            Node91_Data1.remindminute = 0
        Else
            If Val(txtRemindTime) > 60 Then
                Message ("E124")
                txtRemindTime.SetFocus
                Exit Sub
            Else
                Node91_Data1.remindminute = CByte(Val(txtRemindTime) Mod 256)
            End If
        End If
                
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node91_Data1.log = 0
        Else
            Node91_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
            
        Node91_Data1.reserved2 = 0
        
        ' 使用变量ID
        If cmbVar_Card.ListIndex <= 0 Then
            Node91_Data1.var_cardno = 0
        Else
            Node91_Data1.var_cardno = CByte(cmbVar_Card.ItemData(cmbVar_Card.ListIndex))
        End If
        
        If cmbVar_TelNo.ListIndex <= 0 Then
            Node91_Data1.var_telno = 0
        Else
            Node91_Data1.var_telno = CByte(cmbVar_TelNo.ItemData(cmbVar_TelNo.ListIndex))
        End If
        
        If Cb_var_calllength.ListIndex <= 0 Then
            Node91_Data1.var_connectlength = 0
        Else
            Node91_Data1.var_connectlength = CByte(Cb_var_calllength.ItemData(Cb_var_calllength.ListIndex))
        End If
        
        Node91_Data1.reserved3(0) = 0
        
        ' 前导播放语音
        If Trim(T_vox_pred) = "" Then
            Node91_Data2.vox_talklen = 0
        Else
            If CLng(Trim(T_vox_pred)) > 32767 Then
                Message ("E090")
                T_vox_pred.SetFocus
                Exit Sub
            Else
                Node91_Data2.vox_talklen = CInt(Trim(T_vox_pred.Text))
            End If
        End If
        
        ' 超时提醒播放语音
        If Trim(txtRemindVox) = "" Then
            Node91_Data2.vox_timeout = 0
        Else
            If CLng(Trim(txtRemindVox)) > 32767 Then
                Message ("E076")
                txtRemindVox.SetFocus
                Exit Sub
            Else
                Node91_Data2.vox_timeout = CInt(Trim(txtRemindVox.Text))
            End If
        End If
        
        ' 余额不足或卡无效提示语音
        If Trim(txtOutofMoney) = "" Then
            Node91_Data2.vox_noservice = 0
        Else
            If CLng(Trim(txtOutofMoney)) > 32767 Then
                Message ("E107")
                txtOutofMoney.SetFocus
                Exit Sub
            Else
                Node91_Data2.vox_noservice = CInt(Trim(txtOutofMoney.Text))
            End If
        End If
        
        Node91_Data2.reserved1(0) = 0
        
        ' COM接口ID
        If Trim(T_com_iid.Text) = "" Then
            Node91_Data2.com_talklength = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E040")
                T_com_iid.SetFocus
                Exit Sub
            Else
                Node91_Data2.com_talklength = CInt(Trim(T_com_iid.Text))
            End If
        End If

        If Trim(T_com_fee.Text) = "" Then
            Node91_Data2.com_billing = 0
        Else
            If CLng(Trim(T_com_fee)) > 32767 Then
                Message ("E040")
                T_com_fee.SetFocus
                Exit Sub
            Else
                Node91_Data2.com_billing = CInt(Trim(T_com_fee.Text))
            End If
        End If

        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node91_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node91_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If

        ' 子节点
        If Trim(T_nd_child) = "" Then
            Node91_Data2.nd_child = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                Node91_Data2.nd_child = CInt(Trim(T_nd_child.Text))
            End If
         End If

        ' 保留
        Node91_Data2.reserved2(0) = 0
        
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
    If (Node91_Data2.com_billing > 0 Or Node91_Data2.com_talklength > 0) And Node0_Data2.MainCOM = 0 Then
        Message "M144"
    End If

    Unload Me
    
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer

    ' 使用变量ID
    RefreshVariablesList cmbVar_Card
    RefreshVariablesList cmbVar_TelNo
    RefreshVariablesList Cb_var_calllength
    
    ' 节点超时(秒)
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With

    ' 可通话时长播报方式
    With cmbTimeFormat
        .AddItem LoadNationalResString(1591)
        .ItemData(.ListCount - 1) = 0
        .AddItem LoadNationalResString(1592)
        .ItemData(.ListCount - 1) = 1
        .AddItem LoadNationalResString(1593)
        .ItemData(.ListCount - 1) = 2
        .AddItem LoadNationalResString(1594)
        .ItemData(.ListCount - 1) = 3
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
    txtOBGroup = Node91_Data1.obgroup
    txtRemindTime = Node91_Data1.remindminute
    T_vox_pred.Text = Node91_Data2.vox_talklen
    txtRemindVox = Node91_Data2.vox_timeout
    txtOutofMoney = Node91_Data2.vox_noservice
    T_com_iid.Text = Node91_Data2.com_talklength
    T_com_fee.Text = Node91_Data2.com_billing
    T_nd_parent.Text = Node91_Data2.nd_parent
    T_nd_child.Text = Node91_Data2.nd_child
    
    Cb_timeout.Text = Node91_Data1.timeout
    cmbTimeFormat.ListIndex = SearchItemDataIndex(cmbTimeFormat, CLng(Node91_Data1.talklentype), 0)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node91_Data1.log), 0)
    
    cmbVar_Card.ListIndex = SearchItemDataIndex(cmbVar_Card, CLng(Node91_Data1.var_cardno), 0)
    cmbVar_TelNo.ListIndex = SearchItemDataIndex(cmbVar_TelNo, CLng(Node91_Data1.var_telno), 0)
    Cb_var_calllength.ListIndex = SearchItemDataIndex(Cb_var_calllength, CLng(Node91_Data1.var_connectlength), 0)
 
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If f_DataChanged Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub T_com_fee_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_com_fee

End Sub

Private Sub T_com_fee_GotFocus()
    T_com_fee.SelStart = 0
    T_com_fee.SelLength = Len(T_com_fee)
End Sub

Private Sub T_com_fee_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
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

Private Sub T_vox_pred_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_pred
    
End Sub

Private Sub T_vox_pred_GotFocus()
    T_vox_pred.SelStart = 0
    T_vox_pred.SelLength = Len(T_vox_pred)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_pred)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_pred_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtOBGroup_Change()
    f_DataChanged = True
End Sub

Private Sub txtOBGroup_GotFocus()
    txtOBGroup.SelStart = 0
    txtOBGroup.SelLength = Len(txtOBGroup)
End Sub

Private Sub txtOBGroup_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtOutofMoney_Change()
    f_DataChanged = True
End Sub

Private Sub txtOutofMoney_GotFocus()
    txtOutofMoney.SelStart = 0
    txtOutofMoney.SelLength = Len(txtOutofMoney)
End Sub

Private Sub txtOutofMoney_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtRemindTime_Change()
    f_DataChanged = True
End Sub

Private Sub txtRemindTime_GotFocus()
    txtRemindTime.SelStart = 0
    txtRemindTime.SelLength = Len(txtRemindTime)
End Sub

Private Sub txtRemindTime_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtRemindVox_Change()
    f_DataChanged = True
End Sub

Private Sub txtRemindVox_GotFocus()
    txtRemindVox.SelStart = 0
    txtRemindVox.SelLength = Len(txtRemindVox)
End Sub

Private Sub txtRemindVox_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
