VERSION 5.00
Begin VB.Form frm_055 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "传真接收"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "frm_055.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1677"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   7422
      Picture         =   "frm_055.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "1948"
      Top             =   4851
      Width           =   333
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   5940
      TabIndex        =   29
      Tag             =   "1144"
      Top             =   4830
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   4620
      TabIndex        =   28
      Tag             =   "1007"
      Top             =   4830
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   5235
      Left            =   60
      TabIndex        =   0
      Tag             =   "1136"
      Top             =   60
      Width           =   4485
      Begin VB.ComboBox Cb_Var_ExtNo 
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
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2760
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
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   1
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   2
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
         Left            =   2400
         TabIndex        =   3
         ToolTipText     =   "0 - 255 秒"
         Top             =   660
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   3600
         Width           =   1965
      End
      Begin VB.ComboBox cb_FileNameType 
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
         ItemData        =   "frm_055.frx":1F3C
         Left            =   2400
         List            =   "frm_055.frx":1F3E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "请选择"
         Top             =   1080
         Width           =   1965
      End
      Begin VB.ComboBox Cb_Var_FileName 
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
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1500
         Width           =   1965
      End
      Begin VB.ComboBox Cb_VarAppField 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4710
         Width           =   1965
      End
      Begin VB.ComboBox Cb_VarAppField 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4350
         Width           =   1965
      End
      Begin VB.ComboBox Cb_VarAppField 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3990
         Width           =   1965
      End
      Begin VB.ComboBox Cb_Var_Sender 
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
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   1965
      End
      Begin VB.ComboBox Cb_Var_Receiver 
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
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2340
         Width           =   1965
      End
      Begin VB.ComboBox Cb_Var_Result 
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
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3180
         Width           =   1965
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "分机号码变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Tag             =   "1684"
         Top             =   2865
         Width           =   1245
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   2340
         TabIndex        =   46
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "节点超时(秒)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   44
         Tag             =   "1154"
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label Lbl_log 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Left            =   180
         TabIndex        =   43
         Tag             =   "1159"
         Top             =   3660
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "文件名类型"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   42
         Tag             =   "1678"
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "文件名变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Tag             =   "1679"
         Top             =   1605
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "变量：附加字段3"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   40
         Tag             =   "1620"
         Top             =   4770
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "变量：附加字段2"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   39
         Tag             =   "1619"
         Top             =   4410
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "变量：附加字段1"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   38
         Tag             =   "1618"
         Top             =   4050
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "发出号码变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Tag             =   "1681"
         Top             =   2025
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "接收号码变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Tag             =   "1682"
         Top             =   2445
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "结果记录变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Tag             =   "1683"
         Top             =   3285
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   3255
      Index           =   0
      Left            =   4590
      TabIndex        =   25
      Tag             =   "1164"
      Top             =   60
      Width           =   3195
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   14
         Top             =   240
         Width           =   1065
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   18
         Top             =   1110
         Width           =   1065
      End
      Begin VB.TextBox T_fax_fileid 
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   16
         Top             =   660
         Width           =   1065
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2700
         Picture         =   "frm_055.frx":1F40
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2700
         Picture         =   "frm_055.frx":2042
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "1146"
         Top             =   690
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2700
         Picture         =   "frm_055.frx":2144
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "1145"
         Top             =   1140
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   2700
         Picture         =   "frm_055.frx":24CE
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "显示节点列表"
         Top             =   1980
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2700
         Picture         =   "frm_055.frx":2858
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "显示节点列表"
         Top             =   1560
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   22
         Top             =   1950
         Width           =   1065
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1530
         Width           =   1065
      End
      Begin VB.CheckBox chkRecord_cdr 
         Caption         =   "记录传真详单"
         Height          =   225
         Left            =   210
         TabIndex        =   24
         Tag             =   "1680"
         Top             =   2460
         Width           =   2835
      End
      Begin VB.Label Label3 
         Caption         =   "操作提示音"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   34
         Tag             =   "1266"
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Tag             =   "1169"
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "传真文件"
         Height          =   225
         Left            =   180
         TabIndex        =   32
         Tag             =   "1295"
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "失败转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Tag             =   "1171"
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Lbl_nd_succeed 
         AutoSize        =   -1  'True
         Caption         =   "成功转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1170"
         Top             =   1620
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1155
      Index           =   1
      Left            =   4590
      TabIndex        =   27
      Tag             =   "1104"
      Top             =   3390
      Width           =   3165
      Begin VB.TextBox Txt_Description 
         Height          =   825
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   240
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frm_055"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//         Node information
'//文件名：  Frm_055.frm
'//用途：    传真接收节点
'//作者:     Tony Sun
'//创建日期：2006-12/31
'//修改日期：2006-12-31
'//文件描述：传真接收
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub cb_FileNameType_Click()
On Error Resume Next

    If cb_FileNameType.ListIndex >= 0 Then
    
        f_DataChanged = True
        
        T_fax_fileid.Enabled = False
        Cb_Var_FileName.Enabled = False
        cmdShowRes(1).Enabled = False
        
        Select Case cb_FileNameType.ItemData(cb_FileNameType.ListIndex)
        Case DEF_NODE050_FAXN_TYPE_AUTO
        Case DEF_NODE050_FAXN_TYPE_RESID
            T_fax_fileid.Enabled = True
            cmdShowRes(1).Enabled = True
        Case DEF_NODE050_FAXN_TYPE_VAR2RESID
            Cb_Var_FileName.Enabled = True
        Case DEF_NODE050_FAXN_TYPE_VAR2NAME
            Cb_Var_FileName.Enabled = True
        Case DEF_NODE050_FAXN_TYPE_FORMAT
            T_fax_fileid.Enabled = True
            Cb_Var_FileName.Enabled = True
            cmdShowRes(1).Enabled = True
        End Select
        
    End If
    
End Sub

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

Private Sub Cb_Var_ExtNo_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Var_FileName_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Var_Receiver_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Var_Result_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Var_Sender_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_VarAppField_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub chkRecord_cdr_Click()
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
        Set gSystem.crlCurItem = T_vox_op
    Case 1
        gSystem.intCurStep = 2
        Set gSystem.crlCurItem = T_fax_fileid
    End Select
    frmResourceList.Show vbModal

End Sub

Private Sub CommandExit_Click()
   Unload Me
End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        Dim lv_loop As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack

        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node55_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 3600 Then
                Message ("E126")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node55_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If

        ' 文件名类型
        If cb_FileNameType.ListIndex < 0 Then
           Node55_Data1.filenametype = DEF_NODE050_FAXN_TYPE_RESID
        Else
           Node55_Data1.filenametype = CByte(cb_FileNameType.ItemData(cb_FileNameType.ListIndex))
        End If

        ' 文件名变量ID
        If Cb_Var_FileName.ListIndex <= 0 Then
            Node55_Data1.var_faxfile = 0
        Else
            Node55_Data1.var_faxfile = CByte(Cb_Var_FileName.ItemData(Cb_Var_FileName.ListIndex))
        End If

        ' 发出号码变量ID
        If Cb_Var_Sender.ListIndex <= 0 Then
            Node55_Data1.var_fromno = 0
        Else
            Node55_Data1.var_fromno = CByte(Cb_Var_Sender.ItemData(Cb_Var_Sender.ListIndex))
        End If

        ' 接收号码变量ID
        If Cb_Var_Receiver.ListIndex <= 0 Then
            Node55_Data1.var_tono = 0
        Else
            Node55_Data1.var_tono = CByte(Cb_Var_Receiver.ItemData(Cb_Var_Receiver.ListIndex))
        End If

        ' 分机号码变量ID
        If Cb_Var_ExtNo.ListIndex <= 0 Then
            Node55_Data1.var_extno = 0
        Else
            Node55_Data1.var_extno = CByte(Cb_Var_ExtNo.ItemData(Cb_Var_ExtNo.ListIndex))
        End If

        ' 结果记录变量ID
        If Cb_Var_Result.ListIndex <= 0 Then
            Node55_Data1.var_result = 0
        Else
            Node55_Data1.var_result = CByte(Cb_Var_Result.ItemData(Cb_Var_Result.ListIndex))
        End If

        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node55_Data1.log = 0
        Else
            Node55_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        ' 变量：附加字段
        For lv_loop = 0 To 2
            If Cb_VarAppField(lv_loop).ListIndex <= 0 Then
                Node55_Data1.var_appfield(lv_loop) = 0
            Else
                Node55_Data1.var_appfield(lv_loop) = CByte(Cb_VarAppField(lv_loop).ItemData(Cb_VarAppField(lv_loop).ListIndex))
            End If
        Next
        
        ' 传真文件
        If Trim(T_fax_fileid) = "" Then
            Node55_Data2.fax_fileid = 0
        Else
            If CLng(Trim(T_fax_fileid)) > 32767 Then
                Message ("E093")
                T_fax_fileid.SetFocus
                Exit Sub
            Else
                Node55_Data2.fax_fileid = CInt(Trim(T_fax_fileid.Text))
            End If
        End If
        
        Node55_Data2.reserved1(0) = 0
        
        ' 操作语音提示
        If Trim(T_vox_op) = "" Then
            Node55_Data2.vox_op = 0
        Else
            If CLng(Trim(T_vox_op)) > 32767 Then
                Message ("E092")
                T_vox_op.SetFocus
                Exit Sub
            Else
                Node55_Data2.vox_op = CInt(Trim(T_vox_op))
            End If
        End If
        
        Node55_Data2.reserved2(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node55_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node55_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
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
        If Node55_Data2.nd_succ <> lv_nNewNode Then
            Node55_Data2.nd_succ = lv_nNewNode
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
        If Node55_Data2.nd_fail <> lv_nNewNode Then
            Node55_Data2.nd_fail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
        End If
        
        ' 记录传真详单
        Node55_Data1.record_cdr = chkRecord_cdr
    
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

Dim i As Integer

SetMainFormItemsEnableWhenPropertyShow False

    ' 节点超时(秒)
    With Cb_timeout
        For i = 0 To 600 Step 30
            .AddItem Trim(Str(i))
        Next
    End With

    ' 文件名类型
    With cb_FileNameType
        .Clear
        For i = DEF_NODE050_FAXN_TYPE_AUTO To DEF_NODE050_FAXN_TYPE_FORMAT
            .AddItem Trim(Str(i)) & "-" & LoadNationalResString(1685 + i)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
    ' 文件名变量ID
    RefreshVariablesList Cb_Var_FileName
        
    ' 发出号码变量ID
    RefreshVariablesList Cb_Var_Sender
        
    ' 接收号码变量ID
    RefreshVariablesList Cb_Var_Receiver
        
    ' 分机号码变量ID
    RefreshVariablesList Cb_Var_ExtNo
    
    ' 结果记录变量ID
    RefreshVariablesList Cb_Var_Result
        
    ' 变量：附加字段1
    RefreshVariablesList Cb_VarAppField(0)
        
    ' 变量：附加字段2
    RefreshVariablesList Cb_VarAppField(1)
        
    ' 变量：附加字段3
    RefreshVariablesList Cb_VarAppField(2)
    
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
    
    ' 操作提示音
    T_vox_op.Text = Node55_Data2.vox_op
    
    ' 传真文件
    T_fax_fileid.Text = Node55_Data2.fax_fileid
    
    ' 父节点ID
    T_nd_parent.Text = Node55_Data2.nd_parent
    
    ' 成功转节点ID
    T_nd_succeed.Text = Node55_Data2.nd_succ
    
    ' 失败转节点ID
    T_nd_fail.Text = Node55_Data2.nd_fail
    
    ' 记录传真详单
    chkRecord_cdr = Node55_Data1.record_cdr
    
    ' 描述
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    
    Cb_timeout.Text = Node55_Data1.timeout
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node55_Data1.log), 0)
    
    cb_FileNameType.ListIndex = SearchItemDataIndex(cb_FileNameType, CLng(Node55_Data1.filenametype), 1)
    Cb_Var_FileName.ListIndex = SearchItemDataIndex(Cb_Var_FileName, CLng(Node55_Data1.var_faxfile), 0)
    Cb_Var_Sender.ListIndex = SearchItemDataIndex(Cb_Var_Sender, CLng(Node55_Data1.var_fromno), 0)
    Cb_Var_Receiver.ListIndex = SearchItemDataIndex(Cb_Var_Receiver, CLng(Node55_Data1.var_tono), 0)
    Cb_Var_ExtNo.ListIndex = SearchItemDataIndex(Cb_Var_ExtNo, CLng(Node55_Data1.var_extno), 0)
    Cb_Var_Result.ListIndex = SearchItemDataIndex(Cb_Var_Result, CLng(Node55_Data1.var_result), 0)
    
    Cb_VarAppField(0).ListIndex = SearchItemDataIndex(Cb_VarAppField(0), CLng(Node55_Data1.var_appfield(0)), 0)
    Cb_VarAppField(1).ListIndex = SearchItemDataIndex(Cb_VarAppField(1), CLng(Node55_Data1.var_appfield(1)), 0)
    Cb_VarAppField(2).ListIndex = SearchItemDataIndex(Cb_VarAppField(2), CLng(Node55_Data1.var_appfield(2)), 0)
   
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me
    
End Sub

Private Sub T_fax_fileid_Change()
    f_DataChanged = True
End Sub

Private Sub T_fax_fileid_GotFocus()
    T_fax_fileid.SelStart = 0
    T_fax_fileid.SelLength = Len(T_fax_fileid)
End Sub

Private Sub T_fax_fileid_KeyPress(KeyAscii As Integer)
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

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
