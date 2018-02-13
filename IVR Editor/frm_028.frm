VERSION 5.00
Begin VB.Form frm_028 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TTS 放音"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frm_028.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1524"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   7122
      Picture         =   "frm_028.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "1948"
      Top             =   4131
      Width           =   333
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   2430
      TabIndex        =   19
      Tag             =   "1007"
      Top             =   4140
      Width           =   1065
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   3900
      TabIndex        =   20
      Tag             =   "1144"
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   2985
      Left            =   3990
      TabIndex        =   29
      Tag             =   "1164"
      Top             =   60
      Width           =   3465
      Begin VB.ComboBox cb_PlayType 
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
         ItemData        =   "frm_028.frx":2F6C
         Left            =   1590
         List            =   "frm_028.frx":2F6E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "请选择"
         Top             =   600
         Width           =   1785
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   3030
         Picture         =   "frm_028.frx":2F70
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "显示节点列表"
         Top             =   2460
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3030
         Picture         =   "frm_028.frx":32FA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "显示节点列表"
         Top             =   2100
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   16
         Top             =   2430
         Width           =   1425
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   14
         Top             =   2070
         Width           =   1425
      End
      Begin VB.TextBox txtAlter 
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   8
         Top             =   960
         Width           =   1425
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   3030
         Picture         =   "frm_028.frx":3684
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1146"
         Top             =   990
         Width           =   315
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1320
         Width           =   1425
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   3030
         Picture         =   "frm_028.frx":3786
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "显示资源列表"
         Top             =   1350
         Width           =   315
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   5
         Top             =   240
         Width           =   1425
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1710
         Width           =   1425
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   3030
         Picture         =   "frm_028.frx":3888
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3030
         Picture         =   "frm_028.frx":398A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "1145"
         Top             =   1740
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "播放类型"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   39
         Tag             =   "1555"
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "失败转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Tag             =   "1171"
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label Lbl_nd_succeed 
         AutoSize        =   -1  'True
         Caption         =   "成功转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Tag             =   "1170"
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "替代播放语音"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Tag             =   "1527"
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "COM接口ID"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Tag             =   "1168"
         Top             =   1410
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TTS播放资源"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Tag             =   "1526"
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1169"
         Top             =   1800
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2985
      Left            =   60
      TabIndex        =   22
      Tag             =   "1136"
      Top             =   60
      Width           =   3855
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1590
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   915
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
         TabIndex        =   37
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
         Left            =   1740
         TabIndex        =   0
         ToolTipText     =   "0 - 255 秒"
         Top             =   750
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "请选择"
         Top             =   1170
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "请选择"
         Top             =   2010
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2430
         Width           =   1965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Tag             =   "1248"
         Top             =   1695
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "按键中断"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Tag             =   "1264"
         Top             =   2085
         Width           =   945
      End
      Begin VB.Label Label9 
         Caption         =   "节点超时(秒)"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Tag             =   "1154"
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1740
         TabIndex        =   25
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "放音清空标志"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Tag             =   "1244"
         Top             =   1290
         Width           =   1080
      End
      Begin VB.Label Lbl_log 
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Tag             =   "1245"
         Top             =   2490
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   885
      Left            =   60
      TabIndex        =   21
      Tag             =   "1104"
      Top             =   3120
      Width           =   7395
      Begin VB.TextBox Txt_Description 
         Height          =   555
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   240
         Width           =   7185
      End
   End
End
Attribute VB_Name = "frm_028"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_breakkey_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_playclear_Click()
    f_DataChanged = True
End Sub


Private Sub cb_PlayType_Click()
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
        gSystem.intCurStep = 2
        Set gSystem.crlCurItem = T_vox_play
    Case 1
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = txtAlter
    Case 2
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_iid
    End Select
    frmResourceList.Show vbModal

End Sub

Private Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then

        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        Node28_Data1.reserved1 = 0
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node28_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node28_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        ' 放音清空
        If Cb_playclear.ListIndex = -1 Then
            Node28_Data1.playclear = 0
        Else
            Node28_Data1.playclear = CByte(Cb_playclear.ItemData(Cb_playclear.ListIndex))
        End If
        
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node28_Data1.usevar = 0
        Else
            Node28_Data1.usevar = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        ' 按键中断
        If Cb_breakkey.ListIndex < 0 Then
            Node28_Data1.breakkey = 0
        Else
            Node28_Data1.breakkey = CByte(Cb_breakkey.ItemData(Cb_breakkey.ListIndex))
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node28_Data1.log = 0
        Else
            Node28_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        Node28_Data1.reserved2(0) = 0
        
        ' TTS播放资源
        If Trim(T_vox_play) = "" Then
            Node28_Data2.vox_string = 0
        Else
            If CLng(Trim(T_vox_play)) > 32767 Then
                Message ("E104")
                T_vox_play.SetFocus
                Exit Sub
            Else
                Node28_Data2.vox_string = CInt(Trim(T_vox_play.Text))
            End If
        End If
        
        ' 替代播放语音
        If Trim(txtAlter) = "" Then
            Node28_Data2.vox_alter = 0
        Else
            If CLng(Trim(txtAlter)) > 32767 Then
                Message ("E105")
                txtAlter.SetFocus
                Exit Sub
            Else
                Node28_Data2.vox_alter = CInt(Trim(txtAlter))
            End If
        End If
        
        ' 播放类型
        If cb_PlayType.ListIndex < 0 Then
            Node28_Data2.playtype = 0
        Else
            Node28_Data2.playtype = CByte(cb_PlayType.ItemData(cb_PlayType.ListIndex))
        End If
        
        ' 保留
        Node28_Data2.reserved1(0) = 0
        Node28_Data2.reserved2(0) = 0
        Node28_Data2.reserved3(0) = 0
        
        ' COM接口ID
        If Trim(T_com_iid.Text) = "" Then
            Node28_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E040")
                T_com_iid.SetFocus
                Exit Sub
            Else
                Node28_Data2.com_iid = CInt(Trim(T_com_iid.Text))
            End If
        End If
      
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node28_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node28_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
        
        ' 成功转节点ID
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
        If Node28_Data2.nd_succ <> lv_nNewNode Then
            Node28_Data2.nd_succ = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 1)
        End If
        
        ' 失败转节点ID
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
        If Node28_Data2.nd_fail <> lv_nNewNode Then
            Node28_Data2.nd_fail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
        End If
        
        ' 描述
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
    If Node28_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
        Message "M144"
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
    
    ' 使用变量ID
    RefreshVariablesList Cb_usevar
    
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

    '输入终止符
    F_FillPhoneKeyList Cb_breakkey, 1

    ' 节点超时(秒)
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With

    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    ' 播放类型
    With cb_PlayType
        .AddItem LoadNationalResString(1564)
        .ItemData(.ListCount - 1) = 0
        .AddItem LoadNationalResString(1565)
        .ItemData(.ListCount - 1) = 1
        .AddItem LoadNationalResString(1566)
        .ItemData(.ListCount - 1) = 2
        .AddItem LoadNationalResString(1567)
        .ItemData(.ListCount - 1) = 3
    End With
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_play.Text = Node28_Data2.vox_string
    txtAlter.Text = Node28_Data2.vox_alter
    T_com_iid.Text = Node28_Data2.com_iid
    T_nd_parent.Text = Node28_Data2.nd_parent
    T_nd_succeed.Text = Node28_Data2.nd_succ
    T_nd_fail.Text = Node28_Data2.nd_fail
    
    Cb_timeout.Text = Node28_Data1.timeout
    Cb_playclear.ListIndex = SearchItemDataIndex(Cb_playclear, CLng(Node28_Data1.playclear), 0)
    Cb_breakkey.ListIndex = SearchItemDataIndex(Cb_breakkey, CLng(Node28_Data1.breakkey), 11)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node28_Data1.log), 0)
    cb_PlayType.ListIndex = SearchItemDataIndex(cb_PlayType, CLng(Node28_Data2.playtype), 0)
 
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node28_Data1.usevar), 0)
 
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

Private Sub CommandExit_Click()
    Unload Me
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

Private Sub T_vox_play_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_play
    
End Sub

Private Sub T_vox_play_GotFocus()
    T_vox_play.SelStart = 0
    T_vox_play.SelLength = Len(T_vox_play)
End Sub

Private Sub T_vox_play_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtAlter_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtAlter

End Sub

Private Sub txtAlter_GotFocus()
    txtAlter.SelStart = 0
    txtAlter.SelLength = Len(txtAlter)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(txtAlter)
    Call SoundResourceIDChanged

End Sub

Private Sub txtAlter_KeyPress(KeyAscii As Integer)
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

