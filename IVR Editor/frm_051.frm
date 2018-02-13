VERSION 5.00
Begin VB.Form frm_051 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TTF传真"
   ClientHeight    =   3405
   ClientLeft      =   3675
   ClientTop       =   2670
   ClientWidth     =   6495
   Icon            =   "frm_051.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1298"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   2892
      Picture         =   "frm_051.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "1948"
      Top             =   2961
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1155
      Index           =   1
      Left            =   60
      TabIndex        =   35
      Tag             =   "1104"
      Top             =   1680
      Width           =   3165
      Begin VB.TextBox Txt_Description 
         Height          =   825
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   3255
      Index           =   0
      Left            =   3270
      TabIndex        =   24
      Tag             =   "1164"
      Top             =   60
      Width           =   3165
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2700
         Picture         =   "frm_051.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "1145"
         Top             =   2790
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2700
         Picture         =   "frm_051.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "1145"
         Top             =   2430
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   5
         Left            =   2700
         Picture         =   "frm_051.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "1146"
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   4
         Left            =   2700
         Picture         =   "frm_051.frx":2752
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "1146"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   3
         Left            =   2700
         Picture         =   "frm_051.frx":2854
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "1146"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2700
         Picture         =   "frm_051.frx":2956
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1146"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2700
         Picture         =   "frm_051.frx":2A58
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "1146"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2700
         Picture         =   "frm_051.frx":2B5A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtFromNo 
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
         TabIndex        =   12
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox txtHeader 
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
         TabIndex        =   10
         Top             =   1320
         Width           =   1065
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   14
         Top             =   2040
         Width           =   1065
      End
      Begin VB.TextBox T_fax_format 
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
         TabIndex        =   8
         Top             =   960
         Width           =   1065
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   4
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
         TabIndex        =   16
         Top             =   2400
         Width           =   1065
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2760
         Width           =   1065
      End
      Begin VB.TextBox T_fax_logo 
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
         TabIndex        =   6
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "传真标题资源ID"
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   37
         Tag             =   "1296"
         Top             =   1410
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "发出号码资源ID"
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   36
         Tag             =   "1297"
         Top             =   1770
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "COM接口ID"
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   34
         Tag             =   "1168"
         Top             =   2130
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "表格格式文件"
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   33
         Tag             =   "1300"
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "LOGO文件"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   32
         Tag             =   "1299"
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "操作提示音"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Tag             =   "1266"
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1169"
         Top             =   2490
         Width           =   765
      End
      Begin VB.Label Label19 
         Caption         =   "子节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Tag             =   "1252"
         Top             =   2850
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   60
      TabIndex        =   23
      Tag             =   "1136"
      Top             =   60
      Width           =   3165
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
         Left            =   1620
         TabIndex        =   2
         ToolTipText     =   "0 - 255 秒"
         Top             =   660
         Width           =   1395
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
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1050
         Width           =   1395
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "节点超时(秒)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   28
         Tag             =   "1154"
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label Lbl_log 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Left            =   180
         TabIndex        =   27
         Tag             =   "1159"
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1380
         TabIndex        =   26
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   21
      Tag             =   "1007"
      Top             =   2940
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   1410
      TabIndex        =   22
      Tag             =   "1144"
      Top             =   2940
      Width           =   1035
   End
End
Attribute VB_Name = "frm_051"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_051.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：TTF传真
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
        Set gSystem.crlCurItem = T_vox_op
    Case 1
        gSystem.intCurStep = 2
        Set gSystem.crlCurItem = T_fax_logo
    Case 2
        gSystem.intCurStep = 2
        Set gSystem.crlCurItem = T_fax_format
    Case 3
        gSystem.intCurStep = 2
        Set gSystem.crlCurItem = txtHeader
    Case 4
        gSystem.intCurStep = 2
        Set gSystem.crlCurItem = txtFromNo
    Case 5
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_iid
    End Select
    frmResourceList.Show vbModal

End Sub

Private Sub CommandExit_Click()
   Unload Me
End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack

        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node51_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node51_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If

        Node51_Data1.reserved1(0) = 0
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node51_Data1.log = 0
        Else
            Node51_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        Node51_Data1.reserved2(0) = 0
        
        ' 操作语音提示
        If Trim(T_vox_op) = "" Then
            Node51_Data2.vox_op = 0
        Else
            If CLng(Trim(T_vox_op)) > 32767 Then
                Message ("E092")
                T_vox_op.SetFocus
                Exit Sub
            Else
                Node51_Data2.vox_op = CInt(Trim(T_vox_op))
            End If
        End If

        ' 传真标题资源ID
        If Trim(txtHeader) = "" Then
            Node51_Data2.header_id = 0
        Else
            If CLng(Trim(txtHeader)) > 32767 Then
                Message ("E101")
                txtHeader.SetFocus
                Exit Sub
            Else
                Node51_Data2.header_id = CInt(Trim(txtHeader))
            End If
        End If
        
        ' 发出号码资源ID
        If Trim(txtFromNo) = "" Then
            Node51_Data2.from_id = 0
        Else
            If CLng(Trim(txtFromNo)) > 32767 Then
                Message ("E102")
                txtFromNo.SetFocus
                Exit Sub
            Else
                Node51_Data2.from_id = CInt(Trim(txtFromNo))
            End If
        End If
        
        ' LOGO文件
        If Trim(T_fax_logo) = "" Then
            Node51_Data2.fax_logo = 0
        Else
            If CLng(Trim(T_fax_logo)) > 32767 Then
                Message ("E094")
                T_fax_logo.SetFocus
                Exit Sub
            Else
                Node51_Data2.fax_logo = CInt(Trim(T_fax_logo))
            End If
        End If
        
        ' 表格格式文件
        If Trim(T_fax_format) = "" Then
            Node51_Data2.fax_format = 0
        Else
            If CLng(Trim(T_fax_format)) > 32767 Then
                Message ("E095")
                T_fax_format.SetFocus
                Exit Sub
            Else
                Node51_Data2.fax_format = CInt(Trim(T_fax_format))
            End If
        End If
        
        Node51_Data2.reserved1(0) = 0
        
        ' COM接口ID
        If Trim(T_com_iid.Text) = "" Then
            Node51_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E040")
                T_com_iid.SetFocus
                Exit Sub
            Else
                Node51_Data2.com_iid = CInt(Trim(T_com_iid))
            End If
        End If

        Node51_Data2.reserved2(0) = 0
    
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node51_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node51_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If

        ' 子节点
        If Trim(T_nd_child) = "" Then
            Node51_Data2.nd_child = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                Node51_Data2.nd_child = CInt(Trim(T_nd_child.Text))
            End If
        End If
    
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
    
        Node51_Data2.reserved3(0) = 0
   
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no.Text)
   
        f_DataChanged = False
        
    End If
    
    '' Sun added 2007-03-25
    If Node51_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
        Message "M144"
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

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
  
    ' 被访日志
    With Cb_log
        .AddItem LoadNationalResString(1178)
        .ItemData(.ListCount - 1) = 0
        For i = 1 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1179)
            .ItemData(.ListCount - 1) = i
        Next
    End With

    ' 节点超时(秒)
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
    T_vox_op.Text = Node51_Data2.vox_op
    T_fax_logo.Text = Node51_Data2.fax_logo
    T_fax_format.Text = Node51_Data2.fax_format
    txtHeader = Node51_Data2.header_id
    txtFromNo = Node51_Data2.from_id
    T_com_iid.Text = Node51_Data2.com_iid
    T_nd_parent.Text = Node51_Data2.nd_parent
    T_nd_child.Text = Node51_Data2.nd_child
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    Cb_timeout.Text = Node51_Data1.timeout
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node51_Data1.log), 0)
   
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

Private Sub T_fax_format_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_fax_format
    
End Sub

Private Sub T_fax_format_GotFocus()
    T_fax_format.SelStart = 0
    T_fax_format.SelLength = Len(T_fax_format)
End Sub

Private Sub T_fax_format_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_fax_logo_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_fax_logo

End Sub

Private Sub T_fax_logo_GotFocus()
    T_fax_logo.SelStart = 0
    T_fax_logo.SelLength = Len(T_fax_logo)
End Sub

Private Sub T_fax_logo_KeyPress(KeyAscii As Integer)
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

Private Sub T_timeout_KeyPress(KeyAscii As Integer)
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

Private Sub txtHeader_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtHeader

End Sub

Private Sub txtHeader_GotFocus()
    txtHeader.SelStart = 0
    txtHeader.SelLength = Len(txtHeader)
End Sub

Private Sub txtHeader_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtFromNo_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtFromNo

End Sub

Private Sub txtFromNo_GotFocus()
    txtFromNo.SelStart = 0
    txtFromNo.SelLength = Len(txtFromNo)
End Sub

Private Sub txtFromNo_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
