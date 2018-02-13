VERSION 5.00
Begin VB.Form frm_021 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "放音继续"
   ClientHeight    =   3225
   ClientLeft      =   3165
   ClientTop       =   2550
   ClientWidth     =   6525
   ForeColor       =   &H00000000&
   Icon            =   "frm_021.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1246"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   3402
      Picture         =   "frm_021.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "1948"
      Top             =   2841
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   885
      Left            =   3810
      TabIndex        =   32
      Tag             =   "1104"
      Top             =   2250
      Width           =   2655
      Begin VB.TextBox Txt_Description 
         Height          =   555
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   3810
      TabIndex        =   26
      Tag             =   "1164"
      Top             =   30
      Width           =   2655
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2220
         Picture         =   "frm_021.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "1145"
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2220
         Picture         =   "frm_021.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "1145"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   2
         Left            =   2220
         Picture         =   "frm_021.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "显示资源列表"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   1
         Left            =   2220
         Picture         =   "frm_021.frx":2752
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "显示资源列表"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Index           =   0
         Left            =   2220
         Picture         =   "frm_021.frx":2854
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "显示资源列表"
         Top             =   270
         Width           =   315
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
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   15
         Top             =   1680
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
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   11
         Top             =   960
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
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   13
         Top             =   1320
         Width           =   795
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
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox T_vox_succ 
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
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   9
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "子节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Tag             =   "1252"
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "COM接口ID"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1168"
         Top             =   1050
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Tag             =   "1169"
         Top             =   1410
         Width           =   705
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "前导播放语音"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Tag             =   "1250"
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "后续播放语音"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Tag             =   "1251"
         Top             =   690
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2685
      Left            =   60
      TabIndex        =   20
      Tag             =   "1136"
      Top             =   60
      Width           =   3675
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "请选择"
         Top             =   720
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2160
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "请选择"
         Top             =   1800
         Width           =   1965
      End
      Begin VB.ComboBox CB_playtype 
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
         ItemData        =   "frm_021.frx":2956
         Left            =   1560
         List            =   "frm_021.frx":2958
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "请选择"
         Top             =   1080
         Width           =   1965
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1560
         TabIndex        =   34
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "放音清空标志"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Tag             =   "1244"
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Tag             =   "1245"
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "中断按键"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Tag             =   "1249"
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Tag             =   "1248"
         Top             =   1500
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "播放类型"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Tag             =   "1247"
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   1380
      TabIndex        =   19
      Tag             =   "1144"
      Top             =   2850
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   18
      Tag             =   "1007"
      Top             =   2850
      Width           =   1035
   End
End
Attribute VB_Name = "frm_021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_021.frm
'//用途：    放音继续
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：放音继续
'//////////////////////////////////////////////////////////////////

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
        gSystem.intCurStep = 1
        Set gSystem.crlCurItem = T_vox_succ
    Case 2
        gSystem.intCurStep = 3
        Set gSystem.crlCurItem = T_com_iid
    End Select
    frmResourceList.Show vbModal

End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        Node21_Data1.reserved1 = 0
        
        ' 放音清空
        If Cb_playclear.ListIndex = -1 Then
            Node21_Data1.playclear = 0
        Else
            Node21_Data1.playclear = CByte(Cb_playclear.ItemData(Cb_playclear.ListIndex))
        End If
        
        ' 播放类型
        If cb_PlayType.ListIndex = -1 Then
            Node21_Data1.playtype = 0
        Else
            Node21_Data1.playtype = CByte(cb_PlayType.ItemData(cb_PlayType.ListIndex))
        End If
        
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node21_Data1.usevar = 0
        Else
            Node21_Data1.usevar = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        ' 按键中断
        If Cb_breakkey.ListIndex < 0 Then
            Node21_Data1.breakkey = 0
        Else
            Node21_Data1.breakkey = CByte(Cb_breakkey.ItemData(Cb_breakkey.ListIndex))
        End If

        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node21_Data1.log = 0
        Else
            Node21_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
            
        Node21_Data1.reserved2(0) = 0
        
        ' 前导播放语音
        If Trim(T_vox_pred) = "" Then
            Node21_Data2.vox_pred = 0
        Else
            If CLng(Trim(T_vox_pred)) > 32767 Then
                Message ("E090")
                T_vox_pred.SetFocus
                Exit Sub
            Else
                Node21_Data2.vox_pred = CInt(Trim(T_vox_pred.Text))
            End If
        End If
        
        ' 后续播放语音
        If Trim(T_vox_succ) = "" Then
            Node21_Data2.vox_succ = 0
        Else
            If CLng(Trim(T_vox_succ)) > 32767 Then
                Message ("E091")
                T_vox_succ.SetFocus
                Exit Sub
            Else
                Node21_Data2.vox_succ = CInt(Trim(T_vox_succ.Text))
            End If
        End If
        
        Node21_Data2.reserved1(0) = 0
        
        ' COM接口ID
        If Trim(T_com_iid.Text) = "" Then
            Node21_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E040")
                T_com_iid.SetFocus
                Exit Sub
            Else
                Node21_Data2.com_iid = CInt(Trim(T_com_iid.Text))
            End If
        End If

        ' 保留
        Node21_Data2.reserved2(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node21_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node21_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If

        ' 子节点
        If Trim(T_nd_child) = "" Then
            Node21_Data2.nd_child = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                Node21_Data2.nd_child = CInt(Trim(T_nd_child.Text))
            End If
         End If

        ' 保留
        Node21_Data2.reserved3(0) = 0
        
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
    If Node21_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
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

    ' 使用变量ID
    RefreshVariablesList Cb_usevar
     
    ' 播放类型
    With cb_PlayType
        .AddItem LoadNationalResString(1255)
        .ItemData(.ListCount - 1) = 0
        .AddItem LoadNationalResString(1256)
        .ItemData(.ListCount - 1) = 1
        .AddItem LoadNationalResString(1257)
        .ItemData(.ListCount - 1) = 2
        .AddItem LoadNationalResString(1258)
        .ItemData(.ListCount - 1) = 4
        .AddItem LoadNationalResString(1259)
        .ItemData(.ListCount - 1) = 8
        .AddItem LoadNationalResString(1599)
        .ItemData(.ListCount - 1) = 16
    End With
  
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

    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_pred.Text = Node21_Data2.vox_pred
    T_vox_succ.Text = Node21_Data2.vox_succ
    T_com_iid.Text = Node21_Data2.com_iid
    T_nd_parent.Text = Node21_Data2.nd_parent
    T_nd_child.Text = Node21_Data2.nd_child
    
    Cb_breakkey.ListIndex = SearchItemDataIndex(Cb_breakkey, CLng(Node21_Data1.breakkey), 11)
    Cb_playclear.ListIndex = SearchItemDataIndex(Cb_playclear, CLng(Node21_Data1.playclear), 0)
    cb_PlayType.ListIndex = SearchItemDataIndex(cb_PlayType, CLng(Node21_Data1.playtype), -1)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node21_Data1.log), 0)
    
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node21_Data1.usevar), 0)
 
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

Private Sub T_vox_succ_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_succ

End Sub

Private Sub T_vox_succ_GotFocus()
    T_vox_succ.SelStart = 0
    T_vox_succ.SelLength = Len(T_vox_succ)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_succ)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_succ_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
