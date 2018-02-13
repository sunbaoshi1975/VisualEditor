VERSION 5.00
Begin VB.Form frm_023 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "放音转移"
   ClientHeight    =   3390
   ClientLeft      =   3165
   ClientTop       =   2670
   ClientWidth     =   6780
   Icon            =   "frm_023.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1279"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6372
      Picture         =   "frm_023.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "1948"
      Top             =   2931
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1275
      Left            =   3600
      TabIndex        =   18
      Tag             =   "1104"
      Top             =   1530
      Width           =   3105
      Begin VB.TextBox Txt_Description 
         Height          =   795
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   300
         Width           =   2865
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2745
      Left            =   60
      TabIndex        =   15
      Tag             =   "1136"
      Top             =   60
      Width           =   3465
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
         Top             =   2220
         Width           =   1755
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
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1830
         Width           =   1755
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
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "请选择"
         Top             =   1470
         Width           =   1755
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
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "请选择"
         Top             =   1110
         Width           =   1755
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
         Left            =   1590
         TabIndex        =   0
         ToolTipText     =   "0 - 255 秒"
         Top             =   750
         Width           =   1755
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   20
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Tag             =   "1248"
         Top             =   2310
         Width           =   885
      End
      Begin VB.Label Lbl_log 
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Tag             =   "1245"
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "放音清空标志"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Tag             =   "1244"
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1740
         TabIndex        =   22
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         Caption         =   "节点超时(秒)"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Tag             =   "1154"
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "按键中断"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Tag             =   "1264"
         Top             =   1560
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   3600
      TabIndex        =   14
      Tag             =   "1164"
      Top             =   60
      Width           =   3105
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2670
         Picture         =   "frm_023.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "1145"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2670
         Picture         =   "frm_023.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "1145"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2670
         Picture         =   "frm_023.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox T_nd_goto 
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
         TabIndex        =   9
         Top             =   960
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   7
         Top             =   600
         Width           =   1065
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
         Width           =   1065
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "转移节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Tag             =   "1280"
         Top             =   1050
         Width           =   885
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Tag             =   "1169"
         Top             =   690
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "默认播放语音"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Tag             =   "1676"
         Top             =   330
         Width           =   1080
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   3570
      TabIndex        =   13
      Tag             =   "1144"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   2100
      TabIndex        =   12
      Tag             =   "1007"
      Top             =   2940
      Width           =   1065
   End
End
Attribute VB_Name = "frm_023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_023.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：放音转移
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
    
'    If Cb_usevar.ListIndex > 0 Then
'        T_vox_play.Enabled = False
'        cmdShowRes.Enabled = False
'    Else
'        T_vox_play.Enabled = True
'        cmdShowRes.Enabled = True
'    End If
    
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
        Set gSystem.crlCurItem = T_nd_goto
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdShowRes_Click()
    
    gSystem.intCurStep = 1
    Set gSystem.crlCurItem = T_vox_play
    frmResourceList.Show vbModal

End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then

        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node23_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node23_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        ' 放音清空
        If Cb_playclear.ListIndex = -1 Then
            Node23_Data1.playclear = 0
        Else
            Node23_Data1.playclear = CByte(Cb_playclear.ItemData(Cb_playclear.ListIndex))
        End If
        
        ' 按键中断
        If Cb_breakkey.ListIndex < 0 Then
            Node23_Data1.breakkey = 0
        Else
            Node23_Data1.breakkey = CByte(Cb_breakkey.ItemData(Cb_breakkey.ListIndex))
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node23_Data1.log = 0
        Else
            Node23_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        Node23_Data1.reserved1(0) = 0
        Node23_Data1.reserved2(0) = 0
        
        '' Sun added 2006-02-10
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node23_Data1.var_play = 0
        Else
            Node23_Data1.var_play = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        ' 播放语音
        If Trim(T_vox_play) = "" Then
            Node23_Data2.vox_play = 0
        Else
            If CLng(Trim(T_vox_play)) > 32767 Then
                Message ("E076")
                T_vox_play.SetFocus
                Exit Sub
            Else
                Node23_Data2.vox_play = CInt(Trim(T_vox_play.Text))
            End If
        End If
        
        ' 保留
        Node23_Data2.reserved1(0) = 0
        Node23_Data2.reserved2(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node23_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node23_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
        
        ' 转移节点ID
        If Trim(T_nd_goto) = "" Then
           Node23_Data2.nd_goto = 0
        Else
           If (Val(T_nd_goto) > 32767 Or Val(T_nd_goto) < 256) And Val(T_nd_goto) <> 0 Then
              Message ("E110")
              T_nd_goto.SetFocus
              Exit Sub
           Else
              Node23_Data2.nd_goto = CLng(Trim(T_nd_goto))
           End If
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
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
    
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

    ' 使用变量ID
    RefreshVariablesList Cb_usevar

    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_play.Text = Node23_Data2.vox_play
    T_nd_parent.Text = Node23_Data2.nd_parent
    T_nd_goto.Text = Node23_Data2.nd_goto
    
    Cb_timeout.Text = Node23_Data1.timeout
    Cb_playclear.ListIndex = SearchItemDataIndex(Cb_playclear, CLng(Node23_Data1.playclear), 0)
    Cb_breakkey.ListIndex = SearchItemDataIndex(Cb_breakkey, CLng(Node23_Data1.breakkey), 11)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node23_Data1.log), 0)
 
    '' Sun added 2006-02-10
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node23_Data1.var_play), 0)
'    If Node23_Data1.var_play > 0 Then
'        T_vox_play.Enabled = False
'        cmdShowRes.Enabled = False
'    Else
'        T_vox_play.Enabled = True
'        cmdShowRes.Enabled = True
'    End If

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

Private Sub T_nd_goto_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_goto_GotFocus()
    T_nd_goto.SelStart = 0
    T_nd_goto.SelLength = Len(T_nd_goto)
End Sub

Private Sub T_nd_goto_KeyPress(KeyAscii As Integer)
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

Private Sub T_vox_play_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_play

End Sub

Private Sub T_vox_play_GotFocus()
    T_vox_play.SelStart = 0
    T_vox_play.SelLength = Len(T_vox_play)

    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_play)
    Call SoundResourceIDChanged

End Sub

Private Sub T_vox_play_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
