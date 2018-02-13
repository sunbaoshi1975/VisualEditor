VERSION 5.00
Begin VB.Form frm_020 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "放音挂机"
   ClientHeight    =   3960
   ClientLeft      =   4665
   ClientTop       =   2790
   ClientWidth     =   3270
   ForeColor       =   &H00000000&
   Icon            =   "frm_020.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1243"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   2862
      Picture         =   "frm_020.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "1948"
      Top             =   3561
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1065
      Left            =   60
      TabIndex        =   18
      Tag             =   "1104"
      Top             =   2400
      Width           =   3135
      Begin VB.TextBox Txt_Description 
         Height          =   645
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   60
      TabIndex        =   9
      Tag             =   "1136"
      Top             =   60
      Width           =   3135
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Left            =   2700
         Picture         =   "frm_020.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "1145"
         Top             =   1830
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2700
         Picture         =   "frm_020.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "1146"
         Top             =   1470
         Width           =   315
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
         TabIndex        =   15
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   14
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
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1080
         Width           =   1605
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
         TabIndex        =   4
         Top             =   1800
         Width           =   1275
      End
      Begin VB.ComboBox CB_playclear 
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
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "请选择"
         Top             =   720
         Width           =   1605
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
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1380
         TabIndex        =   16
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Tag             =   "1169"
         Top             =   1890
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "放音清空标志"
         Height          =   180
         Left            =   180
         TabIndex        =   12
         Tag             =   "1244"
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Tag             =   "1245"
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "播放语音"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Tag             =   "1100"
         Top             =   1500
         Width           =   720
      End
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Tag             =   "1007"
      Top             =   3570
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   1275
      TabIndex        =   8
      Tag             =   "1144"
      Top             =   3570
      Width           =   1035
   End
End
Attribute VB_Name = "frm_020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_020.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：放音挂机
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_playclear_Click()
    f_DataChanged = True
End Sub

'Mike added this event @ 2008-1-30
Private Sub cmdNodeTag_Click()
    frmNodeTagEdit.iNodeID = CInt(T_n_id)
    frmNodeTagEdit.byNodeNo = CByte(T_n_no.Text)
    frmNodeTagEdit.Show vbModal
End Sub

Private Sub cmdShowNodeList_Click()
    
    Set gSystem.crlCurItem = T_nd_parent
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
        
        Node20_Data1.reserved1 = 0
        
        ' 放音清空
        If Cb_playclear.ListIndex = -1 Then
            Node20_Data1.playclear = 0
        Else
            Node20_Data1.playclear = CByte(Cb_playclear.ItemData(Cb_playclear.ListIndex))
        End If
        
        Node20_Data1.reserved2(0) = 0
        
        ' 被访问日志
        If Cb_log.ListIndex = -1 Then
            Node20_Data1.log = 0
        Else
            Node20_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        Node20_Data1.reserved3(0) = 0
        
        ' 播放语音
        If Trim(T_vox_play) = "" Then
            Node20_Data2.vox_play = 0
        Else
            If CLng(Trim(T_vox_play)) > 32767 Then
                Message ("E076")
                T_vox_play.SetFocus
                Exit Sub
            Else
                Node20_Data2.vox_play = CInt(Trim(T_vox_play.Text))
            End If
        End If
        Node20_Data2.reserved1(0) = 0
        '父节点
        If Trim(T_nd_parent) = "" Then
            Node20_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node20_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
        Node20_Data2.reserved2(0) = 0
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

Private Sub CommandExit_Click()
   Unload Me
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
   
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_play.Text = Node20_Data2.vox_play
    T_nd_parent.Text = Node20_Data2.nd_parent
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node20_Data1.log), 0)
    Cb_playclear.ListIndex = SearchItemDataIndex(Cb_playclear, CLng(Node20_Data1.playclear), 0)
      
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

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
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
