VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Frm_006 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "无条件转移"
   ClientHeight    =   3315
   ClientLeft      =   4845
   ClientTop       =   2745
   ClientWidth     =   3300
   Icon            =   "Frm_006.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1152"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   2892
      Picture         =   "Frm_006.frx":1242
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "1948"
      Top             =   2925
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   795
      Left            =   60
      TabIndex        =   15
      Tag             =   "1104"
      Top             =   2070
      Width           =   3165
      Begin VB.TextBox Txt_Description 
         Height          =   465
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   60
      TabIndex        =   9
      Tag             =   "1136"
      Top             =   60
      Width           =   3165
      Begin VB.TextBox txtDelay 
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
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   1
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox T_n_id 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   1920
         Picture         =   "Frm_006.frx":24B4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "1145"
         Top             =   1500
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   1920
         Picture         =   "Frm_006.frx":283E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "1145"
         Top             =   1140
         Width           =   315
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
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
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1470
         Width           =   825
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
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1110
         Width           =   825
      End
      Begin ComCtl2.UpDown updDelay 
         Height          =   360
         Left            =   1905
         TabIndex        =   2
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDelay"
         BuddyDispid     =   196625
         OrigLeft        =   1905
         OrigTop         =   720
         OrigRight       =   2145
         OrigBottom      =   1080
         Increment       =   500
         Max             =   30000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "延时"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Tag             =   "1240"
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "毫秒"
         Height          =   195
         Left            =   2280
         TabIndex        =   18
         Tag             =   "1553"
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1410
         TabIndex        =   14
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_nd_goto 
         Caption         =   "跳转节点"
         Height          =   225
         Left            =   150
         TabIndex        =   11
         Tag             =   "1151"
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Lbl_nd_parent 
         Caption         =   "父节点"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Tag             =   "1150"
         Top             =   1170
         Width           =   705
      End
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Tag             =   "1007"
      Top             =   2940
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   1335
      TabIndex        =   0
      Tag             =   "1144"
      Top             =   2940
      Width           =   1035
   End
End
Attribute VB_Name = "Frm_006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_006.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：无条件转移
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

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
        Set gSystem.crlCurItem = T_nd_goto
    End Select
    frmNodeList.Show vbModal

End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            T_nd_parent = "0"
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node6_Data2.nd_parent = Val(T_nd_parent)
            End If
        End If
        
        '延时, Sun added 2012-11-23
        If Trim(txtDelay) = "" Then
           Node6_Data1.sleep = 0
        Else
           If Val(txtDelay) > updDelay.Max Or Val(txtDelay) < updDelay.Min Then
                Message ("E061")
                txtDelay.SetFocus
                Exit Sub
           Else
                Node6_Data1.sleep = Val(txtDelay)
           End If
        End If
        
        ' 保留
        Node6_Data2.reserved1(0) = 0
        
        ' 跳转节点
        If Trim(T_nd_goto) = "" Then
            Node6_Data2.nd_goto = 0
        Else
            If (Val(T_nd_goto) > 32767 Or Val(T_nd_goto) < 256) And Val(T_nd_goto) <> 0 Then
                Message ("E110")
                Exit Sub
            Else
                Node6_Data2.nd_goto = Val(T_nd_goto)
            End If
        End If
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
       
       'Scott modeify 2001/08/30
    
        ' 节点数据整和
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
       
        ' 节点保存
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no.Text)
        
        f_DataChanged = False
        
    End If
    
    Unload Me
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub CommandExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

'节点ID
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID

'节点编号
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo

' 延时, Sun added 2012-11-23
    txtDelay = Node6_Data1.sleep

'默认跳转节点ID
    T_nd_goto.Text = Node6_Data2.nd_goto
'默认父节点ID
    T_nd_parent.Text = Node6_Data2.nd_parent
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

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtDelay_Change()
    f_DataChanged = True
End Sub

Private Sub txtDelay_GotFocus()
    txtDelay.SelStart = 0
    txtDelay.SelLength = Len(txtDelay)
End Sub

Private Sub txtDelay_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

