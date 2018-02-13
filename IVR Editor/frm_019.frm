VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm_019 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "无操作"
   ClientHeight    =   3795
   ClientLeft      =   4545
   ClientTop       =   2790
   ClientWidth     =   2595
   Icon            =   "frm_019.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1239"
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   855
      Left            =   60
      TabIndex        =   10
      Tag             =   "1104"
      Top             =   2400
      Width           =   2445
      Begin VB.TextBox Txt_Description 
         Height          =   555
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2265
      Left            =   45
      TabIndex        =   0
      Tag             =   "1136"
      Top             =   60
      Width           =   2475
      Begin VB.CommandButton cmdNodeTag 
         Height          =   333
         Left            =   1992
         Picture         =   "frm_019.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "1948"
         Top             =   1761
         Width           =   333
      End
      Begin VB.CheckBox chkLeaveQueue 
         Caption         =   "离开队列"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Tag             =   "1817"
         Top             =   1800
         Width           =   1095
      End
      Begin ComCtl2.UpDown updDelay 
         Height          =   360
         Left            =   1695
         TabIndex        =   7
         Top             =   900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDelay"
         BuddyDispid     =   196614
         OrigLeft        =   1560
         OrigTop         =   900
         OrigRight       =   1800
         OrigBottom      =   1275
         Increment       =   100
         Max             =   3000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
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
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   6
         Top             =   900
         Width           =   525
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Left            =   2010
         Picture         =   "frm_019.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1145"
         Top             =   1320
         Width           =   315
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
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1290
         Width           =   825
      End
      Begin VB.TextBox T_n_no 
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox T_n_id 
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "毫秒"
         Height          =   195
         Left            =   2040
         TabIndex        =   15
         Tag             =   "1553"
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "延时"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Tag             =   "1240"
         Top             =   990
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Tag             =   "1143"
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Tag             =   "1137"
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "父节点"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Tag             =   "1150"
         Top             =   1350
         Width           =   540
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   165
      TabIndex        =   12
      Tag             =   "1007"
      Top             =   3330
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   1365
      TabIndex        =   13
      Tag             =   "1144"
      Top             =   3330
      Width           =   1035
   End
End
Attribute VB_Name = "frm_019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_019.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：无操作
'//////////////////////////////////////////////////////////////////
'//修改时间 : July,10,07
'//修改内容 : SData1中添加1项属性(leavequeue)
'//修改人   : Michael
'//修改版本 : V1.1
'//////////////////////////////////////////////////////////////////


Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean



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

Public Sub Command1_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        '保留
        Node19_Data1.reserved1(0) = 0
        '保留
        Node19_Data2.reserved1(0) = 0
        
        '延时
        If Trim(txtDelay) = "" Then
           Node19_Data2.delaytime = 0
        Else
           If Val(txtDelay) > updDelay.Max Or Val(txtDelay) < updDelay.Min Then
                Message ("E061")
                txtDelay.SetFocus
                Exit Sub
           Else
                Node19_Data2.delaytime = Val(txtDelay)
           End If
        End If
        
        '父节点
        If Trim(T_nd_parent) = "" Then
           Node19_Data2.nd_parent = 0
        Else
           If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
           Else
                Node19_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
           End If
        End If
        
        '保留
        Node19_Data2.reserved2(0) = 0
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        'Michael Added @Jul,10,07
        Node19_Data1.leavequeue = chkLeaveQueue.value
   
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

    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    T_nd_parent.Text = Node19_Data2.nd_parent
    txtDelay = Node19_Data2.delaytime
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    
    'Michael Added @ Jul,10,07
    chkLeaveQueue.value = Node19_Data1.leavequeue
   
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

Private Sub Command2_Click()
   Unload Me
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

'Michael Add @ Jul,10,07
Private Sub chkLeaveQueue_Click()
    f_DataChanged = True
End Sub
