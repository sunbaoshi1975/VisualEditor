VERSION 5.00
Begin VB.Form frm_069 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "转虚拟分机"
   ClientHeight    =   3225
   ClientLeft      =   4800
   ClientTop       =   2670
   ClientWidth     =   6945
   Icon            =   "frm_069.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1333"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6552
      Picture         =   "frm_069.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "1948"
      Top             =   2745
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1035
      Index           =   1
      Left            =   3750
      TabIndex        =   25
      Tag             =   "1104"
      Top             =   1590
      Width           =   3135
      Begin VB.TextBox Txt_Description 
         Height          =   645
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   3750
      TabIndex        =   18
      Tag             =   "1164"
      Top             =   60
      Width           =   3135
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2700
         Picture         =   "frm_069.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "1145"
         Top             =   990
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2700
         Picture         =   "frm_069.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1145"
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2700
         Picture         =   "frm_069.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
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
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   6
         Top             =   240
         Width           =   1185
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
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   10
         Top             =   960
         Width           =   1185
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
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   8
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "操作提示语音"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   26
         Tag             =   "1266"
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "子节点ID"
         Height          =   225
         Left            =   180
         TabIndex        =   20
         Tag             =   "1252"
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "父节点ID"
         Height          =   225
         Left            =   180
         TabIndex        =   19
         Tag             =   "1169"
         Top             =   690
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3045
      Left            =   60
      TabIndex        =   15
      Tag             =   "1136"
      Top             =   60
      Width           =   3615
      Begin VB.ComboBox cmbTryInterval 
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
         Left            =   1500
         TabIndex        =   5
         ToolTipText     =   "0 - 255 秒"
         Top             =   2520
         Width           =   1965
      End
      Begin VB.ComboBox Cb_MaxTry 
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
         ItemData        =   "frm_069.frx":2752
         Left            =   1500
         List            =   "frm_069.frx":2754
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   2160
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   1965
      End
      Begin VB.ComboBox cb_SwitchType 
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   720
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   1800
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   22
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
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox T_vagency 
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
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "尝试间隔(毫秒)"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Tag             =   "2136"
         Top             =   2610
         Width           =   1170
      End
      Begin VB.Label Label3 
         Caption         =   "转接尝试次数"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Tag             =   "2135"
         Top             =   2250
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Tag             =   "1248"
         Top             =   1155
         Width           =   885
      End
      Begin VB.Label Label11 
         Caption         =   "转接方式"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Tag             =   "1302"
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1770
         TabIndex        =   23
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label11 
         Caption         =   "虚拟分机号"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Tag             =   "1334"
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "被访问日志"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Tag             =   "1159"
         Top             =   1890
         Width           =   945
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Tag             =   "1144"
      Top             =   2730
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   3750
      TabIndex        =   13
      Tag             =   "1007"
      Top             =   2730
      Width           =   1035
   End
End
Attribute VB_Name = "frm_069"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_069.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：转虚拟分机
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_MaxTry_Click()
    f_DataChanged = True
End Sub

Private Sub CB_switchtype_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_usevar_Click()
    f_DataChanged = True
    
    If Cb_usevar.ListIndex > 0 Then
        T_vagency.Enabled = False
    Else
        T_vagency.Enabled = True
    End If
    
End Sub

Private Sub cmbTryInterval_Change()
    f_DataChanged = True
End Sub

Private Sub cmbTryInterval_Click()
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

Private Sub cmdShowRes_Click()

    gSystem.intCurStep = 1
    Set gSystem.crlCurItem = T_vox_op
    frmResourceList.Show vbModal

End Sub

Public Sub Command1_Click()
On Error Resume Next

    If f_DataChanged Then

        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 保留
        Node69_Data1.reserved1(0) = 0

        ' 转接方式
        If cb_SwitchType.ListIndex < 0 Then
            Node69_Data1.switchtype = 0
        Else
            Node69_Data1.switchtype = CByte(cb_SwitchType.ItemData(cb_SwitchType.ListIndex))
        End If
        
        ' 虚拟分机
        If Len(Trim(T_vagency)) < 1 Then
           Node69_Data1.vagency = 0
        Else
           If Val(T_vagency) > 2147483647 Then
              Message ("E114")
              Exit Sub
           Else
              Node69_Data1.vagency = Val(T_vagency.Text)
           End If
        End If
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node69_Data1.log = 0
        Else
            Node69_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        ' Sun added 2012-04-18
        ' 最大转接尝试次数
        If Cb_MaxTry.ListIndex < 0 Then
            Node69_Data2.maxtry = 3
        Else
            Node69_Data2.maxtry = CByte(Cb_MaxTry.ItemData(Cb_MaxTry.ListIndex))
        End If
        
        ' Sun added 2012-04-18
        ' 尝试间隔(毫秒)
        If Trim(cmbTryInterval) = "" Then
            Node69_Data2.tryinterval = 1000
        Else
            If Val(cmbTryInterval.Text) < 0 Or Val(cmbTryInterval.Text) > 30000 Then
                Node69_Data2.tryinterval = 1000
            Else
                Node69_Data2.tryinterval = CInt(Val(cmbTryInterval.Text))
            End If
        End If
        
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node69_Data1.usevar = 0
        Else
            Node69_Data1.usevar = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If

        Node69_Data1.reserved3(0) = 0
        
        ' 操作语音提示
        If Trim(T_vox_op) = "" Then
            Node69_Data2.vox_op = 0
        Else
            If Val(Trim(T_vox_op)) > 32767 Then
                Message ("E092")
                Exit Sub
            Else
                Node69_Data2.vox_op = Val(T_vox_op.Text)
            End If
        End If
        
        Node69_Data2.reserved1(0) = 0
        
        '父节点
        If Trim(T_nd_parent) = "" Then
            Node69_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node69_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If

        ' 子节点
        If Trim(T_nd_child) = "" Then
            Node69_Data2.nd_child = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                Node69_Data2.nd_child = CInt(Trim(T_nd_child.Text))
            End If
        End If
        
        Node69_Data2.reserved2(0) = 0
         
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

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
  
    ' 转接方式
    With cb_SwitchType
        .Clear
        .AddItem LoadNationalResString(1335)
        .ItemData(.ListCount - 1) = 0
        .AddItem LoadNationalResString(1336)
        .ItemData(.ListCount - 1) = 1
    End With
    
    ' 使用变量ID
    RefreshVariablesList Cb_usevar
    
    ' 被访日志
    With Cb_log
        .AddItem LoadNationalResString(1178)
        .ItemData(.ListCount - 1) = 0
        For i = 1 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1179)
            .ItemData(.ListCount - 1) = i
        Next
    End With

    '' Sun added 2012-04-18
    ' 最大转接尝试次数
    With Cb_MaxTry
        For i = 0 To 10
            .AddItem Trim(Str(i))
            .ItemData(.ListCount - 1) = i
        Next
    End With

    '' Sun added 2012-04-18
    ' 尝试间隔(毫秒)
    With cmbTryInterval
        .AddItem "500"
        .AddItem "800"
        .AddItem "1000"
        .AddItem "1200"
        .AddItem "1500"
        .AddItem "1800"
        .AddItem "2000"
        .AddItem "2500"
        .AddItem "3000"
    End With
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_vox_op.Text = Node69_Data2.vox_op
    T_nd_parent.Text = Node69_Data2.nd_parent
    T_nd_child.Text = Node69_Data2.nd_child
    T_vagency.Text = Node69_Data1.vagency
    cb_SwitchType.ListIndex = SearchItemDataIndex(cb_SwitchType, CLng(Node69_Data1.switchtype), 0)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node69_Data1.log), 0)
    
    '' Sun added 2012-04-18
    Cb_MaxTry.ListIndex = SearchItemDataIndex(Cb_MaxTry, CLng(Node69_Data2.maxtry), 3)
    cmbTryInterval = Node69_Data2.tryinterval
    
    '' Sun added 2005-08-05
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node69_Data1.usevar), 0)
    If Node69_Data1.usevar = 0 Then
        T_vagency.Enabled = True
    Else
        T_vagency.Enabled = False
    End If
        
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me
End Sub

Private Sub Command2_Click()
   Unload Me
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

Private Sub T_vagency_Change()
    f_DataChanged = True
End Sub

Private Sub T_vagency_GotFocus()
    T_vagency.SelStart = 0
    T_vagency.SelLength = Len(T_vagency)
End Sub

Private Sub T_vagency_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_vox_op_Click()
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
