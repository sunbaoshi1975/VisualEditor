VERSION 5.00
Begin VB.Form frm_000 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "全程转移规则"
   ClientHeight    =   4395
   ClientLeft      =   4170
   ClientTop       =   2430
   ClientWidth     =   3930
   Icon            =   "frm_000.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1135"
   Begin VB.CommandButton CommandExit 
      Caption         =   "退出(&X)"
      Height          =   315
      Left            =   2085
      TabIndex        =   15
      Tag             =   "1144"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "保存(&S)"
      Default         =   -1  'True
      Height          =   315
      Left            =   765
      TabIndex        =   14
      Tag             =   "1007"
      Top             =   3960
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1095
      Left            =   60
      TabIndex        =   24
      Tag             =   "1104"
      Top             =   3960
      Visible         =   0   'False
      Width           =   3795
      Begin VB.TextBox Txt_Description 
         Height          =   705
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   300
         Visible         =   0   'False
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      Height          =   3795
      Left            =   60
      TabIndex        =   16
      Tag             =   "1136"
      Top             =   60
      Width           =   3795
      Begin VB.TextBox txtSysHookNode 
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
         TabIndex        =   8
         Text            =   "256"
         Top             =   2550
         Width           =   1065
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3330
         Picture         =   "frm_000.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1145"
         Top             =   2580
         Width           =   315
      End
      Begin VB.CheckBox chkLogSwitchOff 
         Caption         =   "不记录节点日志"
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Tag             =   "1951"
         Top             =   3255
         Width           =   1695
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtSysDataFormatNode 
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
         TabIndex        =   6
         Text            =   "256"
         Top             =   2190
         Width           =   1035
      End
      Begin VB.CommandButton cmdShowTypedNodeList 
         Height          =   285
         Left            =   3300
         Picture         =   "frm_000.frx":11CC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "1145"
         Top             =   2220
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3300
         Picture         =   "frm_000.frx":1556
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "1145"
         Top             =   1860
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   3300
         Picture         =   "frm_000.frx":18E0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "1146"
         Top             =   2940
         Width           =   315
      End
      Begin VB.TextBox txtMainCOM 
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
         TabIndex        =   10
         Text            =   "256"
         Top             =   2910
         Width           =   1035
      End
      Begin VB.ComboBox Cb_key_root 
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
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1470
         Width           =   1365
      End
      Begin VB.ComboBox CB_key_return 
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
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1110
         Width           =   1365
      End
      Begin VB.ComboBox CB_key_repeat 
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
         ItemData        =   "frm_000.frx":19E2
         Left            =   2250
         List            =   "frm_000.frx":19E4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   750
         Width           =   1365
      End
      Begin VB.TextBox T_nd_root 
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
         TabIndex        =   4
         Text            =   "256"
         Top             =   1830
         Width           =   1035
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   270
         Width           =   525
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
         TabIndex        =   3
         Text            =   "0"
         Top             =   1830
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "挂机转节点ID"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Tag             =   "2137"
         Top             =   2610
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "缺省数据发送节点ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Tag             =   "1538"
         Top             =   2220
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "COM接口资源ID"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Tag             =   "1142"
         Top             =   2970
         Width           =   1245
      End
      Begin VB.Label Lbl_nd_root 
         AutoSize        =   -1  'True
         Caption         =   "根节点ID"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Tag             =   "1141"
         Top             =   1860
         Width           =   705
      End
      Begin VB.Label Lbl_key_repeat 
         AutoSize        =   -1  'True
         Caption         =   "重复当前节点按键"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Tag             =   "1138"
         Top             =   780
         Width           =   1440
      End
      Begin VB.Label Lbl_key_return 
         AutoSize        =   -1  'True
         Caption         =   "回上一级节点按键"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Tag             =   "1139"
         Top             =   1140
         Width           =   1440
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1980
         TabIndex        =   20
         Tag             =   "1137"
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Tag             =   "1143"
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Lbl_key_root 
         AutoSize        =   -1  'True
         Caption         =   "回到主菜单"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "1140"
         Top             =   1500
         Width           =   900
      End
   End
End
Attribute VB_Name = "frm_000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_000.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：全程转移规则
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub CB_key_repeat_Click()
    f_DataChanged = True
End Sub

Private Sub CB_key_return_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_key_root_Click()
    f_DataChanged = True
End Sub

Private Sub cmdShowNodeList_Click(Index As Integer)
    Select Case Index
    Case 0
        Set gSystem.crlCurItem = T_nd_root
    Case 1
        Set gSystem.crlCurItem = txtSysHookNode
    End Select
    frmNodeList.Show vbModal
End Sub

Private Sub cmdShowRes_Click()
    
    gSystem.intCurStep = 3
    Set gSystem.crlCurItem = txtMainCOM
    frmResourceList.Show vbModal
   
End Sub

Private Sub cmdShowTypedNodeList_Click()

    '' Sun added 2004-12-30
    Set gSystem.crlCurItem = txtSysDataFormatNode
    g_NodeListShowType = 18
    frmNodeList.Show vbModal
    g_NodeListShowType = 0

End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        'Sun modified 2001-05-24
        ' 重复当前节点按键默认为0
        If CB_key_repeat.ListIndex < 0 Then
           Node0_Data2.key_repeat = 0
        Else
            Node0_Data2.key_repeat = CByte(CB_key_repeat.ItemData(CB_key_repeat.ListIndex))
        End If
        
        ' 回到上一级菜单默认为0
        If CB_key_return.ListIndex < 0 Then
           Node0_Data2.key_return = 0
        Else
           Node0_Data2.key_return = CByte(CB_key_return.ItemData(CB_key_return.ListIndex))
        End If
        
        ' 回到主菜单默认为0
        If Cb_key_root.ListIndex < 0 Then
           Node0_Data2.key_root = 0
        Else
           Node0_Data2.key_root = CByte(Cb_key_root.ItemData(Cb_key_root.ListIndex))
        End If
        
        ' 父节点默认为0
        'If Trim(T_nd_parent) = "" Then
           Node0_Data2.nd_parent = 0
        'Else
        '   If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
        '      Message ("E035")
        '      T_nd_parent.SetFocus
        '      Exit Sub
        '   Else
        '      Node0_Data2.nd_parent = Val(T_nd_parent)
        '   End If
        'End If
        
        ' 根节点默认为256
        If Trim(T_nd_root) = "" Then
           Node0_Data2.nd_root = 256
        Else
           If Val(T_nd_root) > 32767 Or Val(T_nd_root) < 256 Then
              Message ("E034")
              T_nd_root.SetFocus
              Exit Sub
           Else
              Node0_Data2.nd_root = Val(T_nd_root.Text)
           End If
        End If
        
        ' Sun added 2004-12-30
        '' 缺省数据发送节点ID
        If Trim(txtSysDataFormatNode) = "" Then
            Node0_Data2.nd_SysSendData = 0
        Else
            Node0_Data2.nd_SysSendData = txtSysDataFormatNode
        End If
        
        ' Sun added 2012-05-07
        '' 挂机前转节点
        If Trim(txtSysHookNode) = "" Then
            Node0_Data2.nd_BeforeHookOn = 0
        Else
            Node0_Data2.nd_BeforeHookOn = txtSysHookNode
        End If
        
        ' COM接口资源ID
        If Trim(txtMainCOM) = "" Then
            Node0_Data2.MainCOM = 0
        Else
            If CLng(Trim(txtMainCOM)) > 32767 Then
                Message ("E088")
                txtMainCOM.SetFocus
                Exit Sub
            Else
                Node0_Data2.MainCOM = CInt(Trim(txtMainCOM))
            End If
        End If
        Node0_Data2.ResourceProject = gCallFlow.ResourceID
        
        '节点日志全局控制开关 <Mike 2008-2-19>
        Node0_Data2.LogSwitchOff = CByte(chkLogSwitchOff.value)
        
        ' 节点描述文字
        If Trim(Txt_Description) = "" Or IsNull(Txt_Description) Then
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description)
        End If
        
        '' Sun added 2002-12-04
        If Node0_Data2.nd_root <> gCallFlow.RootNodedID Then
            F_SwitchRootNodeDisplay Node0_Data2.nd_root
        End If
        
        '' Sun added 2007-03-25
        Node0_Data1.reserved1(0) = 0
        Node0_Data1.MajorVer = Def_CallFlow_MajorVersion
        Node0_Data1.MinorVer = Def_CallFlow_MinorVersion
        
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

'Scott modify 2001/08/30
'初始化节点000的信息
  '重复当前按钮
  F_FillPhoneKeyList CB_key_repeat, 2
  '返回上级菜单
  F_FillPhoneKeyList CB_key_return, 2
  '返回根节点
  F_FillPhoneKeyList Cb_key_root, 2

'节点ID
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
'节点编号
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
  
'重复当前按钮
    CB_key_repeat.ListIndex = SearchItemDataIndex(CB_key_repeat, CLng(Node0_Data2.key_repeat), 0)
'返回上一级菜单
    CB_key_return.ListIndex = SearchItemDataIndex(CB_key_return, CLng(Node0_Data2.key_return), 0)
'返回根节点
    Cb_key_root.ListIndex = SearchItemDataIndex(Cb_key_root, CLng(Node0_Data2.key_root), 0)
'父节点
    T_nd_parent.Text = Node0_Data2.nd_parent
'根节点
    T_nd_root.Text = Node0_Data2.nd_root
'缺省数据发送节点ID
    txtSysDataFormatNode = Node0_Data2.nd_SysSendData
      
' Sun added 2012-05-07
'' 挂机前转节点
    txtSysHookNode = Node0_Data2.nd_BeforeHookOn
      
'COM接口资源ID sun 2002-12-03
    txtMainCOM = Node0_Data2.MainCOM
    
'不记录节点日志 <0-记录(默认),1-不记录> Mike 2008-2-19
    chkLogSwitchOff.value = Node0_Data2.LogSwitchOff
        
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

Private Sub T_nd_root_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_root_GotFocus()
    T_nd_root.SelStart = 0
    T_nd_root.SelLength = Len(T_nd_root)
End Sub

Private Sub T_nd_root_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtMainCOM_Change()
    f_DataChanged = True
    
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtMainCOM
    
End Sub

Private Sub txtMainCOM_GotFocus()
    txtMainCOM.SelStart = 0
    txtMainCOM.SelLength = Len(txtMainCOM)
End Sub

Private Sub txtMainCOM_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtSysDataFormatNode_Change()
    f_DataChanged = True
End Sub

Private Sub txtSysDataFormatNode_GotFocus()
    txtSysDataFormatNode.SelStart = 0
    txtSysDataFormatNode.SelLength = Len(txtSysDataFormatNode)
End Sub

Private Sub txtSysDataFormatNode_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

' Mike 2008-2-19
Private Sub chkLogSwitchOff_Click()
    f_DataChanged = True
End Sub

Private Sub txtSysHookNode_Change()
    f_DataChanged = True
End Sub

Private Sub txtSysHookNode_GotFocus()
    txtSysHookNode.SelStart = 0
    txtSysHookNode.SelLength = Len(txtSysHookNode)
End Sub

Private Sub txtSysHookNode_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
