VERSION 5.00
Begin VB.Form frm_071 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查询座席状态"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frm_071.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1633"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   7122
      Picture         =   "frm_071.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "1948"
      Top             =   4341
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   855
      Index           =   1
      Left            =   60
      TabIndex        =   22
      Tag             =   "1104"
      Top             =   3330
      Width           =   7395
      Begin VB.TextBox Txt_Description 
         Height          =   525
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   3690
      TabIndex        =   24
      Tag             =   "1144"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   23
      Tag             =   "1007"
      Top             =   4320
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3195
      Left            =   60
      TabIndex        =   20
      Tag             =   "1136"
      Top             =   60
      Width           =   7395
      Begin VB.CheckBox chkPos 
         Caption         =   "座席状态"
         Height          =   225
         Left            =   780
         TabIndex        =   8
         Tag             =   "1636"
         Top             =   1740
         Width           =   1245
      End
      Begin VB.ComboBox cmbSymbol 
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
         ItemData        =   "frm_071.frx":15FC
         Left            =   2160
         List            =   "frm_071.frx":1612
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox Cb_param 
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
         Left            =   3330
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cb_Condition 
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
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1230
         Width           =   885
      End
      Begin VB.CheckBox chkDN 
         Caption         =   "分机状态"
         Height          =   225
         Left            =   780
         TabIndex        =   4
         Tag             =   "1635"
         Top             =   1320
         Width           =   1245
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
         Left            =   5730
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   240
         Width           =   1515
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
         TabIndex        =   26
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
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   12
         Top             =   2160
         Width           =   1065
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
         Left            =   5820
         MaxLength       =   6
         TabIndex        =   14
         Top             =   2160
         Width           =   1065
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3240
         Picture         =   "frm_071.frx":162E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "显示节点列表"
         Top             =   2190
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   6900
         Picture         =   "frm_071.frx":19B8
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "显示节点列表"
         Top             =   2190
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3210
         Picture         =   "frm_071.frx":1D42
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "显示节点列表"
         Top             =   2670
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   16
         Top             =   2640
         Width           =   1065
      End
      Begin VB.ComboBox cmbSymbol 
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
         ItemData        =   "frm_071.frx":20CC
         Left            =   2160
         List            =   "frm_071.frx":20E2
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1230
         Width           =   1095
      End
      Begin VB.ComboBox Cb_param 
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
         Left            =   3330
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1230
         Width           =   1335
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   6900
         Picture         =   "frm_071.frx":20FE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "1145"
         Top             =   2670
         Width           =   315
      End
      Begin VB.TextBox txtNdNoAnswer 
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
         Left            =   5820
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2640
         Width           =   1065
      End
      Begin VB.TextBox txtAgentID 
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
         Left            =   2580
         MaxLength       =   6
         TabIndex        =   2
         Top             =   750
         Width           =   1005
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
         Left            =   6180
         TabIndex        =   11
         ToolTipText     =   "0 - 255 秒"
         Top             =   1680
         Width           =   1035
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   750
         Width           =   1965
      End
      Begin VB.ComboBox Cb_Object 
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
         ItemData        =   "frm_071.frx":2488
         Left            =   780
         List            =   "frm_071.frx":248A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "查询"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Tag             =   "1656"
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1380
         TabIndex        =   35
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "被访问日志"
         Height          =   195
         Index           =   0
         Left            =   4470
         TabIndex        =   34
         Tag             =   "1159"
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "否则，转节点ID"
         Height          =   195
         Left            =   4200
         TabIndex        =   33
         Tag             =   "1532"
         Top             =   2250
         Width           =   1365
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "则，转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Tag             =   "1531"
         Top             =   2220
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Tag             =   "1169"
         Top             =   2700
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "如果"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1634"
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答转节点"
         Height          =   180
         Index           =   11
         Left            =   4200
         TabIndex        =   29
         Tag             =   "1313"
         Top             =   2730
         Width           =   1080
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "通信超时(秒)"
         Height          =   195
         Index           =   0
         Left            =   4950
         TabIndex        =   28
         Tag             =   "1570"
         Top             =   1740
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Index           =   1
         Left            =   4050
         TabIndex        =   27
         Tag             =   "1248"
         Top             =   810
         Width           =   885
      End
   End
End
Attribute VB_Name = "frm_071"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_071.frm
'//用途：    查询座席状态
'//作者:     Tony Sun
'//创建日期：2005/08/15
'//修改日期：
'//文件描述：查询座席状态
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub cb_Condition_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Object_Click()
    f_DataChanged = True
    
    If Cb_usevar.ListIndex > 0 Or Cb_Object.ListIndex = 1 Then
        txtAgentID.Enabled = False
    Else
        txtAgentID.Enabled = True
    End If
    
End Sub

Private Sub Cb_param_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_usevar_Click()
    f_DataChanged = True
    
    If Cb_usevar.ListIndex > 0 Or Cb_Object.ListIndex = 0 Then
        txtAgentID.Enabled = False
    Else
        txtAgentID.Enabled = True
    End If
    
End Sub

Private Sub chkDN_Click()
    f_DataChanged = True
    
    If chkDN.value = vbChecked Then
        cmbSymbol(0).Enabled = True
        Cb_param(0).Enabled = True
    Else
        cmbSymbol(0).Enabled = False
        Cb_param(0).Enabled = False
    End If
    
End Sub

Private Sub chkPos_Click()
    f_DataChanged = True

    If chkPos.value = vbChecked Then
        cmbSymbol(1).Enabled = True
        Cb_param(1).Enabled = True
    Else
        cmbSymbol(1).Enabled = False
        Cb_param(1).Enabled = False
    End If

End Sub

Private Sub cmbSymbol_Click(Index As Integer)
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
    Case 3
        Set gSystem.crlCurItem = txtNdNoAnswer
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then

        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node71_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node71_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        Node71_Data1.reserved1(0) = 0
        
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node71_Data1.usevar = 0
        Else
            Node71_Data1.usevar = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        ' 查询对象
        If Cb_Object.ListIndex < 0 Then
            Node71_Data1.querytype = 0
        Else
            Node71_Data1.querytype = CByte(Cb_Object.ItemData(Cb_Object.ListIndex))
        End If
        
        ' 参数
        If Cb_param(0).ListIndex < 0 Then
            Node71_Data1.dn_status = 0
        Else
            Node71_Data1.dn_status = CByte(Cb_param(0).ItemData(Cb_param(0).ListIndex))
        End If
        If Cb_param(1).ListIndex < 0 Then
            Node71_Data1.pos_status = 0
        Else
            Node71_Data1.pos_status = CByte(Cb_param(1).ItemData(Cb_param(1).ListIndex))
        End If
        
        ' 逻辑运算符
        If cmbSymbol(0).ListIndex < 0 Then
            Node71_Data1.dn_logic = 0
        Else
            Node71_Data1.dn_logic = CByte(cmbSymbol(0).ItemData(cmbSymbol(0).ListIndex))
        End If
        If cmbSymbol(1).ListIndex < 0 Then
            Node71_Data1.pos_logic = 0
        Else
            Node71_Data1.pos_logic = CByte(cmbSymbol(1).ItemData(cmbSymbol(1).ListIndex))
        End If

        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node71_Data1.log = 0
        Else
            Node71_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        ' Agent ID
        Node71_Data2.agentid = CLng(Val(txtAgentID))
        
        ' 条件
        If chkDN.value = vbChecked And chkPos.value = vbChecked Then
            If cb_Condition.ListIndex < 0 Then
                Node71_Data1.conditions = DEF_NODE071_CONDITION_BOTH
            Else
                If cb_Condition.ItemData(cb_Condition.ListIndex) = 0 Then
                    Node71_Data1.conditions = DEF_NODE071_CONDITION_BOTH
                Else
                    Node71_Data1.conditions = DEF_NODE071_CONDITION_EITHER
                End If
            End If
        ElseIf chkDN.value = vbChecked And chkPos.value = vbUnchecked Then
            Node71_Data1.conditions = DEF_NODE071_CONDITION_FIRST
        ElseIf chkDN.value = vbUnchecked And chkPos.value = vbChecked Then
            Node71_Data1.conditions = DEF_NODE071_CONDITION_SECOND
        Else
            Node71_Data1.conditions = DEF_NODE071_CONDITION_NONE
        End If
        
        Node71_Data2.reserved1(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node71_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node71_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
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
        If Node71_Data2.nd_yes <> lv_nNewNode Then
            Node71_Data2.nd_yes = lv_nNewNode
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
        If Node71_Data2.nd_no <> lv_nNewNode Then
            Node71_Data2.nd_no = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
        End If
        
        ' 无应答转节点
        If Trim(txtNdNoAnswer) = "" Then
           lv_nNewNode = 0
        Else
            If (Val(txtNdNoAnswer) > 32767 Or Val(txtNdNoAnswer) < 256) And Val(txtNdNoAnswer) <> 0 Then
                Message ("E121")
                txtNdNoAnswer.SetFocus
                Exit Sub
            Else
                lv_nNewNode = CInt(Trim(txtNdNoAnswer.Text))
            End If
        End If
        
        '' Sun added 2007-03-25
        If Node71_Data2.nd_fail <> lv_nNewNode Then
            Node71_Data2.nd_fail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 255)
        End If
        
        Node71_Data2.reserved2(0) = 0
        
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

Private Sub CommandExit_Click()
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

    ' 通信超时(秒)
    With Cb_timeout
        .Clear
        For i = 0 To 120 Step 5
            .AddItem Trim(Str(i))
        Next
    End With
    
    ' 使用变量ID
    RefreshVariablesList Cb_usevar
    
    ' 查询对象
    With Cb_Object
        '' Agent ID
        .AddItem LoadNationalResString(1631)
        .ItemData(.ListCount - 1) = 0
        '' User ID
        .AddItem LoadNationalResString(1632)
        .ItemData(.ListCount - 1) = 1
        
        '' Sun added 2007-04-12
        '' 座席工号
        .AddItem LoadNationalResString(1741)
        .ItemData(.ListCount - 1) = 2
        
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

    ' 逻辑运算符
    For i = DEF_NODE016_LOGIC_EQUE To DEF_NODE016_LOGIC_NE
        cmbSymbol(0).ItemData(i) = i
        cmbSymbol(1).ItemData(i) = i
    Next

    ' 参数 - DN状态
    With Cb_param(0)
        For i = 0 To 7
            .AddItem LoadNationalResString(1637 + i)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
    ' 参数 - POS状态
    With Cb_param(1)
        For i = 0 To 7
            .AddItem LoadNationalResString(1645 + i)
            .ItemData(.ListCount - 1) = i
        Next
    End With
   
    ' 条件
    With cb_Condition
        '' AND
        .AddItem LoadNationalResString(1654)
        .ItemData(.ListCount - 1) = 0
        '' OR
        .AddItem LoadNationalResString(1655)
        .ItemData(.ListCount - 1) = 1
    End With
    
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Cb_timeout.Text = Node71_Data1.timeout
    
    txtAgentID.Text = Node71_Data2.agentid
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_nd_parent.Text = Node71_Data2.nd_parent
    T_nd_succeed.Text = Node71_Data2.nd_yes
    T_nd_fail.Text = Node71_Data2.nd_no
    txtNdNoAnswer.Text = Node71_Data2.nd_fail

    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node71_Data1.log), 0)
    cmbSymbol(0).ListIndex = SearchItemDataIndex(cmbSymbol(0), CLng(Node71_Data1.dn_logic), 0)
    cmbSymbol(1).ListIndex = SearchItemDataIndex(cmbSymbol(1), CLng(Node71_Data1.pos_logic), 0)

    Cb_param(0).ListIndex = SearchItemDataIndex(Cb_param(0), CLng(Node71_Data1.dn_status), 0)
    Cb_param(1).ListIndex = SearchItemDataIndex(Cb_param(1), CLng(Node71_Data1.pos_status), 0)
    
    cb_Condition.ListIndex = SearchItemDataIndex(cb_Condition, IIf(Node71_Data1.conditions = DEF_NODE071_CONDITION_EITHER, 1, 0), 0)
    Select Case Node71_Data1.conditions
    Case DEF_NODE071_CONDITION_NONE
        chkDN.value = vbUnchecked
        chkPos.value = vbUnchecked
    Case DEF_NODE071_CONDITION_FIRST
        chkDN.value = vbChecked
        chkPos.value = vbUnchecked
    Case DEF_NODE071_CONDITION_SECOND
        chkDN.value = vbUnchecked
        chkPos.value = vbChecked
    Case DEF_NODE071_CONDITION_BOTH
        chkDN.value = vbChecked
        chkPos.value = vbChecked
    Case DEF_NODE071_CONDITION_EITHER
        chkDN.value = vbChecked
        chkPos.value = vbChecked
    End Select
    
    If chkDN.value = vbChecked Then
        cmbSymbol(0).Enabled = True
        Cb_param(0).Enabled = True
    Else
        cmbSymbol(0).Enabled = False
        Cb_param(0).Enabled = False
    End If
    
    If chkPos.value = vbChecked Then
        cmbSymbol(1).Enabled = True
        Cb_param(1).Enabled = True
    Else
        cmbSymbol(1).Enabled = False
        Cb_param(1).Enabled = False
    End If
    
    Cb_Object.ListIndex = SearchItemDataIndex(Cb_Object, CLng(Node71_Data1.querytype), 0)
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node71_Data1.usevar), 0)
    
    Cb_Object_Click
    
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me

End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
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

Private Sub txtAgentID_Change()
    f_DataChanged = True
End Sub

Private Sub txtAgentID_GotFocus()
    txtAgentID.SelStart = 0
    txtAgentID.SelLength = Len(txtAgentID)
End Sub

Private Sub txtAgentID_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtNdNoAnswer_Change()
    f_DataChanged = True
End Sub

Private Sub txtNdNoAnswer_GotFocus()
    txtNdNoAnswer.SelStart = 0
    txtNdNoAnswer.SelLength = Len(txtNdNoAnswer)
End Sub

Private Sub txtNdNoAnswer_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

