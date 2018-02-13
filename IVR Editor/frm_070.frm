VERSION 5.00
Begin VB.Form frm_070 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查询路由点"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frm_070.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1600"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   7122
      Picture         =   "frm_070.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "1948"
      Top             =   4371
      Width           =   333
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   22
      Tag             =   "1007"
      Top             =   4350
      Width           =   1035
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   3690
      TabIndex        =   23
      Tag             =   "1144"
      Top             =   4350
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3195
      Left            =   60
      TabIndex        =   19
      Tag             =   "1136"
      Top             =   60
      Width           =   7395
      Begin VB.ComboBox cb_VarResult 
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
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   750
         Width           =   1725
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
         ItemData        =   "frm_070.frx":157C
         Left            =   180
         List            =   "frm_070.frx":157E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
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
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   1725
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
         Left            =   4200
         TabIndex        =   0
         ToolTipText     =   "0 - 255 秒"
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox txtRoutePoint 
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
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1200
         Width           =   1065
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
         TabIndex        =   17
         Top             =   2640
         Width           =   1065
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   6900
         Picture         =   "frm_070.frx":1580
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "1145"
         Top             =   2670
         Width           =   315
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
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtDefaultValue 
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
         Left            =   3120
         MaxLength       =   48
         TabIndex        =   8
         Top             =   1680
         Width           =   945
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
         ItemData        =   "frm_070.frx":190A
         Left            =   1980
         List            =   "frm_070.frx":1920
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
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
         Left            =   5820
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1680
         Width           =   1065
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   6900
         Picture         =   "frm_070.frx":193C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "1146"
         Top             =   1710
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
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   15
         Top             =   2640
         Width           =   1065
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3060
         Picture         =   "frm_070.frx":1A3E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "显示节点列表"
         Top             =   2670
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   6900
         Picture         =   "frm_070.frx":1DC8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "显示节点列表"
         Top             =   2190
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3060
         Picture         =   "frm_070.frx":2152
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "显示节点列表"
         Top             =   2190
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
         Left            =   5820
         MaxLength       =   6
         TabIndex        =   13
         Top             =   2160
         Width           =   1065
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
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2160
         Width           =   1065
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
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
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "结果记录变量"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   37
         Tag             =   "1740"
         Top             =   825
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "使用变量ID"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   36
         Tag             =   "1248"
         Top             =   825
         Width           =   885
      End
      Begin VB.Label Lbl_timeout 
         AutoSize        =   -1  'True
         Caption         =   "通信超时(秒)"
         Height          =   195
         Index           =   0
         Left            =   2970
         TabIndex        =   35
         Tag             =   "1570"
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "无应答转节点"
         Height          =   180
         Index           =   11
         Left            =   4200
         TabIndex        =   34
         Tag             =   "1313"
         Top             =   2730
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "如果：参数"
         Height          =   195
         Left            =   3420
         TabIndex        =   33
         Tag             =   "1602"
         Top             =   1290
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "COM接口ID"
         Height          =   225
         Index           =   0
         Left            =   4380
         TabIndex        =   32
         Tag             =   "1168"
         Top             =   1740
         Width           =   975
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "则，转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Tag             =   "1531"
         Top             =   2220
         Width           =   1065
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "否则，转节点ID"
         Height          =   195
         Left            =   4200
         TabIndex        =   29
         Tag             =   "1532"
         Top             =   2250
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "被访问日志"
         Height          =   195
         Index           =   0
         Left            =   5220
         TabIndex        =   28
         Tag             =   "1159"
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1320
         TabIndex        =   27
         Tag             =   "1137"
         Top             =   300
         Width           =   525
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   855
      Index           =   1
      Left            =   60
      TabIndex        =   21
      Tag             =   "1104"
      Top             =   3360
      Width           =   7395
      Begin VB.TextBox Txt_Description 
         Height          =   525
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frm_070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_070.frm
'//用途：    查询路由点
'//作者:     Tony Sun
'//创建日期：2005/06/28
'//修改日期：
'//文件描述：查询路由点
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_Object_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_param_Click()
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
    
    If Cb_usevar.ListIndex > 0 Then
        txtRoutePoint.Enabled = False
    Else
        txtRoutePoint.Enabled = True
    End If
    
End Sub

Private Sub cb_VarResult_Click()
    f_DataChanged = True
End Sub

Private Sub cmbSymbol_Click()
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

Private Sub cmdShowRes_Click()
    gSystem.intCurStep = 3
    Set gSystem.crlCurItem = T_com_iid
    frmResourceList.Show vbModal
End Sub

Private Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then

        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 节点超时
        If Trim(Cb_timeout) = "" Then
            Node70_Data1.timeout = 0
        Else
            If Val(Cb_timeout) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node70_Data1.timeout = CByte(Val(Cb_timeout.Text) Mod 256)
            End If
        End If
        
        Node70_Data1.reserved1 = 0
        
        ' 使用变量ID
        If Cb_usevar.ListIndex <= 0 Then
            Node70_Data1.usevar = 0
        Else
            Node70_Data1.usevar = CByte(Cb_usevar.ItemData(Cb_usevar.ListIndex))
        End If
        
        '' Sun added 2007-04-12
        ' 结果记录变量
        If cb_VarResult.ListIndex <= 0 Then
            Node70_Data1.var_result = 0
        Else
            Node70_Data1.var_result = CByte(cb_VarResult.ItemData(cb_VarResult.ListIndex))
        End If
        
        ' 查询对象
        If Cb_Object.ListIndex < 0 Then
            Node70_Data1.querytype = 0
        Else
            Node70_Data1.querytype = CByte(Cb_Object.ItemData(Cb_Object.ListIndex))
        End If
        
        ' 参数
        If Cb_param.ListIndex < 0 Then
            Node70_Data1.paramindex = 0
        Else
            Node70_Data1.paramindex = CByte(Cb_param.ItemData(Cb_param.ListIndex))
        End If
        
        ' 逻辑运算符
        If cmbSymbol.ListIndex < 0 Then
            Node70_Data1.logic = 0
        Else
            Node70_Data1.logic = CByte(cmbSymbol.ItemData(cmbSymbol.ListIndex))
        End If
                
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node70_Data1.log = 0
        Else
            Node70_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If
        
        Node70_Data1.reserved2(0) = 0
        
        ' Route Point
        Node70_Data2.routepointid = CInt(Val(txtRoutePoint))
        
        ' 变量比较值
        Node70_Data2.comparedvalue = CInt(Val(txtDefaultValue))
        
        Node70_Data2.reserved1(0) = 0
        
        ' COM接口ID
        If Trim(T_com_iid) = "" Then
            Node70_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E088")
                Exit Sub
            Else
                Node70_Data2.com_iid = CLng(Trim(T_com_iid.Text))
            End If
        End If
        
        Node70_Data2.reserved2(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node70_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node70_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
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
        If Node70_Data2.nd_yes <> lv_nNewNode Then
            Node70_Data2.nd_yes = lv_nNewNode
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
        If Node70_Data2.nd_no <> lv_nNewNode Then
            Node70_Data2.nd_no = lv_nNewNode
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
        If Node70_Data2.nd_fail <> lv_nNewNode Then
            Node70_Data2.nd_fail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 255)
        End If
        
        Node70_Data2.reserved3(0) = 0
        
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
    If Node70_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
        Message "M144"
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
    
    '' Sun added 2007-04-12
    ' ' 结果记录变量
    RefreshVariablesList cb_VarResult
    
    ' 查询对象
    With Cb_Object
        '' Route Point
        .AddItem LoadNationalResString(1603)
        .ItemData(.ListCount - 1) = 0
        '' Queue
        .AddItem LoadNationalResString(1628)
        .ItemData(.ListCount - 1) = 1
        '' Group
        .AddItem LoadNationalResString(1629)
        .ItemData(.ListCount - 1) = 2
        '' Team
        .AddItem LoadNationalResString(1630)
        .ItemData(.ListCount - 1) = 3
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
    With cmbSymbol
        For i = DEF_NODE016_LOGIC_EQUE To DEF_NODE016_LOGIC_NE
            .ItemData(i) = i
        Next
    End With

    '' Sun update 2008-05-16
    ''' Enlarges i from 8 to 19
    ' 参数
    With Cb_param
        For i = 0 To 8
            .AddItem LoadNationalResString(1604 + i)
            .ItemData(i) = i
        Next
        For i = 9 To 19
            .AddItem LoadNationalResString(2101 + i - 9)
            .ItemData(i) = i
        Next
    End With
        
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Cb_timeout.Text = Node70_Data1.timeout
    
    T_com_iid.Text = Node70_Data2.com_iid
    txtRoutePoint.Text = Node70_Data2.routepointid
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    txtDefaultValue.Text = Node70_Data2.comparedvalue
    T_nd_parent.Text = Node70_Data2.nd_parent
    T_nd_succeed.Text = Node70_Data2.nd_yes
    T_nd_fail.Text = Node70_Data2.nd_no
    txtNdNoAnswer.Text = Node70_Data2.nd_fail

    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node70_Data1.log), 0)
    cmbSymbol.ListIndex = SearchItemDataIndex(cmbSymbol, CLng(Node70_Data1.logic), 0)
    Cb_param.ListIndex = SearchItemDataIndex(Cb_param, CLng(Node70_Data1.paramindex), 0)

    '' Sun added 2005-08-05
    Cb_Object.ListIndex = SearchItemDataIndex(Cb_Object, CLng(Node70_Data1.querytype), 0)
    Cb_usevar.ListIndex = SearchItemDataIndex(Cb_usevar, CLng(Node70_Data1.usevar), 0)
    
    '' Sun added 2007-04-12
    cb_VarResult.ListIndex = SearchItemDataIndex(cb_VarResult, CLng(Node70_Data1.var_result), 0)
    
    If Node70_Data1.usevar = 0 Then
        txtRoutePoint.Enabled = True
    Else
        txtRoutePoint.Enabled = False
    End If
    
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

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtDefaultValue_Change()
    f_DataChanged = True
End Sub

Private Sub txtDefaultValue_GotFocus()
    txtDefaultValue.SelStart = 0
    txtDefaultValue.SelLength = Len(txtDefaultValue)
End Sub

Private Sub txtDefaultValue_KeyPress(KeyAscii As Integer)
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

Private Sub txtRoutePoint_Change()
    f_DataChanged = True
End Sub

Private Sub txtRoutePoint_GotFocus()
    txtRoutePoint.SelStart = 0
    txtRoutePoint.SelLength = Len(txtRoutePoint)
End Sub

Private Sub txtRoutePoint_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
