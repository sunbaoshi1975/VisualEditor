VERSION 5.00
Begin VB.Form frm_016 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "条件分支"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frm_016.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1528"
   Visible         =   0   'False
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6822
      Picture         =   "frm_016.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "1948"
      Top             =   3951
      Width           =   333
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Tag             =   "1007"
      Top             =   3930
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   885
      Left            =   60
      TabIndex        =   25
      Tag             =   "1104"
      Top             =   2910
      Width           =   7095
      Begin VB.TextBox Txt_Description 
         Height          =   525
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   270
         Width           =   6915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   60
      TabIndex        =   16
      Tag             =   "1136"
      Top             =   60
      Width           =   7095
      Begin VB.ComboBox Cb_var 
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
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
         Width           =   1755
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
         ItemData        =   "frm_016.frx":20B4
         Left            =   2040
         List            =   "frm_016.frx":20CD
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtParam 
         Alignment       =   2  'Center
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
         Left            =   5220
         MaxLength       =   3
         TabIndex        =   3
         Top             =   780
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtParam 
         Alignment       =   2  'Center
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
         Left            =   6180
         MaxLength       =   3
         TabIndex        =   4
         Top             =   780
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.ComboBox cmbProcess 
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
         ItemData        =   "frm_016.frx":20EF
         Left            =   3840
         List            =   "frm_016.frx":2102
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   975
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1740
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
         Left            =   5460
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1740
         Width           =   1065
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3120
         Picture         =   "frm_016.frx":211F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "显示节点列表"
         Top             =   1770
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   6540
         Picture         =   "frm_016.frx":24A9
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "显示节点列表"
         Top             =   1770
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3120
         Picture         =   "frm_016.frx":2833
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "显示节点列表"
         Top             =   2250
         Width           =   315
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
         ItemData        =   "frm_016.frx":2BBD
         Left            =   5460
         List            =   "frm_016.frx":2BBF
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   240
         Width           =   1455
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
         TabIndex        =   18
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   915
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2220
         Width           =   1065
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
         Left            =   3840
         MaxLength       =   48
         TabIndex        =   6
         Top             =   1260
         Width           =   3045
      End
      Begin VB.Label lblComma 
         AutoSize        =   -1  'True
         Caption         =   ","
         Height          =   195
         Left            =   5940
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   ")"
         Height          =   195
         Left            =   6840
         TabIndex        =   28
         Top             =   840
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "("
         Height          =   195
         Left            =   5040
         TabIndex        =   27
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "否则，转节点ID"
         Height          =   195
         Left            =   3840
         TabIndex        =   26
         Tag             =   "1532"
         Top             =   1830
         Width           =   1365
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1380
         TabIndex        =   24
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "如果：变量"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Tag             =   "1530"
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   4080
         TabIndex        =   21
         Tag             =   "1159"
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "则，转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Tag             =   "1531"
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Tag             =   "1169"
         Top             =   2280
         Width           =   705
      End
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Tag             =   "1144"
      Top             =   3930
      Width           =   1035
   End
End
Attribute VB_Name = "frm_016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_016.frm
'//用途：    条件分支
'//作者:     Sun
'//创建日期：2004/12/30
'//修改日期：
'//文件描述：选择服务语言
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_var_Click()
    f_DataChanged = True
End Sub

Private Sub cmbProcess_Click()
    f_DataChanged = True
    
    Select Case cmbProcess.ListIndex
    Case 1, 2
        txtParam(0).Visible = True
        lblComma.Visible = False
        txtParam(1).Visible = False
    Case 3
        txtParam(0).Visible = True
        lblComma.Visible = True
        txtParam(1).Visible = True
    Case Else
        txtParam(0).Visible = False
        lblComma.Visible = False
        txtParam(1).Visible = False
    End Select
End Sub

Private Sub cmbSymbol_Click()
    f_DataChanged = True
End Sub

'Mike added this event @ 2008-1-30
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
    End Select
    frmNodeList.Show vbModal

End Sub


Private Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer

        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        Node16_Data1.reserved1(0) = 0
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node16_Data1.log = 0
        Else
            Node16_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        ' 转换公式
        If cmbProcess.ListIndex < 0 Then
            Node16_Data1.convert = 0
            Node16_Data1.param1 = 0
            Node16_Data1.param2 = 0
        Else
            Node16_Data1.convert = CByte(cmbProcess.ItemData(cmbProcess.ListIndex))
            Select Case Node16_Data1.convert
            Case 1, 2
                Node16_Data1.param1 = CByte(txtParam(0))
                Node16_Data1.param2 = 0
            Case 3
                Node16_Data1.param1 = CByte(txtParam(0))
                Node16_Data1.param2 = CByte(txtParam(1))
            Case 4
                If Trim(txtDefaultValue) <> Trim(Str(Val(txtDefaultValue))) Then
                    Message ("E031")
                    txtDefaultValue.SetFocus
                    Exit Sub
                End If
            Case Else
                Node16_Data1.param1 = 0
                Node16_Data1.param2 = 0
            End Select
        End If

        ' 逻辑运算符
        If cmbSymbol.ListIndex < 0 Then
            Node16_Data1.logic = 0
        Else
            Node16_Data1.logic = CByte(cmbSymbol.ItemData(cmbSymbol.ListIndex))
        End If

        Node16_Data1.reserved2(0) = 0
        
        ' 变量
        If Cb_var.ListIndex <= 0 Then
            Node16_Data1.var_id = 0
        Else
            Node16_Data1.var_id = CByte(Cb_var.ItemData(Cb_var.ListIndex))
        End If
        
        ' 变量比较值
        Call StringToByteArray(txtDefaultValue, Node16_Data2.var_value, 48)
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node16_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node16_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
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
        If Node16_Data2.nd_succ <> lv_nNewNode Then
            Node16_Data2.nd_succ = lv_nNewNode
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
        If Node16_Data2.nd_fail <> lv_nNewNode Then
            Node16_Data2.nd_fail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
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

    ' 逻辑运算符
    With cmbSymbol
        For i = DEF_NODE016_LOGIC_EQUE To DEF_NODE016_LOGIC_LIKE
            .ItemData(i) = i
        Next
    End With

    ' 转换公式
    With cmbProcess
        For i = DEF_NODE016_CONVERT_NONE To DEF_NODE016_CONVERT_LEN
            .ItemData(i) = i
        Next
    End With

    ' 变量
    RefreshVariablesList Cb_var
    
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    txtParam(0).Text = Node16_Data1.param1
    txtParam(1).Text = Node16_Data1.param2
    txtDefaultValue.Text = ByteArrayToString(Node16_Data2.var_value, 24)
    T_nd_parent.Text = Node16_Data2.nd_parent
    T_nd_succeed.Text = Node16_Data2.nd_succ
    T_nd_fail.Text = Node16_Data2.nd_fail

    Cb_var.ListIndex = SearchItemDataIndex(Cb_var, CLng(Node16_Data1.var_id), 0)
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node16_Data1.log), 0)
    cmbSymbol.ListIndex = SearchItemDataIndex(cmbSymbol, CLng(Node16_Data1.logic), 0)
    cmbProcess.ListIndex = SearchItemDataIndex(cmbProcess, CLng(Node16_Data1.convert), 0)
    
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
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

Private Sub txtParam_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub txtParam_GotFocus(Index As Integer)
    txtParam(Index).SelStart = 0
    txtParam(Index).SelLength = Len(txtParam(Index))
End Sub

Private Sub txtParam_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
