VERSION 5.00
Begin VB.Form frm_102 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "记录变量"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   Icon            =   "frm_102.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1348"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   3372
      Picture         =   "frm_102.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "1948"
      Top             =   4791
      Width           =   333
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   17
      Tag             =   "1007"
      Top             =   4770
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   1380
      TabIndex        =   18
      Tag             =   "1144"
      Top             =   4770
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3705
      Left            =   30
      TabIndex        =   14
      Tag             =   "1136"
      Top             =   30
      Width           =   3675
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
         ItemData        =   "frm_102.frx":16B4
         Left            =   420
         List            =   "frm_102.frx":16C7
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
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
         Index           =   1
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2280
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
         Index           =   0
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   525
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
         Left            =   1560
         MaxLength       =   24
         TabIndex        =   6
         Top             =   1860
         Width           =   1965
      End
      Begin VB.ComboBox Cb_var_key 
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
         TabIndex        =   3
         Top             =   1110
         Width           =   1965
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   10
         Top             =   2820
         Width           =   1605
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   12
         Top             =   3180
         Width           =   1605
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1500
         Width           =   1605
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   915
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
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   0
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   720
         Width           =   1965
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   3180
         Picture         =   "frm_102.frx":16E4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "显示资源列表"
         Top             =   1530
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3180
         Picture         =   "frm_102.frx":17E6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "显示节点列表"
         Top             =   2850
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   3180
         Picture         =   "frm_102.frx":1B70
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "显示节点列表"
         Top             =   3210
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "("
         Height          =   195
         Left            =   1620
         TabIndex        =   29
         Top             =   2340
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   ")"
         Height          =   195
         Left            =   3420
         TabIndex        =   28
         Top             =   2340
         Width           =   105
      End
      Begin VB.Label lblComma 
         AutoSize        =   -1  'True
         Caption         =   ","
         Height          =   195
         Left            =   2520
         TabIndex        =   27
         Top             =   2340
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "变量缺省值"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Tag             =   "1350"
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "记录变量ID"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Tag             =   "1349"
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Tag             =   "1169"
         Top             =   2880
         Width           =   705
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "子节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Tag             =   "1252"
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Tag             =   "1159"
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "COM接口ID"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Tag             =   "1168"
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1680
         TabIndex        =   19
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   885
      Left            =   30
      TabIndex        =   16
      Tag             =   "1104"
      Top             =   3810
      Width           =   3675
      Begin VB.TextBox Txt_Description 
         Height          =   525
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   270
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frm_102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_102.frm
'//用途：    记录变量节点
'//作者:     Sun
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：记录变量节点
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_var_key_Click()
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

    gSystem.intCurStep = 3
    Set gSystem.crlCurItem = T_com_iid
    frmResourceList.Show vbModal

End Sub

Public Sub Command1_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 保留
        Node102_Data1.reserved1(0) = 0
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node102_Data1.log = 0
        Else
            Node102_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        ' 转换公式
        If cmbProcess.ListIndex < 0 Then
            Node102_Data1.convert = 0
            Node102_Data1.param1 = 0
            Node102_Data1.param2 = 0
        Else
            Node102_Data1.convert = CByte(cmbProcess.ItemData(cmbProcess.ListIndex))
            Select Case Node102_Data1.convert
            Case 1, 2
                Node102_Data1.param1 = CByte(txtParam(0))
                Node102_Data1.param2 = 0
            Case 3
                Node102_Data1.param1 = CByte(txtParam(0))
                Node102_Data1.param2 = CByte(txtParam(1))
            Case 4
                If Trim(txtDefaultValue) <> Trim(Str(Val(txtDefaultValue))) Then
                    Message ("E031")
                    txtDefaultValue.SetFocus
                    Exit Sub
                End If
            Case Else
                Node102_Data1.param1 = 0
                Node102_Data1.param2 = 0
            End Select
        End If

        ' 按键记录
        If Cb_var_key.ListIndex <= 0 Then
            Node102_Data1.var_chg = 0
        Else
            Node102_Data1.var_chg = CByte(Cb_var_key.ItemData(Cb_var_key.ListIndex))
        End If
        
        ' 保留
        Node102_Data1.reserved2(0) = 0
        
        ' 保留
        Node102_Data2.reserved1(0) = 0
        
        ' COM接口ID
        If Trim(T_com_iid) = "" Then
            Node102_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E088")
                Exit Sub
            Else
                Node102_Data2.com_iid = CLng(Trim(T_com_iid.Text))
            End If
        End If
        
        ' 缺省值
        Call StringToByteArray(txtDefaultValue, Node102_Data2.value, txtDefaultValue.MaxLength)
        
        ' 保留
        Node102_Data2.reserved2(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
           Node102_Data2.nd_parent = 0
        Else
           If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
           Else
              Node102_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
           End If
        End If
        
        '子节点
        If Trim(T_nd_child) = "" Then
           Node102_Data2.nd_child = 0
        Else
           If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
              Message ("E072")
              T_nd_child.SetFocus
              Exit Sub
           Else
              Node102_Data2.nd_child = CInt(Trim(T_nd_child.Text))
           End If
        End If
        
        ' 保留
        Node102_Data2.reserved3(0) = 0
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
   
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no)
        
        f_DataChanged = False
        
    End If

    '' Sun added 2007-03-25
    If Node102_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
        Message "M144"
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

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
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

    ' 记录变量ID
    RefreshVariablesList Cb_var_key
    
    ' 转换公式
    With cmbProcess
        For i = DEF_NODE016_CONVERT_NONE To DEF_NODE016_CONVERT_LEN
            .ItemData(i) = i
        Next
    End With
    
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    txtParam(0).Text = Node102_Data1.param1
    txtParam(1).Text = Node102_Data1.param2
    T_com_iid.Text = Node102_Data2.com_iid
    T_nd_parent.Text = Node102_Data2.nd_parent
    T_nd_child.Text = Node102_Data2.nd_child
    txtDefaultValue = ByteArrayToString(Node102_Data2.value, txtDefaultValue.MaxLength)
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node102_Data1.log), 0)
    Cb_var_key.ListIndex = SearchItemDataIndex(Cb_var_key, CLng(Node102_Data1.var_chg), 0)
    cmbProcess.ListIndex = SearchItemDataIndex(cmbProcess, CLng(Node102_Data1.convert), 0)
        
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

