VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_018 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据发送"
   ClientHeight    =   7035
   ClientLeft      =   4665
   ClientTop       =   2790
   ClientWidth     =   8400
   Icon            =   "frm_018.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1223"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   7992
      Picture         =   "frm_018.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "1948"
      Top             =   6591
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1035
      Left            =   60
      TabIndex        =   28
      Tag             =   "1104"
      Top             =   5430
      Width           =   8265
      Begin VB.TextBox Txt_Description 
         Height          =   615
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   330
         Width           =   8025
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   3030
      TabIndex        =   16
      Tag             =   "1007"
      Top             =   6600
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   4350
      TabIndex        =   17
      Tag             =   "1144"
      Top             =   6600
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   60
      TabIndex        =   18
      Tag             =   "1136"
      Top             =   60
      Width           =   8265
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   345
         Index           =   5
         Left            =   3270
         TabIndex        =   8
         Top             =   3990
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "<"
         Height          =   345
         Index           =   4
         Left            =   3270
         TabIndex        =   7
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">"
         Height          =   345
         Index           =   3
         Left            =   3270
         TabIndex        =   6
         Top             =   3210
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   345
         Index           =   2
         Left            =   3270
         TabIndex        =   4
         Top             =   1950
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "<"
         Height          =   345
         Index           =   1
         Left            =   3270
         TabIndex        =   3
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">"
         Height          =   345
         Index           =   0
         Left            =   3270
         TabIndex        =   2
         Top             =   1170
         Width           =   495
      End
      Begin VB.TextBox txtPrefix 
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
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2700
         Width           =   585
      End
      Begin VB.TextBox txtSeperator 
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
         Left            =   5010
         MaxLength       =   1
         TabIndex        =   0
         Top             =   240
         Width           =   585
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   6690
         Picture         =   "frm_018.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "显示节点列表"
         Top             =   4800
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   3180
         Picture         =   "frm_018.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "显示节点列表"
         Top             =   4800
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
         TabIndex        =   25
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
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
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
         Left            =   5190
         MaxLength       =   6
         TabIndex        =   13
         Top             =   4770
         Width           =   1485
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   11
         Top             =   4770
         Width           =   1485
      End
      Begin MSComctlLib.ListView lstSysVarList 
         Height          =   3495
         Left            =   150
         TabIndex        =   1
         Top             =   1080
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Description"
            Object.Tag             =   "1545"
            Text            =   "ID"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "1546"
            Text            =   "变量名称"
            Object.Width           =   2823
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "1547"
            Text            =   "长度"
            Object.Width           =   1058
         EndProperty
      End
      Begin MSComctlLib.ListView lstAppData 
         Height          =   1575
         Left            =   3930
         TabIndex        =   9
         Top             =   1080
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Description"
            Object.Tag             =   "1545"
            Text            =   "ID"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "1546"
            Text            =   "变量名称"
            Object.Width           =   2823
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "1547"
            Text            =   "长度"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "1543"
            Text            =   "前缀"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "1539"
            Text            =   "分割符"
            Object.Width           =   1058
         EndProperty
      End
      Begin MSComctlLib.ListView lstUserData 
         Height          =   1575
         Left            =   3930
         TabIndex        =   10
         Top             =   3000
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Description"
            Object.Tag             =   "1545"
            Text            =   "ID"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "1546"
            Text            =   "变量名称"
            Object.Width           =   2823
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "1547"
            Text            =   "长度"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "1543"
            Text            =   "前缀"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "1539"
            Text            =   "分割符"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label lblUserDataCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 / 256 Bytes"
         Height          =   285
         Left            =   6810
         TabIndex        =   32
         Top             =   2700
         Width           =   1290
      End
      Begin VB.Label lblAppDataCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 / 128 Bytes"
         Height          =   285
         Left            =   6810
         TabIndex        =   31
         Top             =   780
         Width           =   1290
      End
      Begin VB.Label lblPrefix 
         AutoSize        =   -1  'True
         Caption         =   "Prefix"
         Height          =   195
         Left            =   3240
         TabIndex        =   30
         Tag             =   "1543"
         Top             =   2430
         Width           =   390
      End
      Begin VB.Label lblUserData 
         AutoSize        =   -1  'True
         Caption         =   "UserData发送串"
         Height          =   195
         Left            =   3990
         TabIndex        =   29
         Tag             =   "1542"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1380
         TabIndex        =   26
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "子节点ID"
         Height          =   195
         Left            =   3990
         TabIndex        =   23
         Tag             =   "1229"
         Top             =   4830
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Tag             =   "1169"
         Top             =   4830
         Width           =   705
      End
      Begin VB.Label lblAppData 
         AutoSize        =   -1  'True
         Caption         =   "AppData发送串"
         Height          =   195
         Left            =   3990
         TabIndex        =   21
         Tag             =   "1541"
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label lblSysVarList 
         AutoSize        =   -1  'True
         Caption         =   "变量列表"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Tag             =   "1540"
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Lbl_tcpip 
         AutoSize        =   -1  'True
         Caption         =   "分隔符"
         Height          =   195
         Left            =   3990
         TabIndex        =   19
         Tag             =   "1539"
         Top             =   300
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_018.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：发送数据
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub cmdMove_Click(Index As Integer)
On Error GoTo BackDoor

    Dim itmSource As ListItem                    ' ListItem 变量
    Dim itmTarget As ListItem                    ' ListItem 变量
    Dim lstTarget As ListView
    Dim lblTarget As Label
    Dim lv_Len As Integer
    Dim lv_nIndex As Integer
    
    ' Target List
    If Index < 3 Then
        Set lstTarget = lstAppData
        Set lblTarget = lblAppDataCount
    Else
        Set lstTarget = lstUserData
        Set lblTarget = lblUserDataCount
    End If
    
    Select Case Index
    Case 0, 3
        ' Get Current Item in Source List
        Set itmSource = lstSysVarList.SelectedItem
        If itmSource Is Nothing Then Exit Sub
        
        ' Find in Target List
        Set itmTarget = lstTarget.FindItem(itmSource)
        If Not (itmTarget Is Nothing) Then
            Message "M140"
            Exit Sub
        End If
        
        ' 增加条目
        If lstTarget.SelectedItem Is Nothing Then
            lv_nIndex = lstTarget.ListItems.Count + 1
        Else
            lv_nIndex = lstTarget.SelectedItem.Index + 1
        End If
        Set itmTarget = lstTarget.ListItems.Add(lv_nIndex, , itmSource)
        itmTarget.SubItems(1) = itmSource.SubItems(1)
        itmTarget.SubItems(2) = itmSource.SubItems(2)
        itmTarget.SubItems(3) = Trim(txtPrefix)
        itmTarget.SubItems(4) = Trim(txtSeperator)
        txtPrefix = ""
        
        ' 计数
        lv_Len = Val(lstTarget.Tag)
        lv_Len = lv_Len + Val(itmSource.SubItems(2)) + Len(Trim(txtPrefix)) + Len(Trim(txtSeperator))
        lstTarget.Tag = Trim(Str(lv_Len))
        
        Set lstTarget.SelectedItem = itmTarget
                
    '' 删除一个条目
    Case 1, 4
        ' Find in Target List
        Set itmTarget = lstTarget.SelectedItem
        If itmTarget Is Nothing Then Exit Sub
                        
        ' 计数
        lv_Len = Val(lstTarget.Tag)
        lv_Len = lv_Len - Val(itmTarget.SubItems(2)) - Len(Trim(itmTarget.SubItems(3))) - Len(Trim(itmTarget.SubItems(4)))
        lstTarget.Tag = Trim(Str(lv_Len))
        
        lstTarget.ListItems.Remove itmTarget.Index

                        
    '' 清空目标列表
    Case 2, 5
        lstTarget.ListItems.Clear
        
        ' 计数
        lstTarget.Tag = "0"
    
    End Select
    
    If Index < 3 Then
        lblTarget = lstTarget.Tag & " / 128 Bytes"
    Else
        lblTarget = lstTarget.Tag & " / 256 Bytes"
    End If
    
    f_DataChanged = True
    
BackDoor:
    Debug.Print Err.Description
    On Error GoTo 0
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
        Set gSystem.crlCurItem = T_nd_child
    End Select
    frmNodeList.Show vbModal

End Sub

Public Sub Command1_Click()
On Error Resume Next

Dim lv_loop As Integer
Dim lv_StrTemp As String

    If f_DataChanged Then

        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        Node18_Data1.reserved1(0) = 0
        
        ' 分隔符
        If Trim(txtSeperator) = "" Then
            Node18_Data1.seperator = 0
        Else
            Node18_Data1.seperator = Asc(Left(txtSeperator, 1))
        End If
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node18_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node18_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If

        ' 子节点
        If Trim(T_nd_child) = "" Then
           Node18_Data2.nd_child = 0
        Else
            If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
                Message ("E072")
                T_nd_child.SetFocus
                Exit Sub
            Else
                Node18_Data2.nd_child = CInt(Trim(T_nd_child.Text))
            End If
        End If
         
        ' AppData & UserData List
        For lv_loop = 0 To 14
            Node18_Data2.typeflags(lv_loop) = 0
            Node18_Data2.prefix1(lv_loop) = 0
            Node18_Data2.prefix2(lv_loop) = 0
            Node18_Data2.valueid(lv_loop) = 0
        Next
        For lv_loop = 1 To lstAppData.ListItems.Count
            Node18_Data2.typeflags(lv_loop - 1) = 1
            lv_StrTemp = lstAppData.ListItems(lv_loop).SubItems(3)
            If Len(lv_StrTemp) > 0 Then
                Node18_Data2.prefix1(lv_loop - 1) = Asc(Left(lv_StrTemp, 1))
            End If
            If Len(lv_StrTemp) > 1 Then
                Node18_Data2.prefix2(lv_loop - 1) = Asc(Right(lv_StrTemp, 1))
            End If
            Node18_Data2.valueid(lv_loop - 1) = CByte(Val(lstAppData.ListItems(lv_loop)))
        Next
        For lv_loop = lstAppData.ListItems.Count + 1 To lstAppData.ListItems.Count + lstUserData.ListItems.Count
            If lv_loop > 14 Then Exit For
            
            Node18_Data2.typeflags(lv_loop - 1) = 2
            lv_StrTemp = lstUserData.ListItems(lv_loop - lstAppData.ListItems.Count).SubItems(3)
            If Len(lv_StrTemp) > 0 Then
                Node18_Data2.prefix1(lv_loop - 1) = Asc(Left(lv_StrTemp, 1))
            End If
            If Len(lv_StrTemp) > 1 Then
                Node18_Data2.prefix2(lv_loop - 1) = Asc(Right(lv_StrTemp, 1))
            End If
            Node18_Data2.valueid(lv_loop - 1) = CByte(Val(lstUserData.ListItems(lv_loop - lstAppData.ListItems.Count)))
        Next
                
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

Private Sub Command2_Click()
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
    
    ' 分隔符
    If Node18_Data1.seperator > Asc(" ") Then
        txtSeperator = Chr$(Node18_Data1.seperator)
    Else
        txtSeperator = ""
    End If
    
    txtPrefix = ""
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_nd_parent.Text = Node18_Data2.nd_parent
    T_nd_child.Text = Node18_Data2.nd_child
        
    ' 变量列表
    Call RefreshVarList
    
    ' AppData发送串 & UserData发送串
    Call FillDataList
    
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

Private Sub txtPrefix_GotFocus()
    txtPrefix.SelStart = 0
    txtPrefix.SelLength = Len(txtPrefix)
End Sub

Private Sub txtSeperator_Change()
    f_DataChanged = True
End Sub

Private Sub txtSeperator_GotFocus()
    txtSeperator.SelStart = 0
    txtSeperator.SelLength = Len(txtSeperator)
End Sub

Private Sub RefreshVarList()
On Error Resume Next

Dim lv_loop As Integer
Dim lv_strName As String
Dim lv_bytType As Byte
Dim lv_bytLen As Byte
Dim lv_nVarCount As Integer

Dim itmX As ListItem                    ' ListItem 变量

    lstSysVarList.ListItems.Clear
    lv_nVarCount = gCallFlow.GetUserVarCount
    For lv_loop = 1 To lv_nVarCount
        
        lv_strName = ""
        lv_bytType = 0
        lv_bytLen = 0
        Call gCallFlow.GetUserVarDefination(lv_loop, lv_strName, lv_bytType, lv_bytLen)
    
        ''' 使用 Add 方法添加新的 ListItem 并为新引用设置对象。
        ''' 使用引用设置属性。
        Set itmX = lstSysVarList.ListItems.Add(lv_loop, , Str(lv_loop))
        itmX.SubItems(1) = lv_strName
        itmX.SubItems(2) = Str(lv_bytLen)
                    
    Next
    
On Error GoTo 0
End Sub

Private Sub FillDataList()
On Error Resume Next

Dim lv_loop As Integer
Dim lv_bytLen As Byte
Dim lv_AppLen As Byte
Dim lv_UserLen As Byte
Dim lv_sTemp As String

Dim itmX As ListItem                    ' ListItem 变量
Dim itmSource As ListItem               ' ListItem 变量

    lstAppData.ListItems.Clear
    lstUserData.ListItems.Clear
    lv_AppLen = 0
    lv_UserLen = 0
    
    For lv_loop = 0 To 14
        
        If Node18_Data2.typeflags(lv_loop) = 0 Then
            Exit For
        Else
            '' Prefix
            lv_sTemp = ""
            If Node18_Data2.prefix1(lv_loop) > 0 Then
                lv_sTemp = Chr(Node18_Data2.prefix1(lv_loop))
                If Node18_Data2.prefix2(lv_loop) > 0 Then
                    lv_sTemp = lv_sTemp & Chr(Node18_Data2.prefix2(lv_loop))
                End If
            End If
            
            '' Source Item
            Set itmSource = lstSysVarList.FindItem(Str(Node18_Data2.valueid(lv_loop)))
            If Not (itmSource Is Nothing) Then
                
                If Node18_Data2.typeflags(lv_loop) = 1 Then
                    Set itmX = lstAppData.ListItems.Add(, , itmSource)
                    lv_bytLen = Val(itmSource.SubItems(2))
                    lv_AppLen = lv_AppLen + lv_bytLen
                Else
                    Set itmX = lstUserData.ListItems.Add(, , itmSource)
                    lv_bytLen = Val(itmSource.SubItems(2))
                    lv_UserLen = lv_UserLen + lv_bytLen
                End If
                itmX.SubItems(1) = itmSource.SubItems(1)
                itmX.SubItems(2) = itmSource.SubItems(2)
                itmX.SubItems(3) = lv_sTemp
                itmX.SubItems(4) = txtSeperator
                
            End If
        End If
                    
    Next
    
    lstAppData.Tag = Trim(Str(lv_AppLen))
    lblAppDataCount = lstAppData.Tag & " / 128 Bytes"
    
    lstUserData.Tag = Trim(Str(lv_UserLen))
    lblUserDataCount = lstUserData.Tag & " / 256 Bytes"
    
On Error GoTo 0
End Sub

