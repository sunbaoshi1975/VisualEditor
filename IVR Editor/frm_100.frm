VERSION 5.00
Begin VB.Form frm_100 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户DLL"
   ClientHeight    =   3855
   ClientLeft      =   4920
   ClientTop       =   2925
   ClientWidth     =   3225
   Icon            =   "frm_100.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1064"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   2832
      Picture         =   "frm_100.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "1948"
      Top             =   3441
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   885
      Left            =   30
      TabIndex        =   15
      Tag             =   "1104"
      Top             =   2430
      Width           =   3135
      Begin VB.TextBox Txt_Description 
         Height          =   525
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   270
         Width           =   2955
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2325
      Left            =   60
      TabIndex        =   10
      Tag             =   "1136"
      Top             =   60
      Width           =   3105
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2640
         Picture         =   "frm_100.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "1145"
         Top             =   1860
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2640
         Picture         =   "frm_100.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "1145"
         Top             =   1500
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2640
         Picture         =   "frm_100.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "1146"
         Top             =   1140
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   750
         Width           =   1395
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1470
         Width           =   1065
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
         TabIndex        =   5
         Top             =   1830
         Width           =   1065
      End
      Begin VB.TextBox T_dll_fid 
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
         TabIndex        =   1
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1380
         TabIndex        =   18
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label20 
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Tag             =   "1169"
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label Label19 
         Caption         =   "子节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Tag             =   "1252"
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "被访问日志"
         Height          =   225
         Left            =   180
         TabIndex        =   12
         Tag             =   "1159"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "DLL文件ID"
         Height          =   225
         Left            =   180
         TabIndex        =   11
         Tag             =   "1347"
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   1380
      TabIndex        =   9
      Tag             =   "1144"
      Top             =   3420
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Tag             =   "1007"
      Top             =   3420
      Width           =   1035
   End
End
Attribute VB_Name = "frm_100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_100.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：用户DLL
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_log_Click()
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

    gSystem.intCurStep = 3
    Set gSystem.crlCurItem = T_dll_fid
    frmResourceList.Show vbModal

End Sub

Public Sub Command1_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 保留
        Node100_Data1.reserved1(0) = 0
    
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
            Node100_Data1.log = 0
        Else
            Node100_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        ' 保留
        Node100_Data1.reserved2(0) = 0
        
        ' DLL文件ID
        If Trim(T_dll_fid) = "" Then
            Node100_Data2.dll_fid = 0
        Else
            If CLng(Trim(T_dll_fid)) > 32767 Then
                Message ("E087")
                T_dll_fid.SetFocus
                Exit Sub
            Else
                Node100_Data2.dll_fid = CLng(Trim(T_dll_fid.Text))
            End If
        End If
        
        ' 保留
        Node100_Data2.reserved1(0) = 0
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
           Node100_Data2.nd_parent = 0
        Else
           If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
           Else
              Node100_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
           End If
        End If
        
        '子节点
        If Trim(T_nd_child) = "" Then
           Node100_Data2.nd_child = 0
        Else
           If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
              Message ("E072")
              T_nd_child.SetFocus
              Exit Sub
           Else
              Node100_Data2.nd_child = CInt(Trim(T_nd_child.Text))
            End If
        End If
         
        ' 保留
        Node100_Data2.reserved2(0) = 0
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no)

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
 
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    T_dll_fid.Text = Node100_Data2.dll_fid
    T_nd_parent.Text = Node100_Data2.nd_parent
    T_nd_child.Text = Node100_Data2.nd_child
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node100_Data1.log), 0)
    
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

Private Sub T_dll_fid_Change()
    f_DataChanged = True

    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_dll_fid

End Sub

Private Sub T_dll_fid_GotFocus()
    T_dll_fid.SelStart = 0
    T_dll_fid.SelLength = Len(T_dll_fid)
End Sub

Private Sub T_dll_fid_KeyPress(KeyAscii As Integer)
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
