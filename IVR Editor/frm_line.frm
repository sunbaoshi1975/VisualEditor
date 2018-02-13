VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm_Line 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "节点连线"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3075
   Icon            =   "frm_line.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1378"
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1095
      Left            =   60
      TabIndex        =   22
      Tag             =   "1104"
      Top             =   3210
      Width           =   2955
      Begin VB.TextBox Txt_Description 
         Height          =   705
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   300
         Width           =   2715
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   330
      TabIndex        =   10
      Tag             =   "1007"
      Top             =   4410
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   375
      Left            =   1620
      TabIndex        =   11
      Tag             =   "1144"
      Top             =   4410
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00FF0000&
      Height          =   3105
      Left            =   60
      TabIndex        =   0
      Tag             =   "1136"
      Top             =   60
      Width           =   2955
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   885
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
      Begin VB.PictureBox LineColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   435
         TabIndex        =   7
         Top             =   2640
         Width           =   465
      End
      Begin VB.CommandButton CustNodeHand 
         Caption         =   "…"
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         ToolTipText     =   "1520"
         Top             =   2640
         Width           =   315
      End
      Begin VB.TextBox txtWidth 
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
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2220
         Width           =   1035
      End
      Begin ComCtl2.UpDown updWidth 
         Height          =   360
         Left            =   2596
         TabIndex        =   6
         Top             =   2220
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196618
         OrigLeft        =   1950
         OrigTop         =   2640
         OrigRight       =   2190
         OrigBottom      =   2865
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cmbStyle 
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
         TabIndex        =   4
         Top             =   1845
         Width           =   1275
      End
      Begin VB.TextBox txtIndex 
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
         MaxLength       =   2
         TabIndex        =   1
         Top             =   735
         Width           =   1275
      End
      Begin VB.TextBox nd_start_text 
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
         TabIndex        =   2
         Top             =   1110
         Width           =   1275
      End
      Begin VB.TextBox nd_end_text 
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
         HelpContextID   =   6
         Left            =   1560
         TabIndex        =   3
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1290
         TabIndex        =   20
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "颜色"
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Tag             =   "1385"
         Top             =   2670
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "序号"
         Height          =   225
         Left            =   180
         TabIndex        =   16
         Tag             =   "1379"
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "起始节点"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Tag             =   "1381"
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label19 
         Caption         =   "终止节点"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Tag             =   "1382"
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "线型"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Tag             =   "1383"
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "线宽"
         Height          =   180
         Left            =   180
         TabIndex        =   12
         Tag             =   "1384"
         Top             =   2280
         Width           =   360
      End
   End
End
Attribute VB_Name = "frm_Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub cmbStyle_Click()
    f_DataChanged = True
End Sub

Private Sub Command1_Click()
On Error Resume Next

    Dim lv_bytIndex As Byte
    Dim lv_blnResort As Boolean
    Dim lv_Str As String
    
    If f_DataChanged Then

        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        lv_bytIndex = Val(txtIndex) Mod 255
        If Node255_Data2.Index <> lv_bytIndex Then
            Node255_Data2.Index = lv_bytIndex
            lv_blnResort = True
        Else
            lv_blnResort = False
        End If
        
        Node255_Data2.StartNode = Val(nd_start_text)
        Node255_Data2.EndNode = Val(nd_end_text)
        If cmbStyle.ListIndex < 0 Then
            Node255_Data2.Style = 1
        Else
            Node255_Data2.Style = cmbStyle.ItemData(cmbStyle.ListIndex)
        End If
        Node255_Data2.Width = Val(txtWidth.Text)
        Node255_Data2.Color = CLng(LineColor.BackColor)
        
        gCallFlow.Node(gCallFlow.NodeSelectedID).Line_Color = Node255_Data2.Color
        gCallFlow.Node(gCallFlow.NodeSelectedID).Line_Style = Node255_Data2.Style
        gCallFlow.Node(gCallFlow.NodeSelectedID).Line_Width = Node255_Data2.Width
            
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = ""
        Else
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        '' Sun added 2007-03-25
        ''' 根据连线序号更新前节点属性
        If Node255_Data2.StartNode > 255 And Node255_Data2.EndNode > 255 And lv_blnResort Then
            lv_Str = gCallFlow.ResetNodeChildrenNodes(gCallFlow.NodeSelectedID, Node255_Data2.StartNode, Node255_Data2.EndNode, Node255_Data2.Index)
            If lv_Str <> "" Then
                gCallFlow.Node(gCallFlow.NodeSelectedID).Description = lv_Str
            End If
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

Private Sub CustNodeHand_Click()
Dim lv_OldColor As OLE_COLOR

    lv_OldColor = LineColor.BackColor
    F_CustomColor lv_OldColor, LineColor
    
    If lv_OldColor <> LineColor.BackColor Then
        f_DataChanged = True
    End If
    
End Sub

Private Sub Form_Load()
On Error Resume Next

    SetMainFormItemsEnableWhenPropertyShow False
    
    cmbStyle.Clear
    cmbStyle.AddItem LoadNationalResString(1386)
    cmbStyle.ItemData(cmbStyle.ListCount - 1) = 0
    cmbStyle.AddItem LoadNationalResString(1387)
    cmbStyle.ItemData(cmbStyle.ListCount - 1) = 1
    cmbStyle.AddItem LoadNationalResString(1388)
    cmbStyle.ItemData(cmbStyle.ListCount - 1) = 2
    cmbStyle.AddItem LoadNationalResString(1389)
    cmbStyle.ItemData(cmbStyle.ListCount - 1) = 3
    cmbStyle.AddItem LoadNationalResString(1390)
    cmbStyle.ItemData(cmbStyle.ListCount - 1) = 4
    cmbStyle.AddItem LoadNationalResString(1391)
    cmbStyle.ItemData(cmbStyle.ListCount - 1) = 5
    cmbStyle.AddItem LoadNationalResString(1392)
    cmbStyle.ItemData(cmbStyle.ListCount - 1) = 6
    
    '节点ID
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    
    '节点编号
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    txtIndex = Trim(Str(Node255_Data2.Index))
    nd_start_text = Node255_Data2.StartNode
    nd_end_text = Node255_Data2.EndNode

    cmbStyle.ListIndex = Node255_Data2.Style Mod cmbStyle.ListCount
    txtWidth = Trim(Str(Node255_Data2.Width))
    LineColor.BackColor = Node255_Data2.Color
                
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    
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

Private Sub nd_end_text_Change()
    f_DataChanged = True
End Sub

Private Sub nd_end_text_GotFocus()
    nd_end_text.SelStart = 0
    nd_end_text.SelLength = Len(nd_end_text)
End Sub

Private Sub nd_end_text_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub nd_start_text_Change()
    f_DataChanged = True
End Sub

Private Sub nd_start_text_GotFocus()
    nd_start_text.SelStart = 0
    nd_start_text.SelLength = Len(nd_start_text)
End Sub

Private Sub nd_start_text_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

Private Sub txtIndex_Change()
    f_DataChanged = True
End Sub

Private Sub txtIndex_GotFocus()
    txtIndex.SelStart = 0
    txtIndex.SelLength = Len(txtIndex)
End Sub

Private Sub txtIndex_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtWidth_Change()
    f_DataChanged = True
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = 0
    txtWidth.SelLength = Len(txtWidth)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
