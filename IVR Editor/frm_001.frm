VERSION 5.00
Begin VB.Form frm_001 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户Buffer定义日志"
   ClientHeight    =   1800
   ClientLeft      =   4290
   ClientTop       =   2550
   ClientWidth     =   3285
   Icon            =   "frm_001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1148"
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Tag             =   "1144"
      Top             =   1380
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   450
      TabIndex        =   4
      Tag             =   "1007"
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Tag             =   "1136"
      Top             =   90
      Width           =   3135
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   6
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
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox T_uservars 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "0"
         ToolTipText     =   "0-255"
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1350
         TabIndex        =   7
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_uservars 
         Caption         =   "用户定义变量数"
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Tag             =   "1149"
         Top             =   780
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1095
      Left            =   60
      TabIndex        =   9
      Tag             =   "1104"
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox Txt_Description 
         Height          =   705
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frm_001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_001.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：用户BUFFER定义-日志
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub CommandExit_Click()
    Unload Me
End Sub

Public Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then

        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        '保留
        Node1_Data1.reserved1(0) = 0
        
        '用户变量定义
        If Trim(T_uservars) = "" Then
           Node1_Data1.uservars = 0
        Else
           If Val(T_uservars) > 255 Then
              Message ("E036")
              T_uservars.SetFocus
              Exit Sub
           Else
              Node1_Data1.uservars = CByte(Trim(T_uservars.Text))
           End If
        End If
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        '节点数据整和
        F_NodeData 2, 1
           
        '保存节点
        gCallFlow.UpdateAnotherIVRRecord 2
        
        f_DataChanged = False
        
        If Val(T_uservars) >= 0 Then
           F_CreateVar Val(T_uservars)
        End If
       
    End If
      
    Unload Me
    
End Sub

Public Sub Form_Load()
On Error Resume Next

    '节点ID
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    
    '节点编号
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    T_uservars.Text = Node1_Data1.uservars    '用户定义变量数
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

Private Sub T_uservars_Change()
    f_DataChanged = True
End Sub

Private Sub T_uservars_GotFocus()
    T_uservars.SelStart = 0
    T_uservars.SelLength = Len(T_uservars)
End Sub

Private Sub T_uservars_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub
