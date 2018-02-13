VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm_buffer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户Buffer定义变量"
   ClientHeight    =   4485
   ClientLeft      =   4545
   ClientTop       =   3045
   ClientWidth     =   5085
   Icon            =   "frm_buffer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1351"
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   1410
      TabIndex        =   16
      Tag             =   "1007"
      Top             =   4080
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   2850
      TabIndex        =   17
      Tag             =   "1144"
      Top             =   4080
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   3915
      Left            =   60
      TabIndex        =   18
      Tag             =   "1136"
      Top             =   60
      Width           =   4965
      Begin VB.Frame famVar 
         Caption         =   "变量4"
         Height          =   1425
         Index           =   3
         Left            =   2550
         TabIndex        =   37
         Tag             =   "1355"
         Top             =   2340
         Width           =   2265
         Begin VB.TextBox txtDataLength 
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
            Index           =   3
            Left            =   990
            TabIndex        =   12
            Top             =   210
            Width           =   915
         End
         Begin VB.ComboBox Cb_datatype 
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
            Index           =   3
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   570
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox T_var_name 
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
            Index           =   3
            Left            =   990
            MaxLength       =   15
            TabIndex        =   15
            ToolTipText     =   "请输入字母或数字"
            Top             =   930
            Width           =   1155
         End
         Begin ComCtl2.UpDown udnDataLength 
            Height          =   360
            Index           =   3
            Left            =   1905
            TabIndex        =   13
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDataLength(3)"
            BuddyDispid     =   196613
            BuddyIndex      =   3
            OrigLeft        =   1890
            OrigTop         =   210
            OrigRight       =   2130
            OrigBottom      =   570
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "变量名称"
            Height          =   225
            Index           =   3
            Left            =   150
            TabIndex        =   40
            Tag             =   "1358"
            Top             =   930
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "变量类型"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   39
            Tag             =   "1357"
            Top             =   600
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label20 
            Caption         =   "变量长度"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   38
            Tag             =   "1356"
            Top             =   270
            Width           =   765
         End
      End
      Begin VB.Frame famVar 
         Caption         =   "变量3"
         Height          =   1425
         Index           =   2
         Left            =   150
         TabIndex        =   33
         Tag             =   "1354"
         Top             =   2340
         Width           =   2265
         Begin VB.TextBox txtDataLength 
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
            Index           =   2
            Left            =   990
            TabIndex        =   8
            Top             =   210
            Width           =   915
         End
         Begin VB.ComboBox Cb_datatype 
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
            Index           =   2
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   570
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox T_var_name 
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
            Index           =   2
            Left            =   990
            MaxLength       =   15
            TabIndex        =   11
            ToolTipText     =   "请输入字母或数字"
            Top             =   930
            Width           =   1155
         End
         Begin ComCtl2.UpDown udnDataLength 
            Height          =   360
            Index           =   2
            Left            =   1905
            TabIndex        =   9
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDataLength(2)"
            BuddyDispid     =   196613
            BuddyIndex      =   2
            OrigLeft        =   1890
            OrigTop         =   210
            OrigRight       =   2130
            OrigBottom      =   570
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "变量名称"
            Height          =   225
            Index           =   2
            Left            =   150
            TabIndex        =   36
            Tag             =   "1358"
            Top             =   930
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "变量类型"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   35
            Tag             =   "1357"
            Top             =   600
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label20 
            Caption         =   "变量长度"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   34
            Tag             =   "1356"
            Top             =   270
            Width           =   765
         End
      End
      Begin VB.Frame famVar 
         Caption         =   "变量2"
         Height          =   1425
         Index           =   1
         Left            =   2550
         TabIndex        =   29
         Tag             =   "1353"
         Top             =   780
         Width           =   2265
         Begin VB.TextBox txtDataLength 
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
            Left            =   990
            TabIndex        =   4
            Top             =   210
            Width           =   915
         End
         Begin VB.ComboBox Cb_datatype 
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
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   570
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox T_var_name 
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
            Left            =   990
            MaxLength       =   15
            TabIndex        =   7
            ToolTipText     =   "请输入字母或数字"
            Top             =   930
            Width           =   1155
         End
         Begin ComCtl2.UpDown udnDataLength 
            Height          =   360
            Index           =   1
            Left            =   1905
            TabIndex        =   5
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDataLength(1)"
            BuddyDispid     =   196613
            BuddyIndex      =   1
            OrigLeft        =   1890
            OrigTop         =   210
            OrigRight       =   2130
            OrigBottom      =   570
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "变量名称"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   32
            Tag             =   "1358"
            Top             =   930
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "变量类型"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   31
            Tag             =   "1357"
            Top             =   600
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label20 
            Caption         =   "变量长度"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   30
            Tag             =   "1356"
            Top             =   270
            Width           =   765
         End
      End
      Begin VB.Frame famVar 
         Caption         =   "变量1"
         Height          =   1425
         Index           =   0
         Left            =   150
         TabIndex        =   25
         Tag             =   "1352"
         Top             =   780
         Width           =   2265
         Begin VB.TextBox txtDataLength 
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
            Left            =   990
            TabIndex        =   0
            Top             =   210
            Width           =   915
         End
         Begin ComCtl2.UpDown udnDataLength 
            Height          =   360
            Index           =   0
            Left            =   1905
            TabIndex        =   1
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDataLength(0)"
            BuddyDispid     =   196613
            BuddyIndex      =   0
            OrigLeft        =   1890
            OrigTop         =   210
            OrigRight       =   2130
            OrigBottom      =   570
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox T_var_name 
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
            Left            =   990
            MaxLength       =   15
            TabIndex        =   3
            ToolTipText     =   "请输入字母或数字"
            Top             =   930
            Width           =   1155
         End
         Begin VB.ComboBox Cb_datatype 
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
            ItemData        =   "frm_buffer.frx":0E42
            Left            =   990
            List            =   "frm_buffer.frx":0E44
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   570
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label20 
            Caption         =   "变量长度"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   28
            Tag             =   "1356"
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label19 
            Caption         =   "变量类型"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   27
            Tag             =   "1357"
            Top             =   600
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "变量名称"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   26
            Tag             =   "1358"
            Top             =   930
            Width           =   855
         End
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
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Tag             =   "1143"
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Tag             =   "1137"
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "描述"
      Height          =   1125
      Left            =   2250
      TabIndex        =   19
      Tag             =   "1104"
      Top             =   2940
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox Txt_Description 
         Height          =   705
         Left            =   150
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   270
         Visible         =   0   'False
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frm_buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_101.frm
'//用途：    创建新的节点
'//作者:     Scott
'//创建日期：2001/09/13
'//修改日期：
'//文件描述：用户BUFFER定义-变量
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_datatype_Click(Index As Integer)
    f_DataChanged = True
End Sub

Public Sub Command1_Click()
On Error Resume Next

Dim lv_loop    As Integer
Dim lv_SubLoop As Integer
Dim lv_Len     As Integer

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        For lv_loop = Cb_datatype.LBound To Cb_datatype.UBound
        
            ' 变量长度
            If Trim(txtDataLength(lv_loop)) = "" Then
                Node2_Data2.uservar(lv_loop * 16) = 0
            Else
                If Val(txtDataLength(lv_loop)) > 255 Then
                    Message ("E078")
                    txtDataLength(lv_loop).SetFocus
                    Exit Sub
                Else
                    Node2_Data2.uservar(lv_loop * 16) = CByte(Val(txtDataLength(lv_loop)) Mod 256)
                End If
            End If
            
            ' 变量数据类型
            If Cb_datatype(lv_loop).ListIndex = -1 Then
                If Node2_Data2.uservar(lv_loop * 16) > 0 Then
                    Message ("E058")
                    Cb_datatype(lv_loop).SetFocus
                    Exit Sub
                Else
                    Node2_Data2.uservar(1 + lv_loop * 16) = 0
                End If
            Else
                Node2_Data2.uservar(1 + lv_loop * 16) = Cb_datatype(lv_loop).ItemData(Cb_datatype(lv_loop).ListIndex)
            End If
            
            ' 变量名称
            If Trim(T_var_name(lv_loop)) = "" Then
                Node2_Data2.uservar(2 + lv_loop * 16) = 0
            Else
                lv_Len = Len(Trim(T_var_name(lv_loop).Text))
                If lv_Len > 14 Then lv_Len = 14
                    
                For lv_SubLoop = 2 To 15
                    If lv_SubLoop <= lv_Len + 1 Then
                        Node2_Data2.uservar(lv_loop * 16 + lv_SubLoop) = Asc(Mid(Trim(T_var_name(lv_loop)), lv_SubLoop - 1, 1))
                    Else
                        Node2_Data2.uservar(lv_loop * 16 + lv_SubLoop) = 0
                    End If
                Next
            End If
            
        Next
        
        If Trim(Txt_Description.Text) = "" Then
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
           gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no)
        f_DataChanged = False
        
    End If
    
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim lv_loop As Integer
Dim lv_SubLoop As Integer
Dim lv_conten As String

    For lv_loop = Cb_datatype.LBound To Cb_datatype.UBound
        With Cb_datatype(lv_loop)
            .AddItem LoadNationalResString(1359)
            .ItemData(.ListCount - 1) = 1
            .AddItem LoadNationalResString(1360)
            .ItemData(.ListCount - 1) = 2
        End With
        
        udnDataLength(lv_loop).Min = 0
        udnDataLength(lv_loop).Max = 255
        udnDataLength(lv_loop).Increment = 1
    Next
   
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
      
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    For lv_loop = Cb_datatype.LBound To Cb_datatype.UBound
        
        udnDataLength(lv_loop).value = Node2_Data2.uservar(lv_loop * 16)
        Cb_datatype(lv_loop).ListIndex = SearchItemDataIndex(Cb_datatype(lv_loop), CLng(Node2_Data2.uservar(1 + lv_loop * 16)), 1)
        
        T_var_name(lv_loop) = ""
        For lv_SubLoop = 2 To 15
            T_var_name(lv_loop) = T_var_name(lv_loop) & Chr(Node2_Data2.uservar(lv_SubLoop + lv_loop * 16))
        Next
        
    Next
    
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

Private Sub T_var_name_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub T_var_name_GotFocus(Index As Integer)
    T_var_name(Index).SelStart = 0
    T_var_name(Index).SelLength = Len(T_var_name(Index))
End Sub

' Only accept letters, numbers and "_"
'
Private Sub T_var_name_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    
    If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
    
    '' Sun added 2005-09-12
    If KeyAscii < 30 Or KeyAscii > 127 Then Exit Sub
    
    If KeyAscii <> vbKeyTab Then
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 95 Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtDataLength_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub txtDataLength_GotFocus(Index As Integer)
    txtDataLength(Index).SelStart = 0
    txtDataLength(Index).SelLength = Len(txtDataLength(Index))
End Sub

Private Sub txtDataLength_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub udnDataLength_Change(Index As Integer)
    f_DataChanged = True
End Sub
