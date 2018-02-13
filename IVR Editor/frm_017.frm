VERSION 5.00
Begin VB.Form frm_017 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择服务语言"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frm_017.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1213"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6672
      Picture         =   "frm_017.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "1948"
      Top             =   3441
      Width           =   333
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   2010
      TabIndex        =   24
      Tag             =   "1007"
      Top             =   3450
      Width           =   1065
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   3480
      TabIndex        =   25
      Tag             =   "1144"
      Top             =   3450
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      Caption         =   "设置1"
      ForeColor       =   &H00000000&
      Height          =   3315
      Left            =   3840
      TabIndex        =   23
      Tag             =   "1164"
      Top             =   60
      Width           =   3165
      Begin VB.Frame famLang 
         Caption         =   "语言选择按键"
         Height          =   1455
         Left            =   30
         TabIndex        =   10
         Tag             =   "1215"
         Top             =   600
         Width           =   3105
         Begin VB.ComboBox Cb_Lang 
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
            Index           =   5
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "请选择"
            Top             =   960
            Width           =   795
         End
         Begin VB.ComboBox Cb_Lang 
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
            Index           =   4
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "请选择"
            Top             =   960
            Width           =   795
         End
         Begin VB.ComboBox Cb_Lang 
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
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "请选择"
            Top             =   630
            Width           =   795
         End
         Begin VB.ComboBox Cb_Lang 
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
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "请选择"
            Top             =   600
            Width           =   795
         End
         Begin VB.ComboBox Cb_Lang 
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
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "请选择"
            Top             =   240
            Width           =   795
         End
         Begin VB.ComboBox Cb_Lang 
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
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "请选择"
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblLang 
            Caption         =   "语言5"
            Height          =   195
            Index           =   5
            Left            =   1650
            TabIndex        =   40
            Tag             =   "1221"
            Top             =   1050
            Width           =   555
         End
         Begin VB.Label lblLang 
            Caption         =   "语言4"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Tag             =   "1220"
            Top             =   1050
            Width           =   555
         End
         Begin VB.Label lblLang 
            Caption         =   "语言3"
            Height          =   195
            Index           =   3
            Left            =   1650
            TabIndex        =   38
            Tag             =   "1219"
            Top             =   690
            Width           =   555
         End
         Begin VB.Label lblLang 
            Caption         =   "语言2"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Tag             =   "1218"
            Top             =   690
            Width           =   555
         End
         Begin VB.Label lblLang 
            Caption         =   "语言1"
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   36
            Tag             =   "1217"
            Top             =   330
            Width           =   555
         End
         Begin VB.Label lblLang 
            Caption         =   "语言0"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Tag             =   "1216"
            Top             =   330
            Width           =   555
         End
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
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2490
         Width           =   915
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
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2850
         Width           =   915
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2670
         Picture         =   "frm_017.frx":1B3C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "显示节点列表"
         Top             =   2520
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   2670
         Picture         =   "frm_017.frx":1EC6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "显示节点列表"
         Top             =   2880
         Width           =   315
      End
      Begin VB.TextBox T_vox_play 
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   8
         Top             =   240
         Width           =   1065
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
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2130
         Width           =   915
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2670
         Picture         =   "frm_017.frx":2250
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1146"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2670
         Picture         =   "frm_017.frx":2352
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "显示节点列表"
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label Lbl_nd_succeed 
         AutoSize        =   -1  'True
         Caption         =   "成功转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Tag             =   "1170"
         Top             =   2580
         Width           =   1065
      End
      Begin VB.Label Lbl_nd_fail 
         AutoSize        =   -1  'True
         Caption         =   "失败转节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Tag             =   "1171"
         Top             =   2940
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "播放语音"
         Height          =   225
         Left            =   180
         TabIndex        =   30
         Tag             =   "1100"
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label20 
         Caption         =   "父节点ID"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Tag             =   "1169"
         Top             =   2220
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   2025
      Left            =   60
      TabIndex        =   5
      Tag             =   "1136"
      Top             =   60
      Width           =   3705
      Begin VB.ComboBox Cb_var_lang 
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
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1530
         Width           =   1965
      End
      Begin VB.ComboBox Cb_maxtrytime 
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
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "请选择"
         Top             =   1140
         Width           =   1965
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
         Left            =   2610
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
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
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
         Left            =   1590
         TabIndex        =   2
         ToolTipText     =   "0 - 255 秒"
         Top             =   750
         Width           =   1965
      End
      Begin VB.Label Lbl_var_password 
         AutoSize        =   -1  'True
         Caption         =   "选择结果记录"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Tag             =   "1214"
         Top             =   1590
         Width           =   1080
      End
      Begin VB.Label Lbl_trytime 
         AutoSize        =   -1  'True
         Caption         =   "最大尝试次数"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Tag             =   "1158"
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label9 
         Caption         =   "节点超时(秒)"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   28
         Tag             =   "1154"
         Top             =   840
         Width           =   1125
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
         Left            =   1620
         TabIndex        =   26
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "描述"
      Height          =   1215
      Left            =   60
      TabIndex        =   7
      Tag             =   "1104"
      Top             =   2160
      Width           =   3705
      Begin VB.TextBox Txt_Description 
         Height          =   885
         Left            =   90
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frm_017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_017.frm
'//用途：    创建新的节点
'//作者:     Sun
'//创建日期：2003/04/25
'//修改日期：
'//文件描述：选择服务语言
'//////////////////////////////////////////////////////////////////
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub Cb_Lang_Click(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub Cb_maxtrytime_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Change()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_Click()
    f_DataChanged = True
End Sub

Private Sub Cb_timeout_GotFocus()
    Cb_timeout.SelStart = 0
    Cb_timeout.SelLength = Len(Cb_timeout)
End Sub

Private Sub Cb_var_lang_Click()
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

Private Sub cmdShowRes_Click()
    gSystem.intCurStep = 1
    Set gSystem.crlCurItem = T_vox_play
    frmResourceList.Show vbModal
End Sub

Private Sub CommandExit_Click()
    Unload Me
End Sub

Private Sub CommandSave_Click()
On Error Resume Next

    Dim lv_loop
    
    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        '节点超时
        If Len(Trim(Cb_timeout)) < 1 Or Trim(Cb_timeout) = "" Then
            Node17_Data1.timeout = 0
        Else
            If CLng(Trim(Cb_timeout)) > 255 Then
                Message ("E051")
                Cb_timeout.SetFocus
                Exit Sub
            Else
                Node17_Data1.timeout = CByte(Val(Cb_timeout) Mod 256)
            End If
        End If
        
        '最大尝试次数
        If Cb_maxtrytime.ListIndex < 0 Then
            Node17_Data1.maxtrytime = 3
        Else
            Node17_Data1.maxtrytime = CByte(Cb_maxtrytime.ItemData(Cb_maxtrytime.ListIndex))
        End If
        
        '语言选择记录
        If Cb_var_lang.ListIndex <= 0 Then
            Node17_Data1.var_lang = 0
        Else
            Node17_Data1.var_lang = CByte(Cb_var_lang.ItemData(Cb_var_lang.ListIndex))
        End If
        
        Node17_Data1.reserved1(0) = 0
        
        
        '提示语音
        If Trim(T_vox_play.Text) = "" Then
            Node17_Data2.vox_play = 0
        Else
            If Val(T_vox_play) > 32767 Then
                Message ("E076")
                T_vox_play.SetFocus
                Exit Sub
            Else
                Node17_Data2.vox_play = Val(T_vox_play.Text)
            End If
        End If
        
        '语言选择按键
        For lv_loop = Cb_Lang.LBound To Cb_Lang.UBound
            With Cb_Lang(lv_loop)
                If .ListIndex < 0 Then
                    Node17_Data2.lang(lv_loop) = 0
                Else
                    Node17_Data2.lang(lv_loop) = CByte(.ItemData(.ListIndex))
                End If
            End With
        Next
        
        Node17_Data2.reserved1(0) = 0
        
        
        '父节点
        If Trim(T_nd_parent.Text) = "" Then
           Node17_Data2.nd_parent = 0
        Else
           If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
           Else
                Node17_Data2.nd_parent = Val(T_nd_parent.Text)
           End If
        End If
        
        'Succeed node id
        If Trim(T_nd_succeed.Text) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_succeed) > 32767 Or Val(T_nd_succeed) < 256) And Val(T_nd_succeed) <> 0 Then
                Message ("E041")
                T_nd_succeed.SetFocus
                Exit Sub
            Else
                lv_nNewNode = Val(T_nd_succeed.Text)
            End If
        End If
        
        '' Sun added 2007-03-25
        If Node17_Data2.nd_succ <> lv_nNewNode Then
            Node17_Data2.nd_succ = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 1)
        End If
        
        'failed node id
        If Trim(T_nd_fail.Text) = "" Then
            lv_nNewNode = 0
        Else
            If (Val(T_nd_fail) > 32767 Or Val(T_nd_fail) < 256) And Val(T_nd_fail) <> 0 Then
                Message ("E042")
                T_nd_fail.SetFocus
                Exit Sub
           Else
              lv_nNewNode = Val(T_nd_fail.Text)
           End If
        End If
        
        '' Sun added 2007-03-25
        If Node17_Data2.nd_fail <> lv_nNewNode Then
            Node17_Data2.nd_fail = lv_nNewNode
            Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, 0)
        End If
        
        Node17_Data2.reserved2(0) = 0
        
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

    Dim lv_loop
    Dim i As Integer
    
    '节点超时
    With Cb_timeout
        For i = 0 To 100 Step 5
            .AddItem Trim(Str(i))
        Next
    End With

    '最大尝试次数
    With Cb_maxtrytime
        For i = 0 To 9
            .AddItem Trim(Str(i)) & LoadNationalResString(1180)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
    '语言记录
    RefreshVariablesList Cb_var_lang
    
    '语言选择按键
    For lv_loop = Cb_Lang.LBound To Cb_Lang.UBound
        With Cb_Lang(lv_loop)
            .AddItem LoadNationalResString(1222)
            .ItemData(.ListCount - 1) = 0
            For i = 0 To 9
                .AddItem Trim(Str(i)) & LoadNationalResString(1172)
                .ItemData(.ListCount - 1) = 48 + i
            Next
            .AddItem LoadNationalResString(1176)
            .ItemData(.ListCount - 1) = 42
            .AddItem LoadNationalResString(1177)
            .ItemData(.ListCount - 1) = 35
            
            If lv_loop >= Node0_Data1.Languages Then
                .Enabled = False
            End If
        End With
    Next
    
    '节点ID
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    '节点编号
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
   
    Cb_timeout.Text = Node17_Data1.timeout
    Cb_maxtrytime.ListIndex = SearchItemDataIndex(Cb_maxtrytime, _
                                    CLng(Node17_Data1.maxtrytime), 3)
    Cb_var_lang.ListIndex = SearchItemDataIndex(Cb_var_lang, CLng(Node17_Data1.var_lang), 0)
    
    T_vox_play.Text = Node17_Data2.vox_play
    T_nd_parent.Text = Node17_Data2.nd_parent
    T_nd_succeed.Text = Node17_Data2.nd_succ
    T_nd_fail.Text = Node17_Data2.nd_fail
          
    For lv_loop = Cb_Lang.LBound To Cb_Lang.UBound
        Cb_Lang(lv_loop).ListIndex = SearchItemDataIndex(Cb_Lang(lv_loop), _
                                    CLng(Node17_Data2.lang(lv_loop)), 0)
    Next
    
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

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
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

Private Sub T_vox_play_Change()
    f_DataChanged = True

    '' Sun added 2002-09-10
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip T_vox_play

End Sub

Private Sub T_vox_play_GotFocus()
    T_vox_play.SelStart = 0
    T_vox_play.SelLength = Len(T_vox_play)
    
    '' Sun added 2002-04-02
    gintSoundResourceID = Val(T_vox_play)
    Call SoundResourceIDChanged
    
End Sub

Private Sub T_vox_play_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub
