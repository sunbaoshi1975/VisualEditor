VERSION 5.00
Begin VB.Form frm_101 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�û�COM"
   ClientHeight    =   3765
   ClientLeft      =   4920
   ClientTop       =   2790
   ClientWidth     =   3240
   Icon            =   "frm_101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1065"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   2832
      Picture         =   "frm_101.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "1948"
      Top             =   3351
      Width           =   333
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   885
      Left            =   30
      TabIndex        =   19
      Tag             =   "1104"
      Top             =   2340
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
      Caption         =   "����"
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   30
      TabIndex        =   10
      Tag             =   "1136"
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   2640
         Picture         =   "frm_101.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "1145"
         Top             =   1830
         Width           =   315
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   2640
         Picture         =   "frm_101.frx":22C6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "1145"
         Top             =   1470
         Width           =   315
      End
      Begin VB.CommandButton cmdShowRes 
         Height          =   285
         Left            =   2640
         Picture         =   "frm_101.frx":2650
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "1146"
         Top             =   1110
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
         ToolTipText     =   "0-����¼��1-16��¼λ"
         Top             =   720
         Width           =   1395
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   915
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
         TabIndex        =   1
         Top             =   1080
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
         Top             =   1800
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "�ڵ�ID"
         Height          =   195
         Left            =   1380
         TabIndex        =   18
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "COM�ӿ�ID"
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Tag             =   "1168"
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "��������־"
         Height          =   225
         Left            =   180
         TabIndex        =   13
         Tag             =   "1159"
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "�ӽڵ�ID"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Tag             =   "1252"
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label Label20 
         Caption         =   "���ڵ�ID"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Tag             =   "1169"
         Top             =   1500
         Width           =   765
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X�˳�"
      Height          =   375
      Left            =   1350
      TabIndex        =   9
      Tag             =   "1144"
      Top             =   3330
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S����"
      Default         =   -1  'True
      Height          =   375
      Left            =   30
      TabIndex        =   8
      Tag             =   "1007"
      Top             =   3330
      Width           =   1035
   End
End
Attribute VB_Name = "frm_101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//�ļ�����  Frm_101.frm
'//��;��    �����µĽڵ�
'//����:     Scott
'//�������ڣ�2001/09/13
'//�޸����ڣ�
'//�ļ��������û�COM
'//////////////////////////////////////////////////////////////////
Option Explicit

' �����Ƿ��޸�
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
    Set gSystem.crlCurItem = T_com_iid
    frmResourceList.Show vbModal

End Sub

Public Sub Command1_Click()
On Error Resume Next

    If f_DataChanged Then
    
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' ����
        Node101_Data1.reserved1(0) = 0
        
        ' ��������־
        If Cb_log.ListIndex < 0 Then
            Node101_Data1.log = 0
        Else
            Node101_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        ' ����
        Node101_Data1.reserved2(0) = 0
        
        ' ����
        Node101_Data2.reserved1(0) = 0
        
        ' COM�ӿ�ID
        If Trim(T_com_iid) = "" Then
            Node101_Data2.com_iid = 0
        Else
            If CLng(Trim(T_com_iid)) > 32767 Then
                Message ("E088")
                Exit Sub
            Else
                Node101_Data2.com_iid = CLng(Trim(T_com_iid.Text))
            End If
        End If
        
        ' ����
        Node101_Data2.reserved2(0) = 0
        
        ' ���ڵ�
        If Trim(T_nd_parent) = "" Then
           Node101_Data2.nd_parent = 0
        Else
           If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
           Else
              Node101_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
           End If
        End If
        
        '�ӽڵ�
        If Trim(T_nd_child) = "" Then
           Node101_Data2.nd_child = 0
        Else
           If (Val(T_nd_child) > 32767 Or Val(T_nd_child) < 256) And Val(T_nd_child) <> 0 Then
              Message ("E072")
              T_nd_child.SetFocus
              Exit Sub
           Else
              Node101_Data2.nd_child = CInt(Trim(T_nd_child.Text))
           End If
        End If
        
        ' ����
        Node101_Data2.reserved3(0) = 0
        
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
    If Node101_Data2.com_iid > 0 And Node0_Data2.MainCOM = 0 Then
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
  
    ' ������־
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
    T_com_iid.Text = Node101_Data2.com_iid
    T_nd_parent.Text = Node101_Data2.nd_parent
    T_nd_child.Text = Node101_Data2.nd_child
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node101_Data1.log), 0)
        
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

