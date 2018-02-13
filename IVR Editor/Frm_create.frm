VERSION 5.00
Begin VB.Form Frm_FlowCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "流程创建"
   ClientHeight    =   4065
   ClientLeft      =   4350
   ClientTop       =   2460
   ClientWidth     =   2880
   Icon            =   "Frm_create.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CommandCreate 
      Caption         =   "创建(&C)"
      Height          =   315
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "保存流程、节点定义"
      Top             =   3630
      Width           =   1065
   End
   Begin VB.CommandButton exit_Command 
      Caption         =   "退出(&E)"
      Height          =   315
      Left            =   1440
      TabIndex        =   17
      ToolTipText     =   "返回主菜单"
      Top             =   3630
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "创建流程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3405
      Left            =   128
      TabIndex        =   0
      Top             =   60
      Width           =   2625
      Begin VB.TextBox Txt_pid 
         Height          =   345
         Left            =   1050
         TabIndex        =   8
         Top             =   330
         Width           =   1395
      End
      Begin VB.TextBox Txt_pname 
         Height          =   345
         Left            =   1050
         TabIndex        =   7
         Top             =   690
         Width           =   1395
      End
      Begin VB.TextBox Txt_pdescription 
         Height          =   345
         Left            =   1050
         TabIndex        =   6
         Text            =   "未定义"
         Top             =   1050
         Width           =   1395
      End
      Begin VB.TextBox Txt_pversion 
         Height          =   345
         Left            =   1050
         TabIndex        =   5
         Top             =   1410
         Width           =   1395
      End
      Begin VB.TextBox Txt_pauther 
         Height          =   345
         Left            =   1050
         TabIndex        =   4
         Top             =   1770
         Width           =   1395
      End
      Begin VB.TextBox Txt_puser 
         Height          =   345
         Left            =   1050
         TabIndex        =   3
         Top             =   2130
         Width           =   1395
      End
      Begin VB.TextBox Txt_pcreatetime 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1050
         TabIndex        =   2
         Top             =   2490
         Width           =   1395
      End
      Begin VB.TextBox Txt_pmodifytime 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1050
         TabIndex        =   1
         Top             =   2850
         Width           =   1395
      End
      Begin VB.Label p_id_lable 
         AutoSize        =   -1  'True
         Caption         =   "流程ID:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   420
         Width           =   570
      End
      Begin VB.Label p_name_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程名称:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   780
         Width           =   765
      End
      Begin VB.Label p_desc_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程描述:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1140
         Width           =   765
      End
      Begin VB.Label p_ver_Label 
         AutoSize        =   -1  'True
         Caption         =   "版本号:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label p_auth_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程作者:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label p_user_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程用户:"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   2220
         Width           =   765
      End
      Begin VB.Label p_crea_Label 
         AutoSize        =   -1  'True
         Caption         =   "创建时间:"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   2550
         Width           =   765
      End
      Begin VB.Label p_modi_Label 
         AutoSize        =   -1  'True
         Caption         =   "修改时间:"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   2940
         Width           =   765
      End
   End
End
Attribute VB_Name = "Frm_FlowCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Create New Flow and Update old Flow
'//文件名：  Frm_FlowCreate.frm
'//用途：    创建新的流程
'//作者:     Scott
'//创建日期：2001/08/29
'//修改日期：
'//文件描述：Create New Flows,So it save data into tabrefivr table
'//////////////////////////////////////////////////////////////////
Option Explicit
Public Sub CommandCreate_Click()
Dim lv_str_sql1 As String
Dim lv_int_row1 As Integer
Dim lv_rs_get1 As New ADODB.Recordset
Dim node As node
Set M_Cn = New ADODB.Connection
      With M_Cn
          .ConnectionString = "DSN=dbcallcenter;UID=sa;PASSWORD="
          '.ConnectionString = "provider=Microsoft.Jet.OLEDB.3.51;data source=" & App.Path & "\data\controls.mdb"
          .Open
      End With
On Error GoTo errFind
'Flow's information
      gGetUserFlow.IntP_id = Txt_pid.Text
      gGetUserFlow.StrP_name = Txt_pname.Text
      gGetUserFlow.StrP_description = Txt_pdescription.Text
      gGetUserFlow.StrP_auther = Txt_pauther.Text
      gGetUserFlow.StrP_user = Txt_pauther.Text
      gGetUserFlow.StrP_version = Txt_pversion.Text
      gGetUserFlow.DataP_modifytime = Now()
      gGetUserFlow.DateP_createtime = Txt_pcreatetime.Text
'Update flow's information  auther:Scott data:2001/08/29
'Save flow's information into tbrefivr table
If CommandCreate.Caption = "&S保存" Then
      M_Cn.Execute "update tbrefivr set p_name='" & _
                 gGetUserFlow.StrP_name & "'," & "p_description='" & _
                 gGetUserFlow.StrP_description & "'," & "p_version='" & _
                 gGetUserFlow.StrP_version & "'," & "p_modifytime='" & _
                 gGetUserFlow.DataP_modifytime & "'," & "p_createtime='" & _
                 CDate(gGetUserFlow.DateP_createtime) & "'," & "p_auther='" & _
                 gGetUserFlow.StrP_auther & "'," & "p_user='" & _
                 gGetUserFlow.StrP_user & "'" & " where p_id=" & CInt(Trim(gGetUserFlow.IntP_id))
                 
      Unload Me
Else
'Add new flows
lv_str_sql1 = "select * from tbrefivr where tbrefivr.p_id =" & CInt(Trim(gGetUserFlow.IntP_id))
         lv_rs_get1.Open lv_str_sql1, M_Cn, adOpenStatic, adLockOptimistic
         lv_int_row1 = lv_rs_get1.RecordCount
If lv_int_row1 > 0 Then
   Message ("M126")
Else
   M_Cn.Execute "Insert into tbrefivr (p_id,p_name,p_description,p_version,p_modifytime,p_createtime,p_auther,p_user) values (" & gGetUserFlow.IntP_id & _
             ",'" & gGetUserFlow.StrP_name & "','" & gGetUserFlow.StrP_description & "','" & gGetUserFlow.StrP_version & "','" & gGetUserFlow.DataP_modifytime & "','" & gGetUserFlow.DateP_createtime & "','" & gGetUserFlow.StrP_auther & "','" & gGetUserFlow.StrP_user & "'" & ")"
   gFlowNo = Txt_pid.Text
'Create node name id Node0_data2 and Node1_data1
'   F_CreateFlow gGetUserFlow.IntP_id
'   CFlowWorks.InitNewFlow
   Unload Me
End If
End If
errFind:
    If Err = -2147467259 Then
       Set M_Cn = Nothing
       Set M_Cn = New ADODB.Connection
        M_Cn.ConnectionString = "provider=Microsoft.Jet.OLEDB.3.51;data source=" & App.Path & "\controls.mdb"
        M_Cn.Open
        Resume Next
    ElseIf Err <> 0 Then ' 其他的错误
        MsgBox "不期望的错误: " & Err.Description
        End
    End If
End Sub
Private Sub exit_Command_Click()
If Message("Q005") = vbYes Then
   Unload Me
Else
   Exit Sub
End If
End Sub
Private Sub Form_Load()
'Flow information
   Dim Rs As ADODB.Recordset
   Set Rs = New ADODB.Recordset
   With Rs
         Set .ActiveConnection = M_Cn
         .CursorType = adOpenKeyset
         .LockType = adLockOptimistic
         .Open "Select * from tbrefivr where p_id=" & gFlowNo
   End With
   If Rs.RecordCount = 1 Then
      F_GetNodedata Rs
      Frm_FlowCreate.Caption = "流程编辑"
      Frame1.Caption = "流程号-" & gFlowNo
      Txt_pid = gFlowNo
      Txt_pdescription = gGetUserFlow.StrP_description
      Txt_pname = gGetUserFlow.StrP_name
      Txt_pauther = gGetUserFlow.StrP_auther
      Txt_puser = gGetUserFlow.StrP_user
      Txt_pversion = gGetUserFlow.StrP_version
      Txt_pcreatetime = gGetUserFlow.DateP_createtime
      Txt_pmodifytime = gGetUserFlow.DataP_modifytime
   Else
      If Rs.RecordCount = 0 Then
         Frm_FlowCreate.Caption = "流程编辑"
         Frame1.Caption = "流程号-" & gFlowNo
         Txt_pid = gFlowNo
         Txt_pdescription = "N/A"
         Txt_pname = "N/A"
         Txt_pauther = "N/A"
         Txt_puser = "N/A"
         Txt_pversion = "V1.00"
         Txt_pcreatetime = Now()
         Txt_pmodifytime = Now()
       End If
   End If
End Sub
