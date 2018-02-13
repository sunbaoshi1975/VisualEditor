VERSION 5.00
Begin VB.Form Frm_NodeDelete 
   Caption         =   "节点删除"
   ClientHeight    =   1710
   ClientLeft      =   4605
   ClientTop       =   2760
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   2355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "流程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   135
      TabIndex        =   2
      Top             =   60
      Width           =   2085
      Begin VB.TextBox T_n_id 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox load_text 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "节点ID："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "流程ID："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   4
         Top             =   330
         Width           =   720
      End
   End
   Begin VB.CommandButton Comm_load 
      Caption         =   "&D删除"
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   1260
      Width           =   885
   End
   Begin VB.CommandButton Comm_exit 
      Caption         =   "&Q退出"
      Height          =   315
      Left            =   1245
      TabIndex        =   0
      Top             =   1260
      Width           =   885
   End
End
Attribute VB_Name = "Frm_NodeDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Private Sub Comm_exit_Click()
'Unload Me
'End Sub
'
'Private Sub Comm_load_Click()
'Dim lv_rs_select As New ADODB.Recordset
'Dim lv_str_sql As String
'Dim lv_int_row As Integer
'Dim lv_i As Integer
'lv_str_sql = "select * from tbivrprogram  where tbivrprogram.p_id=" & CInt(Trim(load_text.Text)) & " and n_id=" & CInt(Trim(T_n_id.Text))
'lv_rs_select.Open lv_str_sql, M_Cn, adOpenStatic, adLockOptimistic
'lv_int_row = lv_rs_select.RecordCount
'lv_rs_select.MoveFirst
'For lv_i = 0 To lv_int_row
'      lv_rs_select.Delete
'      lv_rs_select.Update
'      lv_rs_select.MoveNext
'      AddItemsOnly CInt(Trim(load_text.Text))
''   Unload Me
'Next
'End Sub
'
