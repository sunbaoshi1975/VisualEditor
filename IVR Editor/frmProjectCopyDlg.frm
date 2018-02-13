VERSION 5.00
Begin VB.Form frmProjectCopyDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "复制项目选项"
   ClientHeight    =   1350
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4455
   Icon            =   "frmProjectCopyDlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1905"
   Begin VB.CheckBox chkExt 
      Caption         =   "使用扩展选项"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Tag             =   "1911"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制资源"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Tag             =   "1907"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdNoCopy 
      Caption         =   "不复制资源"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1908"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "取消(&Q)"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Tag             =   "1909"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "目标项目属性"
      Height          =   1815
      Left            =   2280
      TabIndex        =   9
      Tag             =   "1901"
      Top             =   1440
      Width           =   2055
      Begin VB.TextBox txtDesPath 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Text            =   "txtDesPath"
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cboDesType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "资源前缀"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Tag             =   "1902"
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "资源类型"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Tag             =   "1903"
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "源项目属性"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Tag             =   "1900"
      Top             =   1440
      Width           =   2055
      Begin VB.TextBox txtSourcePath 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "txtSourcePath"
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cboSourceType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "资源前缀"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Tag             =   "1902"
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "资源类型"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Tag             =   "1903"
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "是否复制此项目包含的资源?"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Tag             =   "1906"
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "frmProjectCopyDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Michael Added @ 2007-11-28
Option Explicit

Private Sub chkExt_Click()
    Call En_DisEnOpt
    
    If chkExt.value = 1 Then
        Me.Height = 3720
        'cmdExt.Caption = LoadResString(1910)
        Me.Refresh
        Exit Sub
    End If
    
    If chkExt.value = 0 Then
        Me.Height = 1725
        'cmdExt.Caption = LoadResString(1909)
        Me.Refresh
        Exit Sub
    End If
    
End Sub

Private Sub cmdNoCopy_Click()
    frmProjectList.bCopyRes = False
    Unload Me
    frmProjectItem.Show vbModal
End Sub

'Private Sub cmdExt_Click()
'    If Right(Trim(cmdExt.Caption), 2) = ">>" Then
'        Me.Height = 3630
'        cmdExt.Caption = LoadResString(1910)
'        Me.Refresh
'        Exit Sub
'    End If
'
'    If Right(Trim(cmdExt.Caption), 2) = "<<" Then
'        Me.Height = 1440
'        cmdExt.Caption = LoadResString(1909)
'        Me.Refresh
'        Exit Sub
'    End If
'
'End Sub

Private Sub cmdCopy_Click()
    frmProjectList.bCopyRes = True
    If chkExt.value = 1 Then
        If cboDesType.ListIndex = 2 Or cboSourceType.ListIndex = 2 Then
            Message "M150"
            Exit Sub
        End If
        frmProjectItem.bCopyOpt = True
        frmProjectItem.lv_strOriFront = Trim(txtSourcePath)
        frmProjectItem.lv_strOriBack = Right(Trim(cboSourceType.Text), 3)
        If Trim(txtDesPath) <> "" Then
            If Right(Trim(txtDesPath), 1) <> "\" Then
                frmProjectItem.lv_strNewFront = Trim(txtDesPath) & "\"
            Else
                frmProjectItem.lv_strNewFront = Trim(txtDesPath)
            End If
        Else
            frmProjectItem.lv_strNewFront = ""
        End If
        frmProjectItem.lv_strNewBack = Right(Trim(cboDesType.Text), 3)
    End If
    
    Unload Me
    frmProjectItem.Show vbModal
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Init the Dialog
    txtSourcePath.Text = ""
    txtDesPath.Text = ""
    chkExt.value = 0
    Call FillComboList
    Call En_DisEnOpt
    
    'Fill the Source Property
    txtSourcePath = frmProjectList.strFrontPath
    txtDesPath = frmProjectList.strFrontPath
    If frmProjectList.strResourceType = "vox" Then
        cboDesType.ListIndex = 0
        cboSourceType.ListIndex = 0
    ElseIf frmProjectList.strResourceType = "wav" Then
        cboDesType.ListIndex = 1
        cboSourceType.ListIndex = 1
'    Else
'        cboDesType.ListIndex = 2
'        cboSourceType.ListIndex = 2
    End If
    
    LoadResStrings Me
End Sub


Private Sub FillComboList()
    cboSourceType.Clear
    cboDesType.Clear
    
    cboSourceType.AddItem ".vox"
    cboSourceType.AddItem ".wav"
'    cboSourceType.AddItem ""
    cboSourceType.ListIndex = 0
    
    cboDesType.AddItem ".vox"
    cboDesType.AddItem ".wav"
 '   cboDesType.AddItem ""
    cboDesType.ListIndex = 0
End Sub

Public Sub En_DisEnOpt()
    If chkExt.value = 0 Then
        cboSourceType.Enabled = False
        cboDesType.Enabled = False
        txtDesPath.Enabled = False
        txtSourcePath.Enabled = False
    Else
        cboSourceType.Enabled = True
        cboDesType.Enabled = True
        txtDesPath.Enabled = True
        txtSourcePath.Enabled = True
    End If
End Sub

