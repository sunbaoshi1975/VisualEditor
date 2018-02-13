VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   8760
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   15270
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1420"
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   6000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   5850
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame famOpt 
         Height          =   3375
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   4005
         Begin VB.TextBox txtODBCDBName 
            Height          =   270
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   26
            Top             =   2130
            Width           =   2055
         End
         Begin VB.OptionButton optDataSource 
            Caption         =   "使用 ODBC"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   24
            Tag             =   "1433"
            Top             =   1500
            Width           =   1455
         End
         Begin VB.OptionButton optDataSource 
            Caption         =   "使用 OLE DB"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   21
            Tag             =   "1430"
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox txtDBName 
            Height          =   270
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   23
            Top             =   990
            Width           =   2055
         End
         Begin VB.TextBox txtServer 
            Height          =   270
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   22
            Top             =   630
            Width           =   2055
         End
         Begin VB.TextBox txtDSN 
            Height          =   285
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   25
            Top             =   1785
            Width           =   2055
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1650
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   28
            Top             =   2805
            Width           =   2055
         End
         Begin VB.TextBox txtUserID 
            Height          =   285
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   27
            Top             =   2445
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DB Name"
            Height          =   180
            Index           =   6
            Left            =   570
            TabIndex        =   36
            Tag             =   "1432"
            Top             =   2175
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DB Name"
            Height          =   180
            Index           =   5
            Left            =   570
            TabIndex        =   33
            Tag             =   "1432"
            Top             =   1035
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ServerIP"
            Height          =   195
            Index           =   4
            Left            =   570
            TabIndex        =   32
            Tag             =   "1431"
            Top             =   675
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ODBC DSN"
            Height          =   195
            Index           =   3
            Left            =   570
            TabIndex        =   31
            Tag             =   "1434"
            Top             =   1830
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "口  令"
            Height          =   180
            Index           =   2
            Left            =   570
            TabIndex        =   30
            Tag             =   "1436"
            Top             =   2850
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "用户名"
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   29
            Tag             =   "1435"
            Top             =   2490
            Width           =   540
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   60
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame Frame4 
         Caption         =   "节点属性"
         Height          =   735
         Left            =   2760
         TabIndex        =   47
         Tag             =   "1437"
         Top             =   90
         Width           =   2385
         Begin VB.PictureBox NodeHandColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1020
            ScaleHeight     =   255
            ScaleWidth      =   465
            TabIndex        =   41
            Top             =   270
            Width           =   495
         End
         Begin VB.CommandButton CustNodeHand 
            Caption         =   "..."
            Height          =   285
            Left            =   1560
            TabIndex        =   42
            ToolTipText     =   "Custom Color"
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "颜色:"
            Height          =   180
            Left            =   150
            TabIndex        =   57
            Tag             =   "1438"
            Top             =   300
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "画布"
         Height          =   2295
         Left            =   120
         TabIndex        =   52
         Tag             =   "1440"
         Top             =   960
         Width           =   5025
         Begin VB.CheckBox chkShowSysNodes 
            Caption         =   "显示系统节点"
            Height          =   315
            Left            =   2400
            TabIndex        =   46
            Tag             =   "1730"
            Top             =   300
            Width           =   2415
         End
         Begin VB.CommandButton custwpbkcolor 
            Caption         =   "..."
            Height          =   285
            Left            =   1530
            TabIndex        =   45
            ToolTipText     =   "Custom Color"
            Top             =   300
            Width           =   345
         End
         Begin VB.PictureBox wpbkcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   990
            ScaleHeight     =   255
            ScaleWidth      =   465
            TabIndex        =   44
            Top             =   300
            Width           =   495
         End
         Begin VB.Frame Frame2 
            Caption         =   "区域"
            Height          =   795
            Left            =   120
            TabIndex        =   53
            Tag             =   "1441"
            Top             =   720
            Width           =   4605
            Begin VB.CommandButton cmdDefArea 
               Caption         =   "使用默认值(&U)"
               Height          =   285
               Left            =   3120
               TabIndex        =   62
               Tag             =   "1943"
               Top             =   270
               Width           =   1335
            End
            Begin VB.TextBox txtHeight 
               Height          =   285
               Left            =   2040
               TabIndex        =   49
               Top             =   270
               Width           =   945
            End
            Begin VB.TextBox txtWidth 
               Height          =   285
               Left            =   480
               TabIndex        =   48
               Top             =   270
               Width           =   945
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "高:"
               Height          =   180
               Left            =   1680
               TabIndex        =   38
               Tag             =   "1443"
               Top             =   330
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "宽:"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Tag             =   "1442"
               Top             =   330
               Width           =   270
            End
         End
         Begin VB.TextBox txtPageCount 
            Height          =   285
            Left            =   540
            TabIndex        =   50
            Top             =   1710
            Width           =   615
         End
         Begin ComCtl2.UpDown UpdPageCount 
            Height          =   285
            Left            =   1140
            TabIndex        =   51
            Top             =   1710
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "txtPageCount"
            BuddyDispid     =   196632
            OrigLeft        =   930
            OrigTop         =   1830
            OrigRight       =   1170
            OrigBottom      =   2115
            Max             =   50
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "颜色:"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   56
            Tag             =   "1438"
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "页"
            Height          =   180
            Index           =   2
            Left            =   1470
            TabIndex        =   55
            Tag             =   "1133"
            Top             =   1740
            Width           =   180
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "共"
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   54
            Tag             =   "1444"
            Top             =   1740
            Width           =   180
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "背景"
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Tag             =   "1439"
         Top             =   90
         Width           =   2445
         Begin VB.PictureBox wrkbkcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1050
            ScaleHeight     =   255
            ScaleWidth      =   465
            TabIndex        =   39
            Top             =   270
            Width           =   495
         End
         Begin VB.CommandButton custcolor 
            Caption         =   "..."
            Height          =   285
            Left            =   1590
            TabIndex        =   40
            ToolTipText     =   "Custom Color"
            Top             =   270
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "颜色:"
            Height          =   180
            Left            =   180
            TabIndex        =   5
            Tag             =   "1438"
            Top             =   300
            Width           =   450
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   240
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.ComboBox cboVoiceFileType 
         Height          =   300
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox cboRecFileType 
         Height          =   300
         Left            =   4410
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   345
         Index           =   1
         Left            =   5280
         Picture         =   "frmOptions.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "1521"
         Top             =   1140
         Width           =   345
      End
      Begin VB.TextBox txtVoiceEditorPath 
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   1140
         Width           =   5205
      End
      Begin ComCtl2.UpDown UpDMaxUnDoTimes 
         Height          =   315
         Left            =   2625
         TabIndex        =   15
         Top             =   2790
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtMaxUnDoTimes"
         BuddyDispid     =   196644
         OrigLeft        =   2880
         OrigTop         =   2040
         OrigRight       =   3120
         OrigBottom      =   2325
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMaxUnDoTimes 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2790
         Width           =   915
      End
      Begin VB.ComboBox cmbLangID 
         Height          =   300
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2325
         Width           =   3915
      End
      Begin VB.TextBox txtVoxPath 
         Height          =   315
         Left            =   30
         TabIndex        =   6
         Top             =   450
         Width           =   5205
      End
      Begin VB.ComboBox cmbList 
         Height          =   300
         ItemData        =   "frmOptions.frx":0316
         Left            =   30
         List            =   "frmOptions.frx":0318
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1860
         Width           =   5595
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   345
         Index           =   0
         Left            =   5280
         Picture         =   "frmOptions.frx":031A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "1521"
         Top             =   450
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "语音文件类型"
         Height          =   180
         Index           =   6
         Left            =   0
         TabIndex        =   60
         Tag             =   "1801"
         Top             =   3300
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "录音文件类型"
         Height          =   180
         Index           =   5
         Left            =   3120
         TabIndex        =   59
         Tag             =   "1800"
         Top             =   3300
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "语音编辑外接程序路径"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   58
         Tag             =   "1704"
         Top             =   900
         Width           =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "粘贴板可撤销次数"
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   35
         Tag             =   "1661"
         Top             =   2850
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "选择服务语言编号"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   34
         Tag             =   "1429"
         Top             =   2370
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "选择资源项目"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Tag             =   "1428"
         Top             =   1620
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "系统语音根目录"
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Tag             =   "1427"
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   18
      Top             =   90
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "常规"
            Key             =   "General"
            Object.Tag             =   "1421"
            Object.ToolTipText     =   "Set Options for General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "工具条"
            Key             =   "ToolBar"
            Object.Tag             =   "1422"
            Object.ToolTipText     =   "Set Options for ToolBar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "工作页面"
            Key             =   "Work Page"
            Object.Tag             =   "1424"
            Object.ToolTipText     =   "Set Options for Work Page"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数据源"
            Key             =   "Data Source"
            Object.Tag             =   "1426"
            Object.ToolTipText     =   "Set Database Connetion Parameters"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private m_Color As CHOOSECOLOR
Private i As Integer

'Michael Added @ 06-29-07
'Private Sub cboRecFileType_Change()
Private Sub cborecfiletype_click()
    cmdApply.Enabled = True
End Sub

'Michael Added @ 06-29-07
'Private Sub cboVoiceFileType_Change()
Private Sub cbovoicefiletype_click()
    cmdApply.Enabled = True
End Sub

Private Sub chkShowSysNodes_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmbLangID_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmbList_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    ApplyOptions
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim lv_tempFile As String
    
    Select Case Index
    Case 0          '' 查看目录
        frmSearchDir.Show vbModal
    Case 1          '' 查看全路径
        lv_tempFile = Trim(F_OpenFileDialog(Me.hWnd, False, LoadNationalResString(1705), _
                           LoadNationalResString(1706), "EXE"))
        If lv_tempFile <> "" Then
            txtVoiceEditorPath = lv_tempFile
        End If
        
    End Select
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'' Michael Added @ 2007-12-10 for setting default page size
Private Sub cmdDefArea_Click()
    txtWidth.Text = Format(Str(gSystem.intDefPageWidth / Screen.TwipsPerPixelX))
    txtHeight.Text = Format(Str(gSystem.intDefPageHeight / Screen.TwipsPerPixelY))
End Sub

Private Sub cmdOK_Click()
    'Michael Added for create vox dir @ Aug,7,07
    If Not IsNull(txtVoxPath.Text) Then
        If MakeVoxDir = False Then
            txtVoxPath.SetFocus
            Exit Sub
        End If
    End If
    '*** to here ***
    ApplyOptions
    Unload Me
End Sub

'Michael added @ 2007-11-28
Private Sub cmdTTSSetting_Click()
    frmTTSSetting.Show vbModal
End Sub

Private Sub custcolor_Click()
    CustomColor gFrameBackColor, wrkbkcolor
End Sub

Private Sub CustNodeHand_Click()
    CustomColor gNodeHandColor, NodeHandColor
    gNodeHandColor = NodeHandColor.BackColor
End Sub

Private Sub custwpbkcolor_Click()
    CustomColor gPageBackColor, wpbkcolor
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    txtWidth.Text = Format(Str(gPageWidth / Screen.TwipsPerPixelX))
    txtHeight.Text = Format(Str(gPageHeight / Screen.TwipsPerPixelY))
    wrkbkcolor.BackColor = gFrameBackColor
    wpbkcolor.BackColor = gPageBackColor
    NodeHandColor.BackColor = gNodeHandColor
    
    '' Sun added
    txtPageCount = gCallFlow.PageCount
    txtVoxPath = gSystem.strPath_SysVox
    
    txtMaxUnDoTimes = gClipBoard.MaxStacks
    
    If gSystem.intConfigSet = 1 Then
        tbsOptions.Tabs("Data Source").Selected = True
    Else
        Call AddItemsToList  'Michael Modify this sub @ 06-29-07
        cmbList.ListIndex = mdlcommon.SearchItemDataIndex(cmbList, CLng(gCallFlow.ResourceID), 0)
    End If
        
    '' Sun added 2003-04-23
    Call InitLanguageList
    cmbLangID.ListIndex = mdlcommon.SearchItemDataIndex(cmbLangID, CLng(gCallFlow.LanguageID), 0)
    
    'Set form items -- ODBC
    optDataSource(0).value = Not gSystem.blnUserODBC
    optDataSource(1).value = gSystem.blnUserODBC
    txtServer = gSystem.strDBServer
    txtServer.Enabled = Not gSystem.blnUserODBC
    txtDBName = gSystem.strDBName
    txtDBName.Enabled = Not gSystem.blnUserODBC
    txtODBCDBName = gSystem.strDBName
    txtODBCDBName.Enabled = gSystem.blnUserODBC
    txtDSN = gSystem.strDSN
    txtDSN.Enabled = gSystem.blnUserODBC
    txtUserID = gSystem.strUserID
    txtUserID.Enabled = gSystem.blnUserODBC
    txtPassword = gSystem.strPWD
    txtPassword.Enabled = gSystem.blnUserODBC
    
    '' Sun added 2007-03-25
    txtVoiceEditorPath = gSystem.strVoiceEditorPath
    chkShowSysNodes.value = gSystem.intShowSysNodes
    
    Me.Height = 5340
    Me.Width = 6225
    
    cmdApply.Enabled = False
    
    '' Michael Added @ 2007-12-10
    If (CLng(txtHeight) * Screen.TwipsPerPixelY) <> gSystem.intDefPageHeight Or _
       (CLng(txtWidth) * Screen.TwipsPerPixelX) <> gSystem.intDefPageWidth Then
        cmdDefArea.Enabled = True
    Else
        cmdDefArea.Enabled = False
    End If
    '*********************************************************
    
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True
    LoadResStrings Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdApply.Enabled Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub optDataSource_Click(Index As Integer)
    cmdApply.Enabled = True
    
    If optDataSource(0).value Then
        txtServer.Enabled = True
        txtDBName.Enabled = True
        txtUserID.Enabled = False
        txtPassword.Enabled = False
        txtDSN.Enabled = False
        txtODBCDBName.Enabled = False
    End If
    If optDataSource(1).value Then
        txtServer.Enabled = False
        txtDBName.Enabled = False
        txtUserID.Enabled = True
        txtPassword.Enabled = True
        txtDSN.Enabled = True
        txtODBCDBName.Enabled = True
    End If

End Sub

Private Sub tbsOptions_Click()
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Top = 500
            picOptions(i).Visible = True
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
End Sub

Private Sub ApplyOptions()
On Error Resume Next

    Dim lv_nIndex As Integer
    Dim lv_strNewConnString As String

    If Not cmdApply.Enabled Then Exit Sub
    
    ' notify page width
    If Val(txtWidth.Text) * Screen.TwipsPerPixelX < 0 Or Val(txtWidth.Text) * Screen.TwipsPerPixelX > Screen.Width * 5 Then
        MsgBox "Work page width should be between 0 and 5 * screen width.", vbExclamation
        Exit Sub
    End If
    
    ' notify page height
    If Val(txtHeight.Text) * Screen.TwipsPerPixelY < 0 Or Val(txtHeight.Text) * Screen.TwipsPerPixelY > Screen.Height * 5 Then
        MsgBox "Work page height should be between 0 and 5 * screen width.", vbExclamation
        Exit Sub
    End If
      
    With frmMain
        gFrameBackColor = wrkbkcolor.BackColor
        .SetWorkFramesBackColor
                
        gCallFlow.PageCount = Val(txtPageCount)
        gPageWidth = Val(txtWidth.Text) * Screen.TwipsPerPixelX
        gPageHeight = Val(txtHeight.Text) * Screen.TwipsPerPixelY
        .SetWorkPagesSize
        
        gPageBackColor = wpbkcolor.BackColor
        .SetWorkPagesBackColor
        
        .SetDragHanldeColor
        
        gSystem.strPath_SysVox = Trim(txtVoxPath)
        If cmbList.ListIndex >= 0 And gCallFlow.CallFlowID > 0 Then
            lv_nIndex = cmbList.ItemData(cmbList.ListIndex)
            If gCallFlow.ResourceID <> lv_nIndex Then
                Call gCallFlow.OpenResourceRecordSet(lv_nIndex, gCallFlow.LanguageID)
                F_SetResourceID lv_nIndex
                gCallFlow.SavedMark = False
                .StatusBar.Panels("Resource").Text = LoadNationalResString(1070) & Trim(Str(gCallFlow.ResourceID))
            End If
        End If
        
        If cmbLangID.ListIndex >= 0 Then
            lv_nIndex = cmbLangID.ItemData(cmbLangID.ListIndex)
            If gCallFlow.LanguageID <> lv_nIndex Then
                Call gCallFlow.OpenResourceRecordSet(gCallFlow.ResourceID, lv_nIndex)
                .StatusBar.Panels("Language").Text = LoadNationalResString(1071) & Trim(Str(gCallFlow.LanguageID))
            End If
        End If
     
        ' Write to INI
        Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_FRAME_BG, Str(gFrameBackColor), gSystem.strINI_File)
        Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_HE, Str(gPageHeight), gSystem.strINI_File)
        Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_WD, Str(gPageWidth), gSystem.strINI_File)
        Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_BG, Str(gPageBackColor), gSystem.strINI_File)
        
        '' Sun added 2007-03-25
        '' 语音编辑外接程序路径
        If gSystem.strVoiceEditorPath <> txtVoiceEditorPath Then
            gSystem.strVoiceEditorPath = txtVoiceEditorPath
            Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_VoiceEditorPath, gSystem.strVoiceEditorPath, gSystem.strINI_File)
        End If
        
        '' Sun added 2007-03-25
        '' 是否显示系统节点
        If gSystem.intShowSysNodes <> chkShowSysNodes.value Then
            gSystem.intShowSysNodes = chkShowSysNodes.value
            Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_ShowSysNodes, Str(gSystem.intShowSysNodes), gSystem.strINI_File)
            
            ' 切换所有流程窗口的系统节点显示状态
            frmMain.ShowAllCallFlowWindowSystemNodes
            
        End If
        
        Call WriteIniFileString(Def_INI_SEC_SYS, Def_INI_ENTRY_VoxPath, gSystem.strPath_SysVox, gSystem.strINI_File)
        
        .StatusBar.Panels("Page").Text = LoadNationalResString(1554) & Trim(Str(gCallFlow.CurrentPage)) & LoadNationalResString(1132) & Trim(Str(gCallFlow.PageCount)) & LoadNationalResString(1133)
        
    End With
    
    gSystem.blnUserODBC = optDataSource(1).value
    gSystem.strDSN = Trim(txtDSN)
    gSystem.strUserID = Trim(txtUserID)
    gSystem.strPWD = Trim(txtPassword)
    If gSystem.blnUserODBC Then
        gSystem.strDBName = Trim(txtODBCDBName)
    Else
        gSystem.strDBName = Trim(txtDBName)
    End If
    gSystem.strDBServer = Trim(txtServer)
    
    If gSystem.blnUserODBC Then
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_TYPE, Def_Default_TYPE1, gSystem.strINI_File)
    Else
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_TYPE, Def_Default_TYPE0, gSystem.strINI_File)
    End If

    Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DBNAME, gSystem.strDBName, gSystem.strINI_File)
    Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DBSERVER, gSystem.strDBServer, gSystem.strINI_File)
    Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DSN, gSystem.strDSN, gSystem.strINI_File)
    Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_USERID, gSystem.strUserID, gSystem.strINI_File)
    Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_PWD, gSystem.strPWD, gSystem.strINI_File)
    
    If Not gSystem.blnUserODBC Then
        lv_strNewConnString = "Provider=SQLOLEDB;Integrated Security=SSPI;Data Source=" & gSystem.strDBServer & ";Initial Catalog=" & gSystem.strDBName
    Else
        lv_strNewConnString = "DSN=" & gSystem.strDSN & ";UID=" & gSystem.strUserID & ";PWD=" & gSystem.strPWD & ";Initial Catalog=" & gSystem.strDBName
    End If
    
    '' Sun added 2006-01-27
    If lv_strNewConnString <> gSystem.strConString Then
        gSystem.strConString = lv_strNewConnString
        '' Close All Flows
        frmMain.CloseAllForms False
    End If
    
    If gClipBoard.MaxStacks <> CByte(Val(txtMaxUnDoTimes) Mod 256) Then
        gClipBoard.MaxStacks = CByte(Val(txtMaxUnDoTimes) Mod 256)
        Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_MAXCLIPSTACKS, Str(gClipBoard.MaxStacks), gSystem.strINI_File)
    End If
    
    'Michael Added @ 06-29-07
    If cboRecFileType.ListIndex <> gSystem.intRecFileType Then
        gSystem.intRecFileType = cboRecFileType.ListIndex
        WriteIniFileString Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_RecordFileType, Str(cboRecFileType.ListIndex), gSystem.strINI_File
    End If
    
    If cboVoiceFileType.ListIndex <> gSystem.intVoiFileType Then
        gSystem.intVoiFileType = cboVoiceFileType.ListIndex
        WriteIniFileString Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_VoiceFileType, Str(cboVoiceFileType.ListIndex), gSystem.strINI_File
    End If
    ' Add End
    
    'Michael Added @ 2007-11-28
'    If gCallFlow.PageCount <> 0 And cmdApply.Enabled = True Then
'        Call CFlowWorks.CheckStatus
'    End If
    
    'Mike added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Option changed successful! ")
    cmdApply.Enabled = False
    
End Sub

Private Sub CustomColor(nImportColor As OLE_COLOR, oObject As Object)
    With m_Color
        .lStructSize = Len(m_Color)
        .flags = CC_ANYCOLOR + CC_FULLOPEN + CC_RGBINIT
        .hInstance = App.hInstance
        .hwndOwner = oObject.hWnd
        .rgbResult = F_DealColor(nImportColor)
        .lpCustColors = 0
    End With
    If CHOOSECOLOR(m_Color) Then
        oObject.BackColor = m_Color.rgbResult
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtHeight_Change()
    cmdApply.Enabled = True
    '' Michael Added @ 2007-12-10
    If (CLng(txtHeight) * Screen.TwipsPerPixelY) <> gSystem.intDefPageHeight Then
        cmdDefArea.Enabled = True
    Else
        cmdDefArea.Enabled = False
    End If
End Sub

Private Sub txtMaxUnDoTimes_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtMaxUnDoTimes_GotFocus()
    txtMaxUnDoTimes.SelStart = 0
    txtMaxUnDoTimes.SelLength = Len(txtMaxUnDoTimes)
End Sub

Private Sub txtMaxUnDoTimes_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtODBCDBName_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtODBCDBName_GotFocus()
    txtODBCDBName.SelStart = 0
    txtODBCDBName.SelLength = Len(txtODBCDBName)
End Sub

Private Sub txtPageCount_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPageCount_GotFocus()
    txtPageCount.SelStart = 0
    txtPageCount.SelLength = Len(txtPageCount)
End Sub

Private Sub txtPageCount_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtVoiceEditorPath_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtVoxPath_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtVoxPath_GotFocus()
    txtVoxPath.SelStart = 0
    txtVoxPath.SelLength = Len(txtVoxPath)
End Sub

Private Sub txtWidth_Change()
    cmdApply.Enabled = True
    '' Michael Added @ 2007-12-10
    If (CLng(txtWidth) * Screen.TwipsPerPixelX) <> gSystem.intDefPageWidth Then
        cmdDefArea.Enabled = True
    Else
        cmdDefArea.Enabled = False
    End If
End Sub

Private Sub InitLanguageList()
    Dim lv_loop As Integer

    cmbLangID.Clear
    For lv_loop = 0 To MAX_LANGUAGE_TYPE - 1
        cmbLangID.AddItem LoadNationalResString(1445) & Str(lv_loop)
        cmbLangID.ItemData(cmbLangID.ListCount - 1) = lv_loop
    Next
    
End Sub

Private Sub AddItemsToList()
On Error Resume Next

    Dim lv_RS As ADODB.Recordset
    Dim lv_Str As String
    Dim lv_SQL As String
    Dim lv_loop As Integer
    
    Set lv_RS = New ADODB.Recordset
    
    
    lv_Str = gSystem.strConString

    lv_SQL = "select * from tbRefIVR where P_Type='R' order by P_ID"
    lv_RS.CursorLocation = adUseClient
    
    lv_RS.Open lv_SQL, lv_Str, adOpenStatic, adLockReadOnly, adCmdText
    
    cmbList.Clear
    For lv_loop = 1 To lv_RS.RecordCount
        cmbList.AddItem Str(lv_RS("P_ID")) & "   " & Trim(lv_RS("P_Name")) & "   " & Trim(lv_RS("P_Version")) & "   " & Trim(lv_RS("P_Description"))
        cmbList.ItemData(cmbList.ListCount - 1) = lv_RS("P_ID")
        lv_RS.MoveNext
    Next
    lv_RS.Close
    
    ''************  Michael added 06-29-07  ***************************
    '    "[1] - (*.vox)", 0
    '    "[2] - (*.wav)", 1
    cboRecFileType.Clear
    cboRecFileType.AddItem LoadNationalResString(1802), 0
    cboRecFileType.AddItem LoadNationalResString(1803), 1

    cboVoiceFileType.Clear
    cboVoiceFileType.AddItem LoadNationalResString(1802), 0
    cboVoiceFileType.AddItem LoadNationalResString(1803), 1

    cboRecFileType.ListIndex = gSystem.intRecFileType
    cboVoiceFileType.ListIndex = gSystem.intVoiFileType
    ''***************************  Add End  ************************

End Sub

Private Sub txtDBName_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtDBName_GotFocus()
    txtDBName.SelStart = 0
    txtDBName.SelLength = Len(txtDBName)
End Sub

Private Sub txtDSN_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtDSN_GotFocus()
    txtDSN.SelStart = 0
    txtDSN.SelLength = Len(txtDSN)
End Sub

Private Sub txtPassWord_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtServer_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtServer_GotFocus()
    txtServer.SelStart = 0
    txtServer.SelLength = Len(txtServer)
End Sub

Private Sub txtUserID_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtUserID_GotFocus()
    txtUserID.SelStart = 0
    txtUserID.SelLength = Len(txtUserID)
End Sub

' 2005-03-10
' 页面数量验证
'
Private Sub UpdPageCount_DownClick()
Dim lv_nPageCount As Integer

    lv_nPageCount = gCallFlow.GetActualPageCount
    If UpdPageCount.value < lv_nPageCount Then
        Message "M141"
        UpdPageCount.value = UpdPageCount.value + UpdPageCount.Increment
    ElseIf gCallFlow.CurrentPage > lv_nPageCount Then
        GotoAnotherPage lv_nPageCount
    End If
End Sub

' Aug,07,07
' Create Vox Dirctory
' Michael
Private Function MakeVoxDir() As Boolean
On Error Resume Next

    Dim path As String
    'Dim fso As Object
    path = Me.txtVoxPath.Text
    'Set fso = CreateObject("Scripting.FileSystemObject")
    
    MakeVoxDir = True
    
    'Dir is not exist
    'If fso.folderexists(path) = False Then
    If Dir(path, vbDirectory) = "" Then
        'Michael Modified @ 2007-11-28  -> Q Message
        '文件夹不存在,需要创建吗?"
            If Message("Q021") = vbYes Then
            'Driver is not vaild
            'If fso.folderexists(Left(path, 3)) = False Then
            If Dir(Left(path, 3), vbDirectory) = "" Then
                ' Michael Modified @ 2007-11-28 -> W Message
                ' 驱动器不存在或路径错误,请键入绝对路径!
                Call Message("E137")
                Me.txtVoxPath.Text = ""
                MakeVoxDir = False
                'Mike added @ 2008-7-8
                Call WriteLogMessage(0, enu_Warnning, "Failed to create resource folder! ", "Path=" & path)
                Exit Function
            Else
                'fso.CreateFolder (path)
                
                '' Sun updated 2008-02018
                ''' From
                ''MkDir (path)
                ''' To
                If CreateNewDir(path) > 0 Then
                    Call Message("E137")
                    MakeVoxDir = False
                    'Mike added @ 2008-7-8
                    Call WriteLogMessage(0, enu_Warnning, "Failed to create resource folder! ", "Path=" & path)
                    Exit Function
                Else
                    'Mike added @ 2008-7-8
                    Call WriteLogMessage(0, enu_Information, "Create Resource Folder Successful! ", "Path=" & path)
                End If
                
            End If
        End If
    End If
    'Set fso = Nothing
    
End Function









