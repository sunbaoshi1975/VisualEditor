VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmFlowProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "流程属性"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmFlowProperty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1700"
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   120
      Picture         =   "frmFlowProperty.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   50
      Top             =   4410
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   480
      Picture         =   "frmFlowProperty.frx":044E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   49
      Top             =   4380
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   2
      Left            =   1080
      Picture         =   "frmFlowProperty.frx":0D18
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   48
      Top             =   4380
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   2
      Left            =   9720
      ScaleHeight     =   4740
      ScaleWidth      =   9285
      TabIndex        =   46
      Top             =   780
      Visible         =   0   'False
      Width           =   9285
      Begin FPSpreadADO.fpSpread vasVar 
         Height          =   4605
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   9240
         _Version        =   458752
         _ExtentX        =   16298
         _ExtentY        =   8123
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   255
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmFlowProperty.frx":0E62
         UserResize      =   1
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   0
      Left            =   120
      ScaleHeight     =   4740
      ScaleWidth      =   9225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   9225
      Begin VB.TextBox txtFlowID 
         BackColor       =   &H8000000F&
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
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   120
         Width           =   4395
      End
      Begin VB.TextBox Txt_pmodifytime 
         Height          =   375
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox Txt_pcreatetime 
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox Txt_puser 
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
         Left            =   6570
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1740
         Width           =   2595
      End
      Begin VB.TextBox Txt_pauther 
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
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1740
         Width           =   2595
      End
      Begin VB.TextBox Txt_pversion 
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
         Left            =   7830
         MaxLength       =   10
         TabIndex        =   4
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox Txt_pdescription 
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
         Left            =   1650
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1200
         Width           =   7515
      End
      Begin VB.TextBox Txt_pname 
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
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   3
         Top             =   660
         Width           =   4395
      End
      Begin VB.Label p_modi_Label 
         AutoSize        =   -1  'True
         Caption         =   "修改时间:"
         Height          =   195
         Left            =   5010
         TabIndex        =   37
         Tag             =   "1370"
         Top             =   2340
         Width           =   1410
      End
      Begin VB.Label p_crea_Label 
         AutoSize        =   -1  'True
         Caption         =   "创建时间:"
         Height          =   195
         Left            =   90
         TabIndex        =   36
         Tag             =   "1369"
         Top             =   2340
         Width           =   1410
      End
      Begin VB.Label p_user_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程用户:"
         Height          =   195
         Left            =   5010
         TabIndex        =   35
         Tag             =   "1368"
         Top             =   1830
         Width           =   1410
      End
      Begin VB.Label p_auth_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程作者:"
         Height          =   195
         Left            =   90
         TabIndex        =   34
         Tag             =   "1367"
         Top             =   1830
         Width           =   1410
      End
      Begin VB.Label p_ver_Label 
         AutoSize        =   -1  'True
         Caption         =   "版本号:"
         Height          =   195
         Left            =   6510
         TabIndex        =   33
         Tag             =   "1366"
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label p_desc_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程描述:"
         Height          =   195
         Left            =   90
         TabIndex        =   32
         Tag             =   "1365"
         Top             =   1290
         Width           =   1410
      End
      Begin VB.Label p_name_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程名称:"
         Height          =   195
         Left            =   90
         TabIndex        =   31
         Tag             =   "1364"
         Top             =   750
         Width           =   1410
      End
      Begin VB.Label p_id_lable 
         AutoSize        =   -1  'True
         Caption         =   "流程ID:"
         Height          =   195
         Left            =   90
         TabIndex        =   30
         Tag             =   "1363"
         Top             =   210
         Width           =   1410
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   1
      Left            =   120
      ScaleHeight     =   4740
      ScaleWidth      =   9225
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5640
      Width           =   9225
      Begin VB.Frame famSettings 
         Caption         =   "设置"
         Height          =   1935
         Left            =   0
         TabIndex        =   15
         Tag             =   "1136"
         Top             =   1740
         Width           =   9165
         Begin VB.TextBox txtSysHookNode 
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
            Left            =   6630
            MaxLength       =   6
            TabIndex        =   24
            Text            =   "256"
            Top             =   990
            Width           =   1065
         End
         Begin VB.CommandButton cmdShowNodeList 
            Height          =   285
            Index           =   1
            Left            =   7710
            Picture         =   "frmFlowProperty.frx":1318
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "1145"
            Top             =   1020
            Width           =   315
         End
         Begin VB.CheckBox chkLogSwitchOff 
            Caption         =   "不记录节点日志"
            Height          =   360
            Left            =   150
            TabIndex        =   19
            Tag             =   "1951"
            Top             =   1350
            Width           =   1695
         End
         Begin VB.TextBox T_nd_root 
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
            Left            =   6630
            MaxLength       =   6
            TabIndex        =   20
            Text            =   "256"
            Top             =   270
            Width           =   1065
         End
         Begin VB.TextBox txtMainCOM 
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
            Left            =   6630
            MaxLength       =   6
            TabIndex        =   26
            Text            =   "256"
            Top             =   1350
            Width           =   1065
         End
         Begin VB.CommandButton cmdShowRes 
            Height          =   285
            Left            =   7710
            Picture         =   "frmFlowProperty.frx":16A2
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "1146"
            Top             =   1380
            Width           =   315
         End
         Begin VB.CommandButton cmdShowNodeList 
            Height          =   285
            Index           =   0
            Left            =   7710
            Picture         =   "frmFlowProperty.frx":17A4
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "1145"
            Top             =   300
            Width           =   315
         End
         Begin VB.CommandButton cmdShowTypedNodeList 
            Height          =   285
            Left            =   7710
            Picture         =   "frmFlowProperty.frx":1B2E
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "1145"
            Top             =   660
            Width           =   315
         End
         Begin VB.TextBox txtSysDataFormatNode 
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
            Left            =   6630
            MaxLength       =   6
            TabIndex        =   22
            Text            =   "256"
            Top             =   630
            Width           =   1065
         End
         Begin VB.ComboBox CB_key_repeat 
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
            ItemData        =   "frmFlowProperty.frx":1EB8
            Left            =   2400
            List            =   "frmFlowProperty.frx":1EBA
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   270
            Width           =   1335
         End
         Begin VB.ComboBox CB_key_return 
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
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   630
            Width           =   1335
         End
         Begin VB.ComboBox Cb_key_root 
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
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "挂机转节点ID"
            Height          =   195
            Index           =   1
            Left            =   4620
            TabIndex        =   52
            Tag             =   "2137"
            Top             =   1050
            Width           =   1065
         End
         Begin VB.Label Lbl_nd_root 
            AutoSize        =   -1  'True
            Caption         =   "根节点ID"
            Height          =   255
            Left            =   4620
            TabIndex        =   45
            Tag             =   "1141"
            Top             =   330
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "COM接口资源ID"
            Height          =   195
            Left            =   4620
            TabIndex        =   44
            Tag             =   "1142"
            Top             =   1410
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "缺省数据发送节点ID"
            Height          =   195
            Index           =   0
            Left            =   4620
            TabIndex        =   43
            Tag             =   "1538"
            Top             =   690
            Width           =   1605
         End
         Begin VB.Label Lbl_key_root 
            AutoSize        =   -1  'True
            Caption         =   "回到主菜单"
            Height          =   255
            Left            =   150
            TabIndex        =   42
            Tag             =   "1140"
            Top             =   1050
            Width           =   900
         End
         Begin VB.Label Lbl_key_return 
            AutoSize        =   -1  'True
            Caption         =   "回上一级节点按键"
            Height          =   255
            Left            =   150
            TabIndex        =   41
            Tag             =   "1139"
            Top             =   690
            Width           =   1440
         End
         Begin VB.Label Lbl_key_repeat 
            AutoSize        =   -1  'True
            Caption         =   "重复当前节点按键"
            Height          =   255
            Left            =   150
            TabIndex        =   40
            Tag             =   "1138"
            Top             =   330
            Width           =   1440
         End
      End
      Begin VB.Frame famRes 
         Caption         =   "资源"
         Height          =   1665
         Left            =   0
         TabIndex        =   11
         Tag             =   "1703"
         Top             =   0
         Width           =   9165
         Begin VB.ComboBox cmbLANCount 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   945
            Width           =   1725
         End
         Begin VB.ComboBox cmbList 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   510
            Width           =   7905
         End
         Begin VB.ComboBox cmbLangID 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   945
            Width           =   2085
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "语言数量"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   51
            Tag             =   "1743"
            Top             =   990
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "选择资源项目"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   39
            Tag             =   "1428"
            Top             =   270
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "选择服务语言编号"
            Height          =   195
            Index           =   2
            Left            =   4230
            TabIndex        =   38
            Tag             =   "1429"
            Top             =   990
            Width           =   1440
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   6900
      TabIndex        =   28
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   8130
      TabIndex        =   29
      Top             =   5415
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   5205
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   9181
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "描述"
            Key             =   "keyCaption"
            Object.Tag             =   "1701"
            Object.ToolTipText     =   "Flow Descriptions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "流程参数"
            Key             =   "keyParameter"
            Object.Tag             =   "1702"
            Object.ToolTipText     =   "Flow Parameters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "变量使用"
            Key             =   "keyVarList"
            Object.Tag             =   "1457"
            Object.ToolTipText     =   "Variables List"
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
Attribute VB_Name = "frmFlowProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' 选项卡是否改变
Dim m_blnTabChanged() As Boolean
Dim lv_nLoop As Integer
''Mike add @ 2008-4-25 ,Update MDI Caption
Dim m_blnUpdateCaption As Boolean

Private Sub cmbLANCount_Click()
    Call InitLanguageList
    cmbLangID.ListIndex = mdlcommon.SearchItemDataIndex(cmbLangID, CLng(gCallFlow.LanguageID), 0)
    m_blnTabChanged(1) = True
End Sub

Private Sub cmbLangID_Click()
    m_blnTabChanged(1) = True
End Sub

Private Sub cmbList_Click()
    m_blnTabChanged(1) = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' Save Settings
'
Private Sub cmdOK_Click()
    Dim lv_nIndex As Integer
    Dim lv_nVarCount As Integer
    Dim lv_strName As String
    Dim lv_bytType As Byte
    Dim lv_bytLen As Byte
    
    '' 更新流程信息
    If m_blnTabChanged(0) Then
        Call gCallFlow.UpdateCallFlowInfo(Trim(Txt_pname), Trim(Txt_pdescription), Trim(Txt_pauther), Trim(Txt_puser), Trim(Txt_pversion))
        'Change MDI title when the Name of the CallFlow changed,added by Mike @ 2008-4-25
        If m_blnUpdateCaption Then Call frmMain.UpdateSystemStatusBar
        m_blnUpdateCaption = False
        
        m_blnTabChanged(0) = False
    End If
    
    '' 更新属性
    If m_blnTabChanged(1) Then
        
        If cmbList.ListIndex >= 0 And gCallFlow.CallFlowID > 0 Then
            lv_nIndex = cmbList.ItemData(cmbList.ListIndex)
            If gCallFlow.ResourceID <> lv_nIndex Then
                Call gCallFlow.OpenResourceRecordSet(lv_nIndex, gCallFlow.LanguageID)
                Node0_Data2.ResourceProject = lv_nIndex
                frmMain.StatusBar.Panels("Resource").Text = LoadNationalResString(1070) & Trim(Str(Node0_Data2.ResourceProject))
            End If
        End If
        
        If cmbLANCount.ListIndex >= 0 Then
            Node0_Data1.Languages = cmbLANCount.ItemData(cmbLANCount.ListIndex)
        End If
        
        If cmbLangID.ListIndex >= 0 Then
            lv_nIndex = cmbLangID.ItemData(cmbLangID.ListIndex)
            If gCallFlow.LanguageID <> lv_nIndex Then
                Call gCallFlow.OpenResourceRecordSet(Node0_Data2.ResourceProject, lv_nIndex)
                frmMain.StatusBar.Panels("Language").Text = LoadNationalResString(1071) & Trim(Str(gCallFlow.LanguageID))
            End If
        End If
        
        ' 重复当前节点按键默认为0
        If CB_key_repeat.ListIndex < 0 Then
           Node0_Data2.key_repeat = 0
        Else
            Node0_Data2.key_repeat = CByte(CB_key_repeat.ItemData(CB_key_repeat.ListIndex))
        End If
        
        ' 回到上一级菜单默认为0
        If CB_key_return.ListIndex < 0 Then
           Node0_Data2.key_return = 0
        Else
           Node0_Data2.key_return = CByte(CB_key_return.ItemData(CB_key_return.ListIndex))
        End If
        
        ' 回到主菜单默认为0
        If Cb_key_root.ListIndex < 0 Then
           Node0_Data2.key_root = 0
        Else
           Node0_Data2.key_root = CByte(Cb_key_root.ItemData(Cb_key_root.ListIndex))
        End If
        
        ' 根节点默认为256
        If Trim(T_nd_root) = "" Then
           Node0_Data2.nd_root = 256
        Else
           If Val(T_nd_root) > 32767 Or Val(T_nd_root) < 256 Then
              Message ("E034")
              T_nd_root.SetFocus
              Exit Sub
           Else
              Node0_Data2.nd_root = Val(T_nd_root.Text)
           End If
        End If
        
        ' Sun added 2004-12-30
        '' 缺省数据发送节点ID
        If Trim(txtSysDataFormatNode) = "" Then
            Node0_Data2.nd_SysSendData = 0
        Else
            Node0_Data2.nd_SysSendData = txtSysDataFormatNode
        End If
        
        ' Sun added 2012-05-07
        '' 挂机前转节点
        If Trim(txtSysHookNode) = "" Then
            Node0_Data2.nd_BeforeHookOn = 0
        Else
            Node0_Data2.nd_BeforeHookOn = txtSysHookNode
        End If
        
        ' COM接口资源ID
        If Trim(txtMainCOM) = "" Then
            Node0_Data2.MainCOM = 0
        Else
            If CLng(Trim(txtMainCOM)) > 32767 Then
                Message ("E088")
                txtMainCOM.SetFocus
                Exit Sub
            Else
                Node0_Data2.MainCOM = CInt(Trim(txtMainCOM))
            End If
        End If
        Node0_Data2.ResourceProject = gCallFlow.ResourceID
        
        '节点日志全局控制开关 <Mike 2008-2-19>
        Node0_Data2.LogSwitchOff = CByte(chkLogSwitchOff.value)
        
        '' Sun added 2002-12-04
        If Node0_Data2.nd_root <> gCallFlow.RootNodedID Then
            F_SwitchRootNodeDisplay Node0_Data2.nd_root
        End If
        
        '' Sun added 2007-03-25
        Node0_Data1.reserved1(0) = 0
        Node0_Data1.MajorVer = Def_CallFlow_MajorVersion
        Node0_Data1.MinorVer = Def_CallFlow_MinorVersion
        
        F_NodeData 1, 0
        gCallFlow.UpdateAnotherIVRRecord 1
        
        m_blnTabChanged(1) = False
        
    End If
    
    '' 更新变量
    If m_blnTabChanged(2) Then
    
        '' 变量总数
        lv_nVarCount = 0
        For lv_nIndex = vasVar.MaxRows To 1 Step -1
            vasVar.Row = lv_nIndex
            vasVar.Col = 1
            If Trim(vasVar.Text) <> "" Then
                vasVar.Col = 3
                If Val(vasVar.Text) > 0 Then
                    lv_nVarCount = lv_nIndex
                    Exit For
                End If
            End If
        Next
        If Node1_Data1.uservars <> lv_nVarCount Then
            'Mike added 2008-8-20
            Dim lv_olduservars As Byte
            lv_olduservars = Node1_Data1.uservars
            
            '记录
            Node1_Data1.uservars = lv_nVarCount
            
            '节点数据整和
            F_NodeData 2, 1
            
            '保存节点
            gCallFlow.UpdateAnotherIVRRecord 2
            
            F_CreateVar lv_nVarCount
        
        End If
        
        '' 各个变量节点
        For lv_nIndex = 1 To lv_nVarCount
            vasVar.Row = lv_nIndex
            vasVar.Col = 1
            lv_strName = Trim(vasVar.Text)
            vasVar.Col = 2
            If vasVar.TypeComboBoxCurSel >= 0 Then
                lv_bytType = vasVar.TypeComboBoxCurSel + 1
            Else
                lv_bytType = 2
            End If
            vasVar.Col = 3
            lv_bytLen = CByte(Val(vasVar.Text) Mod 256)
            
            If lv_strName <> "" And lv_bytLen > 0 Then
                gCallFlow.SetUserVarDefination lv_nIndex, lv_strName, lv_bytType, lv_bytLen
            End If
        Next
        
        'Mike added 2008-8-20 ;empty null var structor
        Dim lv_temp As Integer
        Dim lv_temploop As Integer
        Dim lv_start As Integer
        lv_temp = (lv_nVarCount - 1) Mod 4
        
        If (lv_nVarCount < lv_olduservars) And (lv_temp < 3) Then
            lv_start = 1
            For lv_temploop = lv_temp + 1 To 3
                gCallFlow.SetUserVarDefination lv_nVarCount + lv_start, "0", 2, 0
                lv_start = lv_start + 1
            Next lv_temploop
        End If
        
        m_blnTabChanged(2) = False
    End If
    
    Unload Me

End Sub

Private Sub cmdShowNodeList_Click(Index As Integer)
    
    Select Case Index
    Case 0
        Set gSystem.crlCurItem = T_nd_root
    Case 1
        Set gSystem.crlCurItem = txtSysHookNode
    End Select
    frmNodeList.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        lv_nLoop = tbsOptions.SelectedItem.Index
        If lv_nLoop = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(lv_nLoop + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    SetMainFormItemsEnableWhenPropertyShow False
    
    ReDim m_blnTabChanged(tbsOptions.Tabs.Count) As Boolean
    tbsOptions.Tabs("keyCaption").Selected = True
    
    Me.Height = 6285
    Me.Width = 9750
    
    ' 设置流程属性
    If gCallFlow.CallFlowID > 0 Then
        '' 描述
        txtFlowID = gCallFlow.CallFlowID
        Txt_pname = gCallFlow.CallFlowName                  '流程名称
        Txt_pdescription = Trim(gCallFlow.Description)      '流程描述
        Txt_pauther = Trim(gCallFlow.CallFlowAuther)        '流程作者
        Txt_puser = Trim(gCallFlow.CallFlowUser)            '流程用户
        Txt_pversion = Trim(gCallFlow.CallFlowVersion)      '版本号
        Txt_pmodifytime = Format(gCallFlow.ModifyTime, "yyyy-mm-dd hh:nn:ss")             '修改时间
        Txt_pcreatetime = Format(gCallFlow.CreateTime, "yyyy-mm-dd hh:nn:ss")             '创建时间
        
        '' 参数
        Call AddItemsToList
        cmbList.ListIndex = mdlcommon.SearchItemDataIndex(cmbList, CLng(gCallFlow.ResourceID), 0)
        
        Call InitLanguageCount
        If Node0_Data1.Languages <= 0 Then
            Node0_Data1.Languages = 1
        End If
        cmbLANCount.ListIndex = mdlcommon.SearchItemDataIndex(cmbLANCount, CLng(Node0_Data1.Languages), 0)
        
        Call InitLanguageList
        cmbLangID.ListIndex = mdlcommon.SearchItemDataIndex(cmbLangID, CLng(gCallFlow.LanguageID), 0)
        
        '重复当前按钮
        F_FillPhoneKeyList CB_key_repeat, 2
        CB_key_repeat.ListIndex = SearchItemDataIndex(CB_key_repeat, CLng(Node0_Data2.key_repeat), 0)
        '返回上级菜单
        F_FillPhoneKeyList CB_key_return, 2
        CB_key_return.ListIndex = SearchItemDataIndex(CB_key_return, CLng(Node0_Data2.key_return), 0)
        '返回根节点
        F_FillPhoneKeyList Cb_key_root, 2
        Cb_key_root.ListIndex = SearchItemDataIndex(Cb_key_root, CLng(Node0_Data2.key_root), 0)
        
        '缺省数据发送节点ID
        txtSysDataFormatNode = Node0_Data2.nd_SysSendData
              
        ' Sun added 2012-05-07
        '' 挂机前转节点
        txtSysHookNode = Node0_Data2.nd_BeforeHookOn
              
        'COM接口资源ID sun 2002-12-03
        txtMainCOM = Node0_Data2.MainCOM
        
        '不记录节点日志 <0-记录(默认),1-不记录> ; Mike 2008-2-19
        chkLogSwitchOff.value = Node0_Data2.LogSwitchOff
        
        ' 变量清单
        RefreshVarSpread
    
    End If
    
    '' 选中指定卡片，
    ''' 注意：卡片从1开始编号
    If gSystem.intConfigSet < tbsOptions.Tabs.Count Then
        tbsOptions.Tabs(gSystem.intConfigSet + 1).Selected = True
    End If
    
    For lv_nLoop = 0 To tbsOptions.Tabs.Count - 1
        m_blnTabChanged(lv_nLoop) = False
    Next
    
    'Mike add @ 2008-4-25
    m_blnUpdateCaption = False
 
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True
    LoadResStrings Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lv_blnChanged As Boolean
    
    lv_blnChanged = False
    For lv_nLoop = 0 To tbsOptions.Tabs.Count - 1
        If m_blnTabChanged(lv_nLoop) Then
            lv_blnChanged = True
            Exit For
        End If
    Next
    
    If lv_blnChanged Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

Private Sub tbsOptions_Click()
    'show and enable the selected tab's controls
    'and hide and disable all others
    Dim i As Integer
    
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

Private Sub Txt_pauther_Change()
    m_blnTabChanged(0) = True
End Sub

Private Sub Txt_pauther_GotFocus()
    Txt_pauther.SelStart = 0
    Txt_pauther.SelLength = Len(Txt_pauther)
End Sub

Private Sub Txt_pdescription_Change()
    m_blnTabChanged(0) = True
End Sub

Private Sub Txt_pdescription_GotFocus()
    Txt_pdescription.SelStart = 0
    Txt_pdescription.SelLength = Len(Txt_pdescription)
End Sub

Private Sub Txt_pname_Change()
    m_blnTabChanged(0) = True
    'Mike @ 2008-4-25
    m_blnUpdateCaption = True
End Sub

Private Sub Txt_pname_GotFocus()
    Txt_pname.SelStart = 0
    Txt_pname.SelLength = Len(Txt_pname)
End Sub

Private Sub Txt_puser_Change()
    m_blnTabChanged(0) = True
End Sub

Private Sub Txt_puser_GotFocus()
    Txt_puser.SelStart = 0
    Txt_puser.SelLength = Len(Txt_puser)
End Sub

Private Sub Txt_pversion_Change()
    m_blnTabChanged(0) = True
End Sub

Private Sub Txt_pversion_GotFocus()
    Txt_pversion.SelStart = 0
    Txt_pversion.SelLength = Len(Txt_pversion)
End Sub

Private Sub T_nd_root_Change()
    m_blnTabChanged(1) = True
End Sub

Private Sub T_nd_root_GotFocus()
    T_nd_root.SelStart = 0
    T_nd_root.SelLength = Len(T_nd_root)
End Sub

Private Sub T_nd_root_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtMainCOM_Change()
    m_blnTabChanged(1) = True
    
    ''' Get Resource Description
    F_RefreshVoxBoxToolTip txtMainCOM
    
End Sub

Private Sub txtMainCOM_GotFocus()
    txtMainCOM.SelStart = 0
    txtMainCOM.SelLength = Len(txtMainCOM)
End Sub

Private Sub txtMainCOM_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub txtSysDataFormatNode_Change()
    m_blnTabChanged(1) = True
End Sub

Private Sub txtSysDataFormatNode_GotFocus()
    txtSysDataFormatNode.SelStart = 0
    txtSysDataFormatNode.SelLength = Len(txtSysDataFormatNode)
End Sub

Private Sub txtSysDataFormatNode_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub CB_key_repeat_Click()
    m_blnTabChanged(1) = True
End Sub

Private Sub CB_key_return_Click()
    m_blnTabChanged(1) = True
End Sub

Private Sub Cb_key_root_Click()
    m_blnTabChanged(1) = True
End Sub

Private Sub cmdShowRes_Click()
    
    gSystem.intCurStep = 3
    Set gSystem.crlCurItem = txtMainCOM
    frmResourceList.Show vbModal
   
End Sub

Private Sub cmdShowTypedNodeList_Click()

    '' Sun added 2004-12-30
    Set gSystem.crlCurItem = txtSysDataFormatNode
    g_NodeListShowType = 18
    frmNodeList.Show vbModal
    g_NodeListShowType = 0

End Sub

Private Sub InitLanguageCount()
    Dim lv_Loop As Integer

    cmbLANCount.Clear
    For lv_Loop = 1 To MAX_LANGUAGE_TYPE
        cmbLANCount.AddItem Str(lv_Loop)
        cmbLANCount.ItemData(cmbLANCount.ListCount - 1) = lv_Loop
    Next
    
End Sub

Private Sub InitLanguageList()
    Dim lv_Loop As Integer

    cmbLangID.Clear
    If cmbLANCount.ListIndex >= 0 Then
        For lv_Loop = 0 To cmbLANCount.ItemData(cmbLANCount.ListIndex) - 1
            cmbLangID.AddItem LoadNationalResString(1445) & Str(lv_Loop)
            cmbLangID.ItemData(cmbLangID.ListCount - 1) = lv_Loop
        Next
    End If
End Sub

Private Sub AddItemsToList()
On Error Resume Next

    Dim lv_RS As ADODB.Recordset
    Dim lv_Str As String
    Dim lv_SQL As String
    Dim lv_Loop As Integer
    
    Set lv_RS = New ADODB.Recordset
    
    
    lv_Str = gSystem.strConString

    lv_SQL = "select * from tbRefIVR where P_Type='R' order by P_ID"
    lv_RS.CursorLocation = adUseClient
    
    lv_RS.Open lv_SQL, lv_Str, adOpenStatic, adLockReadOnly, adCmdText
    
    cmbList.Clear
    For lv_Loop = 1 To lv_RS.RecordCount
        cmbList.AddItem Str(lv_RS("P_ID")) & "   " & Trim(lv_RS("P_Name")) & "   " & Trim(lv_RS("P_Version")) & "   " & Trim(lv_RS("P_Description"))
        cmbList.ItemData(cmbList.ListCount - 1) = lv_RS("P_ID")
        lv_RS.MoveNext
    Next
    lv_RS.Close

End Sub

Private Sub RefreshVarSpread()
On Error Resume Next

Dim lv_Loop As Integer
Dim lv_strName As String
Dim lv_bytType As Byte
Dim lv_bytLen As Byte
Dim lv_intWriteTimes As Integer
Dim lv_intNodeCount As Integer
Dim lv_intRelateNode(50) As Integer
Dim lv_intReadTimes As Integer
Dim lv_intReadCount As Integer
Dim lv_intReadNode(50) As Integer
Dim lv_nVarCount As Integer

'' Sun added 2002-06-12
Dim lv_SubLoop As Integer
Dim lv_StrTemp As String

    vasVar.Enabled = False
    
    For lv_Loop = 1 To vasVar.MaxRows
        vasVar.Row = lv_Loop
        vasVar.Col = 2
        vasVar.CellType = CellTypeComboBox
        vasVar.TypeComboBoxList = LoadNationalResString(1467) & Chr(9) & LoadNationalResString(1468)
    Next

    lv_nVarCount = gCallFlow.GetUserVarCount
    For lv_Loop = 1 To lv_nVarCount
        
        vasVar.Row = lv_Loop
        lv_strName = ""
        lv_bytType = 0
        lv_bytLen = 0
        Call gCallFlow.GetUserVarDefination(lv_Loop, lv_strName, lv_bytType, lv_bytLen)
        lv_intWriteTimes = 0
        lv_intNodeCount = 0
        Call gCallFlow.GetUserVarUsedStatus(lv_Loop, lv_intWriteTimes, lv_intNodeCount, lv_intRelateNode(), _
                                    lv_intReadTimes, lv_intReadCount, lv_intReadNode())
        
        vasVar.Col = 1
        vasVar.Text = Trim(lv_strName)
        
        vasVar.Col = 2
        Select Case lv_bytType
        Case 1
            'vasVar.Text = LoadNationalResString(1467)
            vasVar.TypeComboBoxCurSel = 0
        Case 2
            'vasVar.Text = LoadNationalResString(1468)
            vasVar.TypeComboBoxCurSel = 1
        Case Else
            vasVar.TypeComboBoxCurSel = 1
        End Select
        
        vasVar.Col = 3
        vasVar.value = lv_bytLen
        
        vasVar.Col = 4
        If lv_intWriteTimes <= 0 Then
            vasVar.TypePictPicture = picStatus(0).Picture
            vasVar.Col = 5
            vasVar.Text = ""
        ElseIf lv_intWriteTimes = 1 Then
            vasVar.TypePictPicture = picStatus(1).Picture
            vasVar.Col = 5
            vasVar.Text = Trim(gCallFlow.Node(lv_intRelateNode(0)).NodeCaption)
        Else
            vasVar.TypePictPicture = picStatus(2).Picture
            vasVar.Col = 5
            lv_StrTemp = Trim(Str(lv_intWriteTimes)) & LoadNationalResString(1466) & Trim(gCallFlow.Node(lv_intRelateNode(0)).NodeCaption)
            
            '' Sun added 2002-06-12
            For lv_SubLoop = 1 To lv_intNodeCount - 1
                lv_StrTemp = lv_StrTemp & "； " & Trim(gCallFlow.Node(lv_intRelateNode(lv_SubLoop)).NodeCaption)
            Next
            vasVar.Text = lv_StrTemp

        End If
    
        vasVar.Col = 6
        If lv_intReadTimes <= 0 Then
            vasVar.Text = ""
        ElseIf lv_intReadTimes = 1 Then
            vasVar.Text = Trim(gCallFlow.Node(lv_intReadNode(0)).NodeCaption)
        Else
            lv_StrTemp = Trim(Str(lv_intReadTimes)) & LoadNationalResString(1466) & Trim(gCallFlow.Node(lv_intReadNode(0)).NodeCaption)
            
            '' Sun added 2002-06-12
            For lv_SubLoop = 1 To lv_intReadCount - 1
                lv_StrTemp = lv_StrTemp & "； " & Trim(gCallFlow.Node(lv_intReadNode(lv_SubLoop)).NodeCaption)
            Next
            vasVar.Text = lv_StrTemp

        End If
    
    Next
    
    vasVar.Enabled = True
    
On Error GoTo 0
End Sub

Private Sub txtSysHookNode_Change()
    m_blnTabChanged(1) = True
End Sub

Private Sub txtSysHookNode_GotFocus()
    txtSysHookNode.SelStart = 0
    txtSysHookNode.SelLength = Len(txtSysHookNode)
End Sub

Private Sub txtSysHookNode_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub vasVar_Change(ByVal Col As Long, ByVal Row As Long)
    If Col <= 3 Then
        m_blnTabChanged(2) = True
    End If
End Sub

' Mike 2008-2-19
Private Sub chkLogSwitchOff_Click()
    m_blnTabChanged(1) = True
End Sub
