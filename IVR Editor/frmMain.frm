VERSION 5.00
Object = "{5F1D7864-ED0C-4F12-B10C-659366B50BF7}#1.2#0"; "OutLookBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{19C577A9-275A-451D-A97E-CFEF5120930D}#1.0#0"; "VOXPLA~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Visual IVR Workflow Editor"
   ClientHeight    =   6210
   ClientLeft      =   5490
   ClientTop       =   4590
   ClientWidth     =   8910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Tag             =   "1066"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Splitter 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   1455
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5505
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   420
      Width           =   30
   End
   Begin VB.PictureBox picToolBarHolder 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   0
      ScaleHeight     =   5505
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   420
      Width           =   1455
      Begin OutLookBar.VerticalMenu ToolBar 
         Height          =   9975
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   17595
         Enabled         =   -1  'True
         MenusMax        =   5
         MenuCur         =   2
         MenuStartup     =   2
         MenuCaption1    =   "1032"
         MenuItemsMax1   =   8
         MenuItemIcon11  =   "frmMain.frx":0442
         MenuItemCaption11=   "无条件转移"
         MenuItemKey11   =   "6"
         MenuItemTag11   =   "1037"
         MenuItemIcon12  =   "frmMain.frx":075C
         MenuItemCaption12=   "身份验证"
         MenuItemKey12   =   "7"
         MenuItemTag12   =   "1038"
         MenuItemIcon13  =   "frmMain.frx":0A76
         MenuItemCaption13=   "修改口令"
         MenuItemKey13   =   "8"
         MenuItemTag13   =   "1039"
         MenuItemIcon14  =   "frmMain.frx":0D90
         MenuItemCaption14=   "时间分支"
         MenuItemKey14   =   "9"
         MenuItemTag14   =   "1040"
         MenuItemIcon15  =   "frmMain.frx":10AA
         MenuItemCaption15=   "工作日设定"
         MenuItemKey15   =   "10"
         MenuItemTag15   =   "1041"
         MenuItemIcon16  =   "frmMain.frx":13C4
         MenuItemCaption16=   "选择服务语言"
         MenuItemKey16   =   "17"
         MenuItemTag16   =   "1042"
         MenuItemIcon17  =   "frmMain.frx":16DE
         MenuItemCaption17=   "条件分支"
         MenuItemKey17   =   "16"
         MenuItemTag17   =   "1528"
         MenuItemIcon18  =   "frmMain.frx":19F8
         MenuItemCaption18=   "节点连线"
         MenuItemKey18   =   "255"
         MenuItemTag18   =   "1045"
         MenuCaption2    =   "1033"
         MenuItemsMax2   =   8
         MenuItemIcon21  =   "frmMain.frx":1D12
         MenuItemCaption21=   "放音挂机"
         MenuItemKey21   =   "20"
         MenuItemTag21   =   "1046"
         MenuItemIcon22  =   "frmMain.frx":202C
         MenuItemCaption22=   "放音继续"
         MenuItemKey22   =   "21"
         MenuItemTag22   =   "1047"
         MenuItemIcon23  =   "frmMain.frx":2346
         MenuItemCaption23=   "放音等待按键"
         MenuItemKey23   =   "22"
         MenuItemTag23   =   "1048"
         MenuItemIcon24  =   "frmMain.frx":2660
         MenuItemCaption24=   "放音转移"
         MenuItemKey24   =   "23"
         MenuItemTag24   =   "1049"
         MenuItemIcon25  =   "frmMain.frx":297A
         MenuItemCaption25=   "TTS 放音"
         MenuItemKey25   =   "28"
         MenuItemTag25   =   "1524"
         MenuItemIcon26  =   "frmMain.frx":2C94
         MenuItemCaption26=   "建立留言"
         MenuItemKey26   =   "40"
         MenuItemTag26   =   "1050"
         MenuItemIcon27  =   "frmMain.frx":2FAE
         MenuItemCaption27=   "察看留言"
         MenuItemKey27   =   "41"
         MenuItemTag27   =   "1051"
         MenuItemIcon28  =   "frmMain.frx":32C8
         MenuItemCaption28=   "节点连线"
         MenuItemKey28   =   "255"
         MenuItemTag28   =   "1045"
         MenuCaption3    =   "1034"
         MenuItemsMax3   =   7
         MenuItemIcon31  =   "frmMain.frx":35E2
         MenuItemCaption31=   "简单传真"
         MenuItemKey31   =   "50"
         MenuItemTag31   =   "1052"
         MenuItemIcon32  =   "frmMain.frx":38FC
         MenuItemCaption32=   "TTF传真"
         MenuItemKey32   =   "51"
         MenuItemTag32   =   "1053"
         MenuItemIcon33  =   "frmMain.frx":3C16
         MenuItemCaption33=   "传真接收"
         MenuItemKey33   =   "55"
         MenuItemTag33   =   "1677"
         MenuItemIcon34  =   "frmMain.frx":3F30
         MenuItemCaption34=   "呼叫外线号码"
         MenuItemKey34   =   "90"
         MenuItemTag34   =   "1054"
         MenuItemIcon35  =   "frmMain.frx":424A
         MenuItemCaption35=   "异步通信"
         MenuItemKey35   =   "96"
         MenuItemTag35   =   "1568"
         MenuItemIcon36  =   "frmMain.frx":4564
         MenuItemCaption36=   "Calling Card"
         MenuItemKey36   =   "91"
         MenuItemTag36   =   "1584"
         MenuItemIcon37  =   "frmMain.frx":487E
         MenuItemCaption37=   "节点连线"
         MenuItemKey37   =   "255"
         MenuItemTag37   =   "1045"
         MenuCaption4    =   "1035"
         MenuItemsMax4   =   8
         MenuItemIcon41  =   "frmMain.frx":4B98
         MenuItemCaption41=   "转接座席"
         MenuItemKey41   =   "60"
         MenuItemTag41   =   "1055"
         MenuItemIcon42  =   "frmMain.frx":4EB2
         MenuItemCaption42=   "转接座席组"
         MenuItemKey42   =   "61"
         MenuItemTag42   =   "1056"
         MenuItemIcon43  =   "frmMain.frx":51CC
         MenuItemCaption43=   "增强转接座席"
         MenuItemKey43   =   "62"
         MenuItemTag43   =   "1057"
         MenuItemIcon44  =   "frmMain.frx":54E6
         MenuItemCaption44=   "增强转接座席组"
         MenuItemKey44   =   "63"
         MenuItemTag44   =   "1058"
         MenuItemIcon45  =   "frmMain.frx":5800
         MenuItemCaption45=   "转虚拟分机"
         MenuItemKey45   =   "69"
         MenuItemTag45   =   "1059"
         MenuItemIcon46  =   "frmMain.frx":5B1A
         MenuItemCaption46=   "路由点查询"
         MenuItemKey46   =   "70"
         MenuItemTag46   =   "1600"
         MenuItemIcon47  =   "frmMain.frx":5E34
         MenuItemCaption47=   "查询座席状态"
         MenuItemKey47   =   "71"
         MenuItemTag47   =   "1633"
         MenuItemIcon48  =   "frmMain.frx":614E
         MenuItemCaption48=   "节点连线"
         MenuItemKey48   =   "255"
         MenuItemTag48   =   "1045"
         MenuCaption5    =   "1036"
         MenuItemsMax5   =   6
         MenuItemIcon51  =   "frmMain.frx":6468
         MenuItemCaption51=   "记录变量"
         MenuItemKey51   =   "102"
         MenuItemTag51   =   "1063"
         MenuItemIcon52  =   "frmMain.frx":6782
         MenuItemCaption52=   "用户DLL"
         MenuItemKey52   =   "100"
         MenuItemTag52   =   "1064"
         MenuItemIcon53  =   "frmMain.frx":6A9C
         MenuItemCaption53=   "用户COM"
         MenuItemKey53   =   "101"
         MenuItemTag53   =   "1065"
         MenuItemIcon54  =   "frmMain.frx":6DB6
         MenuItemCaption54=   "发送数据"
         MenuItemKey54   =   "18"
         MenuItemTag54   =   "1043"
         MenuItemIcon55  =   "frmMain.frx":70D0
         MenuItemCaption55=   "无操作"
         MenuItemKey55   =   "19"
         MenuItemTag55   =   "1044"
         MenuItemIcon56  =   "frmMain.frx":73EA
         MenuItemCaption56=   "节点连线"
         MenuItemKey56   =   "255"
         MenuItemTag56   =   "1045"
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   5505
      Left            =   7740
      ScaleHeight     =   5445
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
      Begin VOXPLAYERLib.VOXPlayer VOXPlayer 
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   5040
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   38
         Left            =   510
         Stretch         =   -1  'True
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   1
         Left            =   270
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   2
         Left            =   540
         Picture         =   "frmMain.frx":7704
         Stretch         =   -1  'True
         Top             =   30
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   3
         Left            =   750
         Picture         =   "frmMain.frx":7806
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   4
         Left            =   90
         Picture         =   "frmMain.frx":7908
         Stretch         =   -1  'True
         Top             =   270
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   5
         Left            =   300
         Picture         =   "frmMain.frx":7A0A
         Stretch         =   -1  'True
         Top             =   270
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   6
         Left            =   540
         Stretch         =   -1  'True
         Top             =   270
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   7
         Left            =   750
         Stretch         =   -1  'True
         Top             =   270
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   8
         Left            =   60
         Stretch         =   -1  'True
         Top             =   510
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   9
         Left            =   300
         Stretch         =   -1  'True
         Top             =   510
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   10
         Left            =   510
         Stretch         =   -1  'True
         Top             =   510
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   11
         Left            =   750
         Stretch         =   -1  'True
         Top             =   510
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   12
         Left            =   30
         Stretch         =   -1  'True
         Top             =   750
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   13
         Left            =   240
         Picture         =   "frmMain.frx":7B0C
         Stretch         =   -1  'True
         Top             =   750
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   14
         Left            =   480
         Picture         =   "frmMain.frx":7C0E
         Stretch         =   -1  'True
         Top             =   750
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   15
         Left            =   720
         Stretch         =   -1  'True
         Top             =   750
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   16
         Left            =   30
         Stretch         =   -1  'True
         Top             =   990
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   17
         Left            =   270
         Stretch         =   -1  'True
         Top             =   990
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   18
         Left            =   480
         Picture         =   "frmMain.frx":7D10
         Stretch         =   -1  'True
         Top             =   990
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   19
         Left            =   720
         Picture         =   "frmMain.frx":8252
         Stretch         =   -1  'True
         Top             =   990
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   20
         Left            =   30
         Picture         =   "frmMain.frx":8354
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   22
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   23
         Left            =   690
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   24
         Left            =   30
         Stretch         =   -1  'True
         Top             =   1470
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   25
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1470
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   26
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   27
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1470
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   28
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1710
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   29
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1710
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   30
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   31
         Left            =   690
         Picture         =   "frmMain.frx":8456
         Stretch         =   -1  'True
         Top             =   1710
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   32
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   21
         Left            =   270
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   33
         Left            =   240
         Picture         =   "frmMain.frx":8998
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   34
         Left            =   450
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   35
         Left            =   660
         Picture         =   "frmMain.frx":8EDA
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   36
         Left            =   30
         Stretch         =   -1  'True
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image Img 
         Height          =   240
         Index           =   37
         Left            =   270
         Picture         =   "frmMain.frx":8FDC
         Stretch         =   -1  'True
         Top             =   2190
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "新建"
            Object.ToolTipText     =   "1083"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打开"
            Object.ToolTipText     =   "1084"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "保存"
            Object.ToolTipText     =   "1085"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打印"
            Object.ToolTipText     =   "1086"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "剪切"
            Object.ToolTipText     =   "1087"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "复制"
            Object.ToolTipText     =   "1088"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "粘贴"
            Object.ToolTipText     =   "1089"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "撤销"
            Object.ToolTipText     =   "1090"
            ImageKey        =   "UnDo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "重复"
            Object.ToolTipText     =   "1091"
            ImageKey        =   "ReDo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "删除节点"
            Object.ToolTipText     =   "1092"
            ImageKey        =   "Dustbin"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "节点属性"
            Object.ToolTipText     =   "1093"
            ImageKey        =   "Property"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "显示节点描述"
            Object.ToolTipText     =   "1018"
            ImageKey        =   "Label"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "节点列表"
            Object.ToolTipText     =   "1094"
            ImageKey        =   "NodeList"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "变量使用"
            Object.ToolTipText     =   "1095"
            ImageKey        =   "Viewer"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "左对齐"
            Object.ToolTipText     =   "1096"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "置中"
            Object.ToolTipText     =   "1097"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "右对齐"
            Object.ToolTipText     =   "1098"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "节点连线"
            Object.ToolTipText     =   "1045"
            ImageKey        =   "NodeLine"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "选项"
            Object.ToolTipText     =   "1099"
            ImageKey        =   "Tools"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "拨放语音"
            Object.ToolTipText     =   "1523"
            ImageKey        =   "Play"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "停止放音"
            Object.ToolTipText     =   "1101"
            ImageKey        =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6720
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":90DE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":91F0
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9302
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9414
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9526
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9638
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":974A
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":985C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":996E
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A80
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B92
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9CA4
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DB6
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9EC8
            Key             =   "Description"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A31A
            Key             =   "Dustbin"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CACC
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CC26
            Key             =   "Label"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CD80
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CEDA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D1F4
            Key             =   "Tools"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E9B6
            Key             =   "Viewer"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ECD0
            Key             =   "UnDo"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EE2A
            Key             =   "ReDo"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EF84
            Key             =   "NodeList"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F31E
            Key             =   "NodeLine"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FFF8
            Key             =   "Stop2"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   7470
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   5925
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2647
            Picture         =   "frmMain.frx":10312
            Key             =   "keyPosition"
            Object.ToolTipText     =   "1073"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Picture         =   "frmMain.frx":10460
            Key             =   "keySize"
            Object.ToolTipText     =   "1074"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Key             =   "keyTip"
            Object.ToolTipText     =   "1075"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "1072"
            TextSave        =   "1072"
            Key             =   "Root"
            Object.ToolTipText     =   "1076"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "1070"
            TextSave        =   "1070"
            Key             =   "Resource"
            Object.ToolTipText     =   "1077"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "1071"
            TextSave        =   "1071"
            Key             =   "Language"
            Object.ToolTipText     =   "1078"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "Page"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "16:54"
            Object.ToolTipText     =   "1079"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1060
            MinWidth        =   1060
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "1080"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1060
            MinWidth        =   1060
            TextSave        =   "NUM"
            Object.ToolTipText     =   "1081"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mFile 
      Caption         =   "1001"
      Begin VB.Menu mnuFile 
         Caption         =   "1692"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1006"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1007"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1693"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1067"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1008"
         Index           =   6
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1068"
         Index           =   8
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1009"
         Index           =   9
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1010"
         Index           =   10
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1622"
         Index           =   11
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "1011"
         Index           =   19
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mResource 
      Caption         =   "1694"
      Begin VB.Menu mnuResource 
         Caption         =   "1712"
         Index           =   0
      End
      Begin VB.Menu mnuResource 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuResource 
         Caption         =   "1695"
         Index           =   2
      End
      Begin VB.Menu mnuResource 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuResource 
         Caption         =   "1696"
         Index           =   4
      End
      Begin VB.Menu mnuResource 
         Caption         =   "1697"
         Index           =   5
      End
      Begin VB.Menu mnuResource 
         Caption         =   "1713"
         Index           =   6
      End
      Begin VB.Menu mnuResource 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuResource 
         Caption         =   "1008"
         Index           =   8
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "1002"
      Begin VB.Menu mnuEdit 
         Caption         =   "1012"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1013"
         Index           =   1
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1014"
         Index           =   3
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1021"
         Index           =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1022"
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1015"
         Index           =   7
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1016"
         Index           =   8
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1017"
         Index           =   9
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1018"
         Index           =   10
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1945"
         Index           =   11
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1019"
         Enabled         =   0   'False
         Index           =   13
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "1020"
         Enabled         =   0   'False
         Index           =   14
      End
   End
   Begin VB.Menu mData 
      Caption         =   "1003"
      Visible         =   0   'False
      Begin VB.Menu mnudata 
         Caption         =   "1021"
         Index           =   0
      End
      Begin VB.Menu mnudata 
         Caption         =   "1022"
         Index           =   1
      End
      Begin VB.Menu mnudata 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnudata 
         Caption         =   "1622"
         Index           =   3
      End
   End
   Begin VB.Menu mView 
      Caption         =   "1004"
      Begin VB.Menu mnuView 
         Caption         =   "1023"
         Index           =   0
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuView 
         Caption         =   "1024"
         Index           =   1
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuView 
         Caption         =   "1729"
         Index           =   3
      End
      Begin VB.Menu mnuView 
         Caption         =   "1025"
         Index           =   4
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuView 
         Caption         =   "1026"
         Index           =   6
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuView 
         Caption         =   "1912"
         Index           =   7
      End
   End
   Begin VB.Menu mWindows 
      Caption         =   "1665"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindows 
         Caption         =   "1666"
         Index           =   0
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "1667"
         Index           =   1
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "1668"
         Index           =   2
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "1005"
      Begin VB.Menu mnuHelp 
         Caption         =   "1027"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "1028"
         Index           =   1
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "1029"
         Index           =   2
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "1030"
         Index           =   4
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "1031"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'variable to hold the last-sized postion
Private currSplitPosX As Long

Private m_frmCallFlow As CFlowWorks
Private m_lstCFForms(256) As CFlowWorks
Private m_SysMenu As Long
Private wlOldProc As Long

Private Sub MDIForm_Load()

    mnuHelp(5).Caption = LoadNationalResString(1031)
    Set m_frmCallFlow = New CFlowWorks
        
    'set the current splitter bar position to an arbitrary value that will always be outside
    'the possibe range. This allows us to check for movement of the spltter bar in subsequent
    'mousexxx subs.
    currSplitPosX = &H7FFFFFFF
    
    '========================================
    'Append a system menu called "About..."
    '========================================
    m_SysMenu = GetSystemMenu(Me.hWnd, False)
    If m_SysMenu Then
        If mnuHelp(5).Caption <> "" Then
            AppendMenu m_SysMenu, MF_SEPARATOR, ByVal 0&, ByVal 0&
            AppendMenu m_SysMenu, MF_STRING, IDM_ABOUT, Me.mnuHelp(5).Caption
        End If
    End If
    
    wlOldProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf OwnMenuProc)
    
    '' Sun added 2002-06-11
    Dim lv_Loop As Integer
    For lv_Loop = Me.Img.LBound To Me.Img.UBound
        Img(lv_Loop).MousePointer = 99
        Img(lv_Loop).Stretch = True
        Img(lv_Loop).Height = Screen.TwipsPerPixelY * 32
        Img(lv_Loop).Width = Screen.TwipsPerPixelX * 32
    Next
    
    If gblnShowNodeCaption Then
        mnuEdit(10).Caption = "1069"
    Else
        mnuEdit(10).Caption = "1018"
    End If
    
    'Mike add this @ 08-1-29 for show/hide the node tag
    If gblnShowNodeTag Then
        mnuEdit(11).Caption = "1946"
    Else
        mnuEdit(11).Caption = "1945"
    End If
    
    LoadResStrings Me
    
    '' Sun added 2012-05-07
    ''' Load MRU menu list
    Call LoadMRUMenu
    
    StatusBar.Panels("Resource").Text = LoadNationalResString(1070)
    StatusBar.Panels("Language").Text = LoadNationalResString(1071)
    tbToolBar.Buttons("显示节点描述").ToolTipText = mnuEdit(10).Caption
    ToolBar.MenuCur = 2
    
    m_frmCallFlow.Show
    
End Sub

'' Sun added 2012-05-07
''' Load MRU menu list
Private Sub LoadMRUMenu()

Dim lv_blnShowItem As Boolean
Dim lv_Loop As Integer
Dim lv_nMenuBase As Integer
Dim lv_nMenuIndex As Integer

lv_blnShowItem = False
lv_nMenuBase = 12
For lv_Loop = 1 To PCS_MAX_MRUITEMS
    lv_nMenuIndex = lv_nMenuBase + lv_Loop
    If gMRUMenu(lv_Loop).blnVisible Then
        mnuFile(lv_nMenuIndex).Visible = True
        mnuFile(lv_nMenuIndex).Caption = "&" & Trim(Str(lv_Loop)) & " [" & Trim(Str(gMRUMenu(lv_Loop).intItemID)) & "] " & gMRUMenu(lv_Loop).strItemName
        mnuFile(lv_nMenuIndex).Tag = gMRUMenu(lv_Loop).intItemID
        lv_blnShowItem = True
    Else
        Exit For
    End If
Next

'' Seperator
mnuFile(lv_nMenuBase + PCS_MAX_MRUITEMS + 1).Visible = lv_blnShowItem

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    If wlOldProc <> 0 Then
         SetWindowLong hWnd, GWL_WNDPROC, wlOldProc
    End If
    
    'Unload all froms before main window is destroyed
    Call CloseAllForms
    
    '' Sun added 2012-05-07
    ''' Save MRU Data
    Call SaveMRUMData
    
    'Michael Added @ 2008-7-4 for write program shutdown log
    WriteLogMessage 0, enu_Information, App.Title & " Shutdown..."
    
    'Michael Added @ 2008-7-1 for delete empty log file
    'Call DelBlankLog

End Sub

'' Sun added 2006-01-27
Public Sub CloseAllForms(Optional ByVal blnAllForm As Boolean = True)

    Dim objForm As Form

    For Each objForm In Forms
        
        If blnAllForm Or objForm.Tag = "1662" Then
            Unload objForm
        End If
    Next

End Sub

Private Function OpenSaveFlowFile(lOwnerhWnd As Long, lStyle As Integer) As String
    Dim lv_strTitle As String
    Dim lv_strFilter As String
    Dim lv_strDefExt As String
    Dim lv_blnSaveDlg As Boolean
    
    lv_blnSaveDlg = True
    Select Case lStyle
        Case 0
            lv_strTitle = LoadNationalResString(1732)
            lv_blnSaveDlg = False
        Case 1
            lv_strTitle = LoadNationalResString(1733)
        Case 2            '' Import Call Flow
            lv_strTitle = LoadNationalResString(1734)
            lv_blnSaveDlg = False
        Case 3            '' Export Call Flow
            lv_strTitle = LoadNationalResString(1735)
        Case 4            '' Import Resource
            lv_strTitle = LoadNationalResString(1736)
            lv_blnSaveDlg = False
        Case 5            '' Export Resource
            lv_strTitle = LoadNationalResString(1737)
    End Select
    
    If lStyle < 4 Then
        lv_strFilter = LoadNationalResString(1738)
        lv_strDefExt = "PVC"
    Else
        lv_strFilter = LoadNationalResString(1739)
        lv_strDefExt = "CSV"
    End If
    
    OpenSaveFlowFile = F_OpenFileDialog(lOwnerhWnd, lv_blnSaveDlg, lv_strTitle, lv_strFilter, lv_strDefExt)

End Function

' Owner menu
Public Function MsgProc(ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long

    'This procedure is called because we've subclassed
    'this form. We will catch DRAWITEM and MEASUREITEM
    'messages and pass the rest of them on.

    'Various structs we'll need
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim DrawInfo As DRAWITEMSTRUCT
    Dim mii As MENUITEMINFO
    'Set later for separator flag:
    Dim IsSep As Boolean
    'Our custom brush and the old one
    Dim hBr As Long, hOldBr As Long
    'Our custom pen and the old one
    Dim hPen As Long, hOldPen As Long
    'The text color of the menu items
    Dim lTextColor As Long
    'Now much to bump the menu's selection
    'rectangle over
    Dim iRectOffset As Integer

    Select Case wMsg
        Case WM_SYSCOMMAND
            If LoWord(wParam) = IDM_ABOUT Then frmAbout.Show vbModal

'        Case WM_DRAWITEM
'            If wParam = 0 Then 'It was sent by the menu
'                'Get DRAWINFOSTRUCT -- copy it to our
'                'empty structure from the pointer in lParam
'                Call CopyMem(DrawInfo, ByVal lParam, Len(DrawInfo))
'                IsSep = IsSeparator(DrawInfo.itemID)
'
'                '===Set the menu font through its hDC...===
'                MyFont = SendMessage(Me.hWnd, WM_GETFONT, 0&, 0&)
'                OldFont = SelectObject(DrawInfo.hdc, MyFont)
'                'We draw the item based on Un/Selected:
'
'                If DrawInfo.itemState = ODS_SELECTED Then
'                    hBr = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
'                    hPen = GetPen(1, GetSysColor(COLOR_HIGHLIGHT))
'                    lTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
'                Else
'
'                    hBr = CreateSolidBrush(GetSysColor(COLOR_MENU))
'                    hPen = GetPen(1, GetSysColor(COLOR_MENU))
'                    lTextColor = GetSysColor(COLOR_MENUTEXT)
'                End If
'
'                  'We're going to draw on the menu
'                  QuickGDI.TargethDC = DrawInfo.hdc
'                  'Select our new, correctly colored objects:
'                  hOldBr = SelectObject(DrawInfo.hdc, hBr)
'                  hOldPen = SelectObject(DrawInfo.hdc, hPen)
'                    With DrawInfo.rcItem
'                        If DrawInfo.itemState <> ODS_SELECTED Then
'                                  'Clear the space where the image is
'                             QuickGDI.DrawRect .Left, .Top, _
'                                  28, .Bottom
'                        End If
'                        'Check to see if the menu item is one of the
'                        'ones with a picture. If so, then we need to
'                        'move the edge of the drawing rectangle a little
'                        'to the left to make room for the image.
'                        iRectOffset = IIf(Img(DrawInfo.itemID).Picture _
'                             <> 0, 23, 0)
'                        'Do we have a separator bar?
'                        If Not IsSep Then
'                             'Draw the rectangle onto the item's space
'                             QuickGDI.DrawRect .Left + iRectOffset, _
'                                  .Top, .Right, .Bottom
'                             'Print the item's text
'                             '(held in the Caps() array)
'                             .Left = .Left + 25
'                             hPrint DrawInfo.rcItem, _
'                                  Caps(DrawInfo.itemID), _
'                                  lTextColor
'                        End If
'                    End With
'                'Select the old objects into the menu's DC
'                Call SelectObject(DrawInfo.hdc, hOldBr)
'                Call SelectObject(DrawInfo.hdc, hOldPen)
'                'Delete the ones we created
'                Call DeleteObject(hBr)
'                Call DeleteObject(hPen)
'                With DrawInfo
'                    'If the item had an image:
'                    '2 = New, 3 = Open, 4 = Save, etc.
'                    If Img(.itemID).Picture.Handle <> 0 Then
'                        pnt.PaintTransparentStdPic .hdc, _
'                             4, .rcItem.Top + 2, _
'                             16, 16, Img(.itemID).Picture, _
'                             0, 0, &HC0C0C0
'                        'If this item is selected, draw a raised
'                        'box around the image
'                        If DrawInfo.itemState = ODS_SELECTED Then ThreedBox 1, .rcItem.Top, 21, .rcItem.Bottom - 1
'                    End If
'                    If IsSep Then
'                        'Draw the special separator bar
'                        ThreedBox .rcItem.Left, .rcItem.Top + 2, .rcItem.Right - 1, .rcItem.Bottom - 2, True
'                    End If
'                End With
'            End If
'            'Don't pass this message on:
'            MsgProc = False
'            Exit Function
'
'        Case WM_MEASUREITEM
'            'Get the MEASUREITEM struct from the pointer
'            Call CopyMem(MeasureInfo, ByVal lParam, Len(MeasureInfo))
'            IsSep = IsSeparator(MeasureInfo.itemID)
'            'Tell Windows how big our items are.
'            MeasureInfo.itemWidth = 170
'
'            'If the item being measured is the separator
'            'bar, the height should be 5 pixels, 18 if
'            'otherwise...
'            MeasureInfo.itemHeight = IIf(IsSep, 5, GetSystemMetrics(SM_CYMENU))
'            'Return the information back to Windows
'            Call CopyMem(ByVal lParam, MeasureInfo, Len(MeasureInfo))
'            'Don't pass this message on:
'            MsgProc = False
'            Exit Function
'
'        Case WM_MENUSELECT

    End Select

    'We didn't handle this message,
    'pass it on to the next WndProc
    MsgProc = CallWindowProc(wlOldProc, hWnd, wMsg, wParam, lParam)
End Function

Private Function HiWord(LongIn As Long) As Integer
     HiWord = (LongIn And &HFFFF0000) \ &H10000
End Function

Private Function LoWord(LongIn As Long) As Integer
     If (LongIn And &HFFFF&) > &H7FFF Then
          LoWord = (LongIn And &HFFFF&) - &H10000
     Else
          LoWord = LongIn And &HFFFF&
     End If
End Function

Public Function IsSeparator(ByVal IID As Integer) As Boolean
     Dim mii As MENUITEMINFO
     mii.cbSize = Len(mii)
     mii.fMask = MIIM_TYPE
     mii.wID = IID
     GetMenuItemInfo GetMenu(hWnd), IID, False, mii
     IsSeparator = ((mii.fType And MFT_SEPARATOR) = MFT_SEPARATOR)
End Function

Private Sub mEdit_Click()
    CheckEditPopMenu (gCallFlow.NodeSelectedID > 0)
End Sub

Private Sub mnudata_Click(Index As Integer)

    ' Sun added 2002-04-02
    If m_frmCallFlow.WorkFrame.Enabled = False Then Exit Sub
    
        Select Case Index
            Case 0
               SystemDefault.Show vbModal
               
            Case 1
               ' Show LOG and Variables View
                If gCallFlow.CallFlowID = 0 Then
                    Message ("M131")
                    Exit Sub
                Else
                    frmViewLogVar.Show vbModal, Me
                End If

            Case 2
               
            Case 3
                ' Call Flow Synchronize
                If gCallFlow.CallFlowID = 0 Then
                    Message ("M131")
                    Exit Sub
                Else
                    gSystem.intConfigSet = 0
                    frmSyncCallFlow.Show vbModal, Me
                End If
                
       End Select
End Sub

'**********************************************************************************************************************
' -----------------------------------------   菜单栏函数   ------------------------------------------------------------
'**********************************************************************************************************************
'Michael : Menu File event   "流程"菜单
Private Sub mnuFile_Click(Index As Integer)
On Error Resume Next
        
    Dim msgresult As VbMsgBoxResult
    Dim lv_fromPage As Integer, lv_toPage As Integer
    Dim lv_Index As Integer
    Dim lv_tempFile As String
    Dim lv_bytPID As Byte
    Dim lv_bytOldPID As Byte
    
    ' Sun added 2002-04-02
    If m_frmCallFlow.WorkFrame.Enabled = False Then Exit Sub
    
    Select Case Index
        Case 0   'New
            Frm_FlowCreate.Show 0, Me
            If gSystem.intConfigSet > 0 Then
                frmOptions.Show vbModal
            End If
            
        Case 1   'Open
            '' Sun added 2007-03-25
            ''' 打开已有流程
            frm_FlowOpen.Show 0, Me
            If gSystem.intConfigSet > 0 Then
                frmOptions.Show vbModal
            End If
            
        Case 2   'Save
            If Not gCallFlow.SavedMark Then
                ' Add storing flow data to disk/database procedure code here...
                gCallFlow.UpdateIvrTable
            End If

        Case 3   'Save As
            If gCallFlow.CallFlowID = 0 Then
                Message ("M131")
                Exit Sub
            Else
                frm_FlowSaveAs.Show vbModal
            End If
        '--------------------------------------------------------------------------------
        Case 5   'Zoom
        
        Case 6   'Print
            If gCallFlow.CallFlowID <> 0 Then
                With dlgCommonDialog
                    .DialogTitle = "Print"
                    .CancelError = True
                    
                    '''' Sun modified 2002-03-31
                    '.flags = cdlPDReturnDC + cdlPDHidePrintToFile
                    '.flags = .flags + cdlPDNoPageNums
                    '.flags = .flags + cdlPDAllPages
                    .flags = cdlPDReturnDC + cdlPDAllPages
                    
                    .ShowPrinter
                    If Err <> MSComDlg.cdlCancel Then
                        Debug.Print .FromPage, .ToPage, .flags
                        If gCallFlow.CallFlowID = gGetUserFlow.IntP_id Then
                            Printer.Print LoadNationalResString(1103) & Trim(Str(gCallFlow.CallFlowID)) & "）" & Trim(gGetUserFlow.StrP_name)
                            Printer.Print LoadNationalResString(1104) & Trim(gGetUserFlow.StrP_description)
                            Printer.Print LoadNationalResString(1105) & Trim(gGetUserFlow.StrP_auther) & "  用户：" & Trim(gGetUserFlow.StrP_user)
                            Printer.Print LoadNationalResString(1106) & Format(gGetUserFlow.DateP_createtime, "YYYY-MM-DD hh:mm:ss") & "  最新修改：" & Format(gGetUserFlow.DataP_modifytime, "YYYY-MM-DD hh:mm:ss")
                        End If
                        
                        If .FromPage > 0 Then
                            lv_fromPage = .FromPage
                        Else
                            lv_fromPage = 1
                        End If
                        
                        If .ToPage = 0 Then
                            lv_toPage = gCallFlow.PageCount
                        ElseIf .ToPage > lv_fromPage Then
                            lv_toPage = .ToPage
                        Else
                            lv_toPage = lv_fromPage
                        End If
                        
                        For lv_Index = lv_fromPage To lv_toPage
                            Call m_frmCallFlow.PrintWorkPage(lv_Index)
                            If lv_Index < lv_toPage Then
                                Printer.NewPage
                            End If
                            SleepVB 200
                        Next

                        Printer.EndDoc
                    End If
                End With
            End If
      '-----------------------------------------------------------------------------------------------------------------
        Case 8   'Load Flow on to IVR Server
            Frm_load.Show vbModal
            
        Case 9   'Import Flow
            lv_tempFile = Trim(OpenSaveFlowFile(Me.hWnd, 2))
            If lv_tempFile <> "" Then
            
                '' Confirm File Info
                lv_bytOldPID = gCallFlow.CallFlowID
                Set gCallFlow = New clsIVRProgram
                
                lv_bytPID = gCallFlow.GetFlowDataFileInfo(lv_tempFile)
                gCallFlow.DestroyAllNodes
                If lv_bytPID > 0 Then
                
                    '' Whether Opened?
                    If IsCallFlowOpened(lv_bytPID) Then
                        m_lstCFForms(lv_bytPID).ClearFormContent
                        SwitchToMDIForm lv_bytPID
                    Else
                        If lv_bytOldPID > 0 Then
                            '' Initalize new call flow worksheet
                            CreateNewMDIForm lv_bytPID
                        Else
                            AssignMDIForm lv_bytPID
                            m_lstCFForms(lv_bytPID).SetFormActive
                        End If
                    End If
                    
                    '' Import File Data
                    gCallFlow.OpenIvrRecordSet lv_bytPID
                    If gCallFlow.ImportFlowData(lv_tempFile) Then
                        gCallFlow.OpenIvrRecordSet lv_bytPID
                        ShowCallFlowOnScreen
                        Call Message("M134")
                      
                    End If
                    
                End If
            
            End If
            
        Case 10   'Export Flow
            
            If gCallFlow.CallFlowID = 0 Then
                Message ("M131")
            Else
                lv_tempFile = Trim(OpenSaveFlowFile(Me.hWnd, 3))
                If lv_tempFile <> "" Then
                    If gCallFlow.ExportFlowData(lv_tempFile) Then
                        Call Message("M132")
                    End If
                End If
            End If
            
        Case 11     ' 流程同步
            Call mnudata_Click(3)
        '------------------------------------------------------------------------------------------------------
        
        '' Sun added 2012-05-07
        Case 13, 14, 15, 16, 17             '' MRU List
            
            lv_bytPID = CByte(Val(mnuFile(Index).Tag))
            If lv_bytPID > 0 Then
            
                ' 需要判断流程是否已经打开
                If frmMain.IsCallFlowOpened(lv_bytPID) Then
                    Message "M142"
                    frmMain.SwitchToMDIForm lv_bytPID
                    Exit Sub
                End If
                              
                If gCallFlow.CallFlowID > 0 Then
                    '' Enable Original Call Flow Form
                    SetMainFormItemsEnableWhenPropertyShow True
                End If
                 
                '' Open New CallFlow Window
                If Not frmMain.CreateNewMDIForm(lv_bytPID) Then
                    Exit Sub
                Else
                    ' 打开新流程
                    Call gCallFlow.OpenIvrRecordSet(lv_bytPID)

                    frmMain.ShowCallFlowOnScreen
                End If
                
            End If

        Case 19   ' Exit
            Unload Me
    End Select
End Sub

'Michael : Menu Resource event   "资源"菜单
Private Sub mnuResource_Click(Index As Integer)
    Dim lv_tempFile As String
    Dim lv_bytPID As Byte
    
    ' Sun added 2002-04-02
    If m_frmCallFlow.WorkFrame.Enabled = False Then Exit Sub

    Select Case Index
        'Michael : this action is not complate ,add @Jul,12,07
        Case 0   ' 资源项目管理
            gSystem.intCurStep = -1
            Set gSystem.crlCurItem = Nothing
            frmProjectList.Show vbModal
            
        '--------------------------------------------------------
        Case 2   ' 资源清单
            gSystem.intCurStep = -1
            Set gSystem.crlCurItem = Nothing
            gbCallFromPro = 0
            frmResourceList.Show vbModal
        '--------------------------------------------------------
        Case 4   ' 导入资源清单
            lv_tempFile = Trim(OpenSaveFlowFile(Me.hWnd, 4))
            If lv_tempFile <> "" Then
            
                '' Confirm File Info & Import File Data
                lv_bytPID = F_ConfirmResourceFileInfo(lv_tempFile)
                If lv_bytPID > 0 Then
                
                    '' Refresh Opened Resource
                    Call ReopenResourceList(lv_bytPID)
                    
                    Call Message("M147")
                    
                End If
            End If

        Case 5   ' 导出资源清单
            If gCallFlow.CallFlowID = 0 Then
                Message ("M131")
                Exit Sub
            Else
                lv_tempFile = Trim(OpenSaveFlowFile(Me.hWnd, 5))
                If lv_tempFile <> "" Then
                    If gCallFlow.ExportResourceData(lv_tempFile) Then
                        Call Message("M146")
                    End If
                End If
            End If
        
        Case 6   ' 资源同步
            If gCallFlow.CallFlowID = 0 Then
                Message ("M131")
                Exit Sub
            Else
                gSystem.intConfigSet = 1
                frmSyncCallFlow.Show vbModal, Me
            End If
        '------------------------------------------------------------------------------
        Case 8   ' 打印资源清单
            If gCallFlow.CallFlowID = 0 Then
                Message ("M131")
                Exit Sub
            Else
                'Michael Note: Print function
                gSystem.intCurStep = -1
                Set gSystem.crlCurItem = Nothing
                'Modified @ Sep,7,07 -> Michael
                'frmPrintRes.Show vbModal
                frmResPrintDoc.Show vbModal
            End If
            
    End Select
End Sub

'Michael : Menu Edit event  "编辑"菜单
Private Sub mnuEdit_Click(Index As Integer)
On Error Resume Next

    Dim lv_nNodeCount As Integer
    Dim lv_nIndex As Integer
    
    ' Sun added 2002-04-02
    If m_frmCallFlow.WorkFrame.Enabled = False Then Exit Sub
    
    Select Case Index
        Case 0   'Copy Code
        
            m_frmCallFlow.CopyContentsOnWorkPage
            
        Case 1   'Paste Code
            
            If gClipBoard.ClipboardMark Then
                m_frmCallFlow.PasteContentsOnWorkPage
            End If
    '----------------------------------------------------------
        Case 3  ' Flow information  流程属性
            gSystem.intConfigSet = 0
            frmFlowProperty.Show vbModal, Me

        Case 4  ' 流程参数
            gSystem.intConfigSet = 1
            frmFlowProperty.Show vbModal, Me

        Case 5  ' 变量使用
            gSystem.intConfigSet = 2
            frmFlowProperty.Show vbModal, Me
    '-----------------------------------------------------------
'        Case 6   'Create Node
'            Frm_NodeCreate.Show vbModal

        Case 7  ' Set Root Node
            
            If gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID >= 256 And gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID <> gCallFlow.RootNodedID Then
                F_SetRootNode gCallFlow.NodeSelectedID
            End If
            
        Case 8  'Delete Node
            lv_nNodeCount = gCallFlow.SelectedCount
            If lv_nNodeCount > 0 Then
                
                If gClipBoard.MultiPushInitialize(lv_nNodeCount, DEF_OPERATION_DELETE) Then
                    
                    For lv_nIndex = 1 To gCallFlow.NewNodeID
                        If gCallFlow.Node(lv_nIndex).IsSelected Then
                                                    
                            If gCallFlow.Node(lv_nIndex).NodeID >= 256 Then
                                
                                If gCallFlow.Node(lv_nIndex).NodeID <> gCallFlow.RootNodedID Then
                                
                                    gClipBoard.MultiPushClipBoardStack lv_nIndex
                                    gCallFlow.DeleteIvrRecord lv_nIndex
                                    gCallFlow.Node(lv_nIndex).Destroy
                                    
                                Else
                                    gClipBoard.MultiPushClipBoardStack DEF_OPERATION_INIT
                                    Call Message("M135")
                                End If
                                
                            Else
                                gClipBoard.MultiPushClipBoardStack DEF_OPERATION_INIT
                                Call Message("M133")
                            End If
                            
                        End If
                    Next
                
                End If
                
                gCallFlow.NodeSelectedID = 0
                
            End If
            
        Case 9   'Node Property   Modify Node
            m_frmCallFlow.ShowNodeProp gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
        
        Case 10  ' Show/Hide Node Caption
            gblnShowNodeCaption = Not gblnShowNodeCaption
            gCallFlow.SetAllNodeCaptionVisible gblnShowNodeCaption
            
            If gblnShowNodeCaption Then
                mnuEdit(10).Caption = LoadNationalResString(1069)
                Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_CAPTION, "1", gSystem.strINI_File)
            Else
                mnuEdit(10).Caption = LoadNationalResString(1018)
                Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_CAPTION, "0", gSystem.strINI_File)
            End If
            tbToolBar.Buttons("显示节点描述").ToolTipText = mnuEdit(10).Caption
        
        'Mike Add @ 2008-1-29
        Case 11  'Show/Hide Node Tag
            gblnShowNodeTag = Not gblnShowNodeTag
            gCallFlow.SetAllNodeTagVisible gblnShowNodeTag
            
            If gblnShowNodeTag Then
                mnuEdit(11).Caption = LoadNationalResString(1946)
                Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_TAG, "1", gSystem.strINI_File)
            Else
                mnuEdit(11).Caption = LoadNationalResString(1945)
                Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_TAG, "0", gSystem.strINI_File)
            End If
        ' Mike Added End
        '-------------------------------------------------------------------------------------------------------
        '' Sun added 2002-06-11
        Case 13  ' UnDo
            Call gClipBoard.UnDoOperation
            m_frmCallFlow.RefreshHandlesPosition
            
        Case 14 ' ReDo
            Call gClipBoard.ReDoOperation
            m_frmCallFlow.RefreshHandlesPosition
            
        End Select
End Sub
    
'Michael : Menu View action     "查看"菜单
' ! : 工作页面属性未实现
Private Sub mnuView_Click(Index As Integer)
    ' Sun added 2002-04-02
    If m_frmCallFlow.WorkFrame.Enabled = False Then Exit Sub
    
    Dim lv_Str As String
    
    Select Case Index
        'Michael : UnComplate
        Case 0   ' WorkPage Properties
        
        Case 1   ' Goto Page...
'            lv_Str = InputBox(LoadNationalResString(1107), LoadNationalResString(1108), Trim(Str(gCallFlow.CurrentPage)))
'            If Len(Trim(lv_Str)) > 0 Then
'                Mdlfunction.GotoAnotherPage Val(lv_Str)
'            End If
            frmGotoPage.Show vbModal
        '--------------------------------------------------------------------------------
        Case 3   ' 日志列表
            Call mnudata_Click(1)
            
        Case 4   ' 节点列表
            If gCallFlow.CallFlowID = 0 Then
                Message ("M131")
                Exit Sub
            Else
                Set gSystem.crlCurItem = Nothing
                frmNodeList.Show vbModal
            End If
        '--------------------------------------------------------------------------------
        Case 6   ' Options...
            gSystem.intConfigSet = 0
            frmOptions.Show vbModal
        'Michael Added @ 2007-11-28
        Case 7   ' TTS Setting
            frmTTSSetting.Show vbModal
    End Select
End Sub

'Michael : Menu Window event  "窗口"菜单
Private Sub mnuWindows_Click(Index As Integer)
    Select Case Index
    Case 0              '' 层叠
        Me.Arrange vbCascade
    Case 1              '' 横排
        Me.Arrange vbTileHorizontal
    Case 2              '' 纵排
        Me.Arrange vbTileVertical
    End Select
End Sub

'Michael : Menu Help event  "帮助"菜单
' ! : no "VisualIVREditor.hlp" file
Private Sub mnuHelp_Click(Index As Integer)
    Select Case Index
        Case 0  ' Contents...
            ShellExecute Me.hWnd, "Open", App.path + "\help\VisualIVREditor.hlp", vbNullString, vbNullString, 1
        Case 1  ' Index...
            ShellExecute Me.hWnd, "Open", App.path + "\help\VisualIVREditor.hlp", vbNullString, vbNullString, 1
        Case 2  ' Search...
            ShellExecute Me.hWnd, "Open", App.path + "\help\VisualIVREditor.hlp", vbNullString, vbNullString, 1
        Case 4  ' Corporation Web Site
            ShellExecute Me.hWnd, "Open", Def_Web_NetSCSI, vbNullString, vbNullString, 1
        Case 5  ' About Visual IVR WorkFlow Editor
            frmAbout.Show vbModal
    End Select
End Sub
'----------------------------------------  菜单栏函数结束     ------------------------------------------------------
'*******************************************************************************************************************

Private Sub picToolBarHolder_Resize()
    ToolBar.Width = picToolBarHolder.Width
    ToolBar.Height = picToolBarHolder.Height
    
    '' Sun added to refresh menu bar property when resize the toolbar
    ToolBar.MenusMax = ToolBar.MenusMax
    ToolBar.Refresh

End Sub

'****************************  工具栏函数 *********************************************************
'Michael : Tool Bar action
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "新建"
            Call mnuFile_Click(0)
        Case "打开"
            Call mnuFile_Click(1)
        Case "保存"
            Call mnuFile_Click(2)
        Case "打印"
            Call mnuFile_Click(6)
        Case "剪切"
        Case "复制"
            Call mnuEdit_Click(0)
        Case "粘贴"
            Call mnuEdit_Click(1)
        
        '' Sun added 2002-06-11
        Case "撤销"
            If mnuEdit(13).Enabled Then
                Call mnuEdit_Click(13)
            End If
        Case "重复"
            If mnuEdit(14).Enabled Then
                Call mnuEdit_Click(14)
            End If
            
        Case "删除节点"
            Call mnuEdit_Click(8)
        Case "节点属性"
            Call mnuEdit_Click(9)
        Case "显示节点描述"
            Call mnuEdit_Click(10)
        Case "节点列表"
            Call mnuView_Click(4)
        Case "变量使用"
            Call mnuEdit_Click(5)
        Case "左对齐"
        Case "置中"
        Case "右对齐"
        
        '-------------------------------------------
        ' Sun added 2004-12-30
        Case "节点连线"
            ToolBar_MenuItemDbClick 1, 8
        '-------------------------------------------
        
        Case "选项"
            Call mnuView_Click(6)
        Case "拨放语音"
            If gintSoundResourceID > 0 Then
                Call F_PlayVoxFile(gintSoundResourceID)
            End If
        Case "停止放音"
            StopSound
    End Select

End Sub
'**************************    工具栏函数结束 ********************************************

Private Sub ToolBar_MenuItemDbClick(MenuNumber As Long, MenuItem As Long)
   
    ' Sun added 2002-04-02
    If ToolBar.Enabled = False Then
        Exit Sub
    End If
   
    If gCallFlow.CallFlowID = 0 Then
        
        Message ("M131")
        Exit Sub
   
    End If
      
    gCallFlow.AddUserNode
    
    If gCallFlow.CreateNode(gCallFlow.NewNodeID) Then
        
        gCallFlow.Node(gCallFlow.NewNodeID).FlowID = gCallFlow.CallFlowID
        ToolBar.MenuItemCur = MenuItem
        gCallFlow.Node(gCallFlow.NewNodeID).NodeNo = Val(ToolBar.MenuItemKey)
        F_NodeDefaultInfo gCallFlow.NewNodeID, gCallFlow.Node(gCallFlow.NewNodeID).NodeNo
        gCallFlow.Node(gCallFlow.NewNodeID).InPage = gCallFlow.CurrentPage
        
        gCallFlow.Node(gCallFlow.NewNodeID).Left = 600 - m_frmCallFlow.WorkPage.Left
        gCallFlow.Node(gCallFlow.NewNodeID).Top = 600 - m_frmCallFlow.WorkPage.Top
        If gCallFlow.NewNodeID - 1 > 0 Then
            If gCallFlow.Node(gCallFlow.NewNodeID - 1).Left = gCallFlow.Node(gCallFlow.NewNodeID).Left And _
               gCallFlow.Node(gCallFlow.NewNodeID - 1).Top = gCallFlow.Node(gCallFlow.NewNodeID).Top Then
               gCallFlow.Node(gCallFlow.NewNodeID).Left = gCallFlow.Node(gCallFlow.NewNodeID).Left + 500
               gCallFlow.Node(gCallFlow.NewNodeID).Top = gCallFlow.Node(gCallFlow.NewNodeID).Left + 500
            End If
        End If
        
        gCallFlow.Node(gCallFlow.NewNodeID).Height = Screen.TwipsPerPixelY * 32          '' 480
        gCallFlow.Node(gCallFlow.NewNodeID).Width = Screen.TwipsPerPixelX * 32           '' 480

        If gCallFlow.Node(gCallFlow.NewNodeID).NodeNo = 255 Then
            ''Create a line control
            Call gCallFlow.Node(gCallFlow.NewNodeID).AddLine
        End If
        gCallFlow.Node(gCallFlow.NewNodeID).Description = "N/A"
        gCallFlow.Node(gCallFlow.NewNodeID).Visible = True
        
        '' Sun added 2002-03-30
        gCallFlow.Node(gCallFlow.NewNodeID).NodeCaptionVisible = gblnShowNodeCaption
        
        '' Sun added 2008-01-18
        gCallFlow.Node(gCallFlow.NewNodeID).NodeTagVisible = gblnShowNodeTag
        
        ' Sun added 2002-03-20
        '' Add new node to DB
        gCallFlow.AddNewIvrRecord gCallFlow.NewNodeID
        
        '' Sun added 2005-10-25
        Dim lv_mousemove_pt As POINTAPI
        lv_mousemove_pt.x = gCallFlow.Node(gCallFlow.NewNodeID).Left / Screen.TwipsPerPixelX
        lv_mousemove_pt.y = gCallFlow.Node(gCallFlow.NewNodeID).Top / Screen.TwipsPerPixelY
        ClientToScreen m_frmCallFlow.WorkPage.hWnd, lv_mousemove_pt
        SetCursorPos lv_mousemove_pt.x, lv_mousemove_pt.y
        gCallFlow.Node(gCallFlow.NewNodeID).SelectThisNode
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack DEF_OPERATION_NEW

    Else
        ' let gCallFlow.NewNodeID decrease 1 because of the error
        gCallFlow.NewNodeID = gCallFlow.NewNodeID - 1
    End If
    
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        'change the splitter colour
        Splitter.BackColor = &H808080
        
        'set the current position to x
        currSplitPosX = CLng(x)
    Else
        'not the left button, so... if the current position <> default, cause a mouseup
        If currSplitPosX <> &H7FFFFFFF Then Splitter_MouseUp Button, Shift, x, y
        
        'set the current position to the default value
        currSplitPosX = &H7FFFFFFF
    End If
End Sub

Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'if the splitter has been moved...
    If currSplitPosX& <> &H7FFFFFFF Then
        'if the current position <> default, reposition the splitter and set this as the current value
        Dim lv_nDeltaWidth As Long
        lv_nDeltaWidth = x - currSplitPosX
        If Abs(lv_nDeltaWidth) > 60 And picToolBarHolder.Width + lv_nDeltaWidth > 600 And picToolBarHolder.Width + lv_nDeltaWidth < (Me.ScaleWidth - 100) Then
            picToolBarHolder.Width = picToolBarHolder.Width + lv_nDeltaWidth
            currSplitPosX = CLng(x)
        End If
    End If
End Sub

Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    'if the splitter has been moved...
    If currSplitPosX <> &H7FFFFFFF Then
        'if the current postition <> the last position do a final move of the splitter
        'if the current position <> default, reposition the splitter and set this as the current value
        Dim lv_nDeltaWidth As Long
        lv_nDeltaWidth = x - currSplitPosX
        If lv_nDeltaWidth <> 0 And picToolBarHolder.Width + lv_nDeltaWidth > 600 And picToolBarHolder.Width + lv_nDeltaWidth < (Me.ScaleWidth - 100) Then
            picToolBarHolder.Width = picToolBarHolder.Width + lv_nDeltaWidth
        End If
    
        'call this the default position
        currSplitPosX = &H7FFFFFFF
        
        'restore the normal splitter colour
        Splitter.BackColor = &H8000000F
    
    End If
End Sub

Public Sub CheckEditPopMenu(bInNode As Boolean)

    If bInNode Then   ' Mouse in a node
        mnuEdit(1).Enabled = False                      'paste
        mnuEdit(7).Enabled = True                       'Set Root
        mnuEdit(9).Enabled = True                       'property
    Else              ' Mouse in work page
        mnuEdit(1).Enabled = gClipBoard.ClipboardMark   'paste
        mnuEdit(7).Enabled = False                      'Set Root
        mnuEdit(9).Enabled = False                      'property
    End If
    
    mnuEdit(0).Enabled = (gCallFlow.SelectedCount > 0)  'copy
    mnuEdit(8).Enabled = (gCallFlow.SelectedCount > 0)  'delete

End Sub

Public Sub Shell_MenuItem_Delete()
    mnuEdit_Click 8
End Sub

Public Sub SetWorkFramesBackColor()
    m_frmCallFlow.WorkFrame.BackColor = gFrameBackColor
End Sub

Public Sub SetWorkPagesSize()
    m_frmCallFlow.WorkPage.Move 0, 0, gPageWidth, gPageHeight
End Sub

Public Sub SetWorkPagesBackColor()
    m_frmCallFlow.WorkPage.BackColor = gPageBackColor
End Sub

Public Sub SetDragHanldeColor()
    Dim i
    For i = 0 To 7
        m_frmCallFlow.picHandle(i).BackColor = gNodeHandColor
    Next
End Sub

Public Sub SetActiveFormItemsEnable(ByVal f_Enabled As Boolean)
    
    m_frmCallFlow.WorkFrame.Enabled = f_Enabled
    ToolBar.Enabled = f_Enabled
    m_frmCallFlow.VScroll.Enabled = f_Enabled
    m_frmCallFlow.HScroll.Enabled = f_Enabled

End Sub

Public Function IsCallFlowOpened(ByVal f_nID As Byte) As Boolean

    '' 判断是否存在
    If m_lstCFForms(f_nID) Is Nothing Then
        IsCallFlowOpened = False
    Else
        IsCallFlowOpened = True
    End If
    
End Function

Public Function CreateNewMDIForm(ByVal f_nID As Byte) As Boolean

    Set m_lstCFForms(f_nID) = New CFlowWorks
    CreateNewMDIForm = True
    
    '' 显示窗口
    m_lstCFForms(f_nID).Show
    
End Function

Public Sub AssignMDIForm(ByVal f_nID As Byte)
    Set m_lstCFForms(f_nID) = m_frmCallFlow
    Debug.Print "AssignMDIForm Form: " & f_nID

    '' Sun added 2007-03-25
    If f_nID > 0 Then
        Me.mnuFile(2).Enabled = True
        Me.mnuFile(3).Enabled = True
        Me.mnuFile(5).Enabled = True
        Me.mnuFile(6).Enabled = True
        Me.mnuFile(10).Enabled = True
    End If
End Sub

Public Sub DeassignMDIForm(ByVal f_nID As Byte)
    Set m_lstCFForms(f_nID) = Nothing
    Debug.Print "DeassignMDIForm Form: " & f_nID
    
    '' Sun added 2007-03-25
    UpdateFileMenuStatus False
    
End Sub

Public Sub SetActiveMDIForm(ByVal f_Form As CFlowWorks)
    Set m_frmCallFlow = f_Form
    
    '' Sun added 2006-04-21
    UpdateSystemStatusBar
    
    '' Sun added 2007-03-25
    UpdateFileMenuStatus gCallFlow.CallFlowID > 0
    
    '' Michael added 2008-8-28
    UpdateSystemNode
    
    Debug.Print "Active Form: " & m_frmCallFlow.Caption
End Sub

Public Sub SwitchToMDIForm(ByVal f_nID As Byte)
    If IsCallFlowOpened(f_nID) Then
        m_lstCFForms(f_nID).Show
        m_lstCFForms(f_nID).ZOrder
        m_lstCFForms(f_nID).SetFocus
        m_lstCFForms(f_nID).SetFormActive
    End If
End Sub

' 切换所有流程窗口的系统节点显示状态
'
Public Sub ShowAllCallFlowWindowSystemNodes()
On Error Resume Next

    Dim objForm As Form
    Dim lv_frmFlow As CFlowWorks
    
    For Each objForm In Forms
   
        If objForm.Tag = "1662" Then
            Set lv_frmFlow = objForm
            If Not lv_frmFlow.m_objCallFlow Is Nothing Then
                ' 设置系统节点的显示状态
                lv_frmFlow.m_objCallFlow.SetAllSysNodesVisible
            End If
        End If
    Next
    
End Sub

' 重新流程中打开指定资源编号的资源清单
'
Public Function ReopenResourceList(ByVal f_bytPID As Byte) As Byte
On Error Resume Next

    Dim objForm As Form
    Dim lv_frmFlow As CFlowWorks
    
    ReopenResourceList = 0
    
    For Each objForm In Forms
   
        If objForm.Tag = "1662" Then
            Set lv_frmFlow = objForm
            If Not lv_frmFlow.m_objCallFlow Is Nothing Then
                If lv_frmFlow.m_objCallFlow.ResourceID = f_bytPID Then
                    lv_frmFlow.m_objCallFlow.OpenResourceRecordSet f_bytPID, lv_frmFlow.m_objCallFlow.LanguageID
                    ReopenResourceList = lv_frmFlow.m_objCallFlow.CallFlowID
                End If
            End If
        End If
    Next
    
End Function

Public Sub ShowCallFlowOnScreen()

    Debug.Print "step 1"
    
    ' 显示流程
    ' Mike modified @ 2008-1-30
    Call gCallFlow.ShowCallFlowOnScreen(Me, m_frmCallFlow, m_frmCallFlow.WorkPage, gblnShowNodeCaption, gblnShowNodeTag, gSystem.intShowSysNodes > 0)

    '' Sun added 2012-05-07
    Call PushMRUItem(gCallFlow.CallFlowID, gCallFlow.CallFlowName)
    
    ' 更新状态栏
    UpdateSystemStatusBar

    UpdateFileMenuStatus True
    
    '' Sun added 2012-05-07
    Call LoadMRUMenu
    
End Sub

' 更新状态栏
' Mike Modified the Scrope of this sub to "Public" @ 2008-4-25
Public Sub UpdateSystemStatusBar()

    Me.StatusBar.Panels("Page").Text = LoadNationalResString(1554) & Trim(Str(gCallFlow.CurrentPage)) & _
            LoadNationalResString(1132) & Trim(Str(gCallFlow.PageCount)) & LoadNationalResString(1133)
    m_frmCallFlow.Caption = LoadNationalResString(1374) & gCallFlow.CallFlowName & "(" & Trim(Str(gCallFlow.CallFlowID)) & ")"
    Me.StatusBar.Panels("Resource").Text = LoadNationalResString(1070) & Trim(Str(gCallFlow.ResourceID))
    Me.StatusBar.Panels("Resource").ToolTipText = gCallFlow.ResourceName
    Me.StatusBar.Panels("Language").Text = LoadNationalResString(1071) & Trim(Str(gCallFlow.LanguageID))
    Me.StatusBar.Panels("Root").Text = LoadNationalResString(1072) & Trim(Str(gCallFlow.RootNodedID))

End Sub

'' 更新文件菜单状态
'
Private Sub UpdateFileMenuStatus(ByVal blnEnabled As Boolean)
    Me.mnuFile(2).Enabled = blnEnabled
    Me.mnuFile(3).Enabled = blnEnabled
    Me.mnuFile(5).Enabled = blnEnabled
    Me.mnuFile(6).Enabled = blnEnabled
    Me.mnuFile(10).Enabled = blnEnabled
    Me.mnuFile(11).Enabled = blnEnabled
    
    Me.mnuResource(2).Enabled = blnEnabled
    Me.mnuResource(5).Enabled = blnEnabled
    Me.mnuResource(6).Enabled = blnEnabled
    Me.mnuResource(8).Enabled = blnEnabled
    
    Me.mnuEdit(3).Enabled = blnEnabled
    Me.mnuEdit(4).Enabled = blnEnabled
    Me.mnuEdit(5).Enabled = blnEnabled

End Sub

'' Mike added 2008-8-28
'' 切换流程时更新相关系统节点
Private Sub UpdateSystemNode()
    Call gCallFlow.UpdateSystemNode
End Sub
