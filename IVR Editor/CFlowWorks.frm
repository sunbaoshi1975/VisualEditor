VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{19C577A9-275A-451D-A97E-CFEF5120930D}#1.0#0"; "voxplayer_v2.ocx"
Begin VB.Form CFlowWorks 
   Caption         =   "新流程"
   ClientHeight    =   6810
   ClientLeft      =   1035
   ClientTop       =   930
   ClientWidth     =   9660
   Icon            =   "CFlowWorks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleMode       =   0  'User
   ScaleWidth      =   9660
   Tag             =   "1662"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox WorkFrame 
      BackColor       =   &H8000000C&
      Height          =   6585
      Left            =   0
      ScaleHeight     =   6525
      ScaleWidth      =   9315
      TabIndex        =   0
      Tag             =   "1662"
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton cmdPage 
         Height          =   375
         Index           =   1
         Left            =   8760
         Picture         =   "CFlowWorks.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "1961"
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton cmdPage 
         Height          =   375
         Index           =   0
         Left            =   8760
         Picture         =   "CFlowWorks.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "1960"
         Top             =   3840
         Width           =   375
      End
      Begin VB.PictureBox WorkPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6045
         Left            =   0
         ScaleHeight     =   6015
         ScaleWidth      =   8625
         TabIndex        =   5
         Top             =   0
         Width           =   8655
         Begin MSComctlLib.ImageList imgListNodes 
            Left            =   6600
            Top             =   4920
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   36
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":0AC6
                  Key             =   "node000"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":1918
                  Key             =   "node001"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":276A
                  Key             =   "node002"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":35BC
                  Key             =   "node006"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":480E
                  Key             =   "node007"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":54E8
                  Key             =   "node008"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":61C2
                  Key             =   "node009"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":6E9C
                  Key             =   "node010"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":71B6
                  Key             =   "node016"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":8008
                  Key             =   "node017"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":88E2
                  Key             =   "node018"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":95BC
                  Key             =   "node019"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":A296
                  Key             =   "node020"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":AF70
                  Key             =   "node021"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":BC4A
                  Key             =   "node022"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":C924
                  Key             =   "node023"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":D5FE
                  Key             =   "node028"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":F308
                  Key             =   "node040"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":FFE2
                  Key             =   "node041"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":10CBC
                  Key             =   "node050"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":11996
                  Key             =   "node051"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":12670
                  Key             =   "node055"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":1334A
                  Key             =   "node060"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":14024
                  Key             =   "node061"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":14CFE
                  Key             =   "node062"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":159D8
                  Key             =   "node063"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":166B2
                  Key             =   "node069"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":1738C
                  Key             =   "node070"
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":176A6
                  Key             =   "node071"
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":179C0
                  Key             =   "node090"
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":1869A
                  Key             =   "node091"
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":189B4
                  Key             =   "node096"
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":18CCE
                  Key             =   "node100"
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":199A8
                  Key             =   "node101"
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":1A682
                  Key             =   "node102"
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CFlowWorks.frx":1AAD4
                  Key             =   "node255"
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox picHandle 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   6870
            ScaleHeight     =   315
            ScaleWidth      =   345
            TabIndex        =   6
            Top             =   420
            Visible         =   0   'False
            Width           =   375
         End
         Begin PicClip.PictureClip clpArrows 
            Left            =   780
            Top             =   5340
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Rows            =   4
            Cols            =   4
            Picture         =   "CFlowWorks.frx":1ABAE
         End
         Begin VB.Image imgHyperLink 
            Height          =   480
            Left            =   5640
            Picture         =   "CFlowWorks.frx":1AE40
            Top             =   4800
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image imgDefaNode 
            Height          =   360
            Left            =   5160
            Picture         =   "CFlowWorks.frx":1AF92
            Top             =   4920
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Image imgSelect 
            Height          =   105
            Left            =   4920
            Picture         =   "CFlowWorks.frx":1B134
            Top             =   4920
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Line lnHMargin 
            BorderStyle     =   5  'Dash-Dot-Dot
            Visible         =   0   'False
            X1              =   3870
            X2              =   3870
            Y1              =   2160
            Y2              =   5760
         End
         Begin VB.Line lnVMargin 
            BorderStyle     =   5  'Dash-Dot-Dot
            Index           =   0
            Visible         =   0   'False
            X1              =   1395
            X2              =   6285
            Y1              =   3555
            Y2              =   3555
         End
         Begin VB.Shape shpSelectRegion 
            BorderStyle     =   3  'Dot
            Height          =   435
            Left            =   2850
            Top             =   1530
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin VB.HScrollBar HScroll 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   6240
         Width           =   9045
      End
      Begin VB.VScrollBar VScroll 
         Height          =   6260
         Left            =   9030
         TabIndex        =   3
         Top             =   4920
         Width           =   285
      End
      Begin VB.PictureBox Sizer 
         BorderStyle     =   0  'None
         Height          =   235
         Left            =   9060
         ScaleHeight     =   240
         ScaleWidth      =   225
         TabIndex        =   2
         Top             =   6255
         Width           =   225
      End
      Begin VOXPLAYERLib.VOXPlayer VOXPlayer 
         Height          =   465
         Left            =   1305
         TabIndex        =   1
         Top             =   5445
         Visible         =   0   'False
         Width           =   690
         _Version        =   65536
         _ExtentX        =   1217
         _ExtentY        =   820
         _StockProps     =   0
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog 
         Left            =   6240
         Top             =   420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6900
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B212
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B324
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B436
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B548
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B65A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B76C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B87E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1B990
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1BAA2
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1BBB4
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1BCC6
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1BDD8
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1BEEA
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1BFFC
            Key             =   "Description"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1C44E
            Key             =   "Dustbin"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1EC00
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1ED5A
            Key             =   "Label"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1EEB4
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1F00E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":1F328
            Key             =   "Tools"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":20AEA
            Key             =   "Viewer"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":20E04
            Key             =   "UnDo"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":20F5E
            Key             =   "ReDo"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":210B8
            Key             =   "NodeList"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFlowWorks.frx":21452
            Key             =   "NodeLine"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CFlowWorks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New CRect
Private m_DragPoint As POINTAPI

Private sysXscroll As Integer, sysYscroll As Integer, vbsysXscroll As Integer, vbsysYscroll As Integer
Private rc As RECT, rc1 As RECT

Private pnt As PaintEffects

Private MyFont As Long, OldFont As Long
Private Caps(0 To 38) As String

' Global flow and node information variable
Public m_objCallFlow As clsIVRProgram

' Last Mouse Click Point
Private m_ptMouseClick As POINTAPI

Public Sub ClearFormContent()
    '清屏
    m_objCallFlow.DestroyAllNodes
    m_objCallFlow.ClearWorkPage
End Sub

Public Sub SetFormActive()
    Set gCallFlow = m_objCallFlow
    frmMain.SetActiveMDIForm Me
End Sub

'Michael Added @ 2007-11-27
'页滚动按钮
Private Sub cmdPage_Click(Index As Integer)
'    Dim lv_ScrollStep As Integer    ' 翻页步长
'    If gCallFlow.PageCount > 1 Then
'        lv_ScrollStep = Int(VScroll.Max / (gCallFlow.PageCount - 1))
'    Else
'        lv_ScrollStep = VScroll.Max
'    End If
    
    '上一页
    If Index = 0 Then
        If gCallFlow.CurrentPage <> 1 Then
            Mdlfunction.GotoAnotherPage gCallFlow.CurrentPage - 1
            '移动滚动条
            'VScroll.value = lv_ScrollStep * (gCallFlow.CurrentPage - 1)
        End If
    '下一页
    ElseIf Index = 1 Then
        If gCallFlow.CurrentPage <> gCallFlow.PageCount Then
            Mdlfunction.GotoAnotherPage gCallFlow.CurrentPage + 1
            '移动滚动条
            'VScroll.value = lv_ScrollStep * (gCallFlow.CurrentPage - 1)
        End If
    End If
    'Michael Commented @2007-12-5
    'Call CheckStatus
End Sub

Private Sub Form_Activate()
    '' Sun added 2002-03-29
    ShowMarginLine
    
    SetFormActive
    
    'Michael Added @ 2007-11-28
    'Michael Commented @2007-12-5
    'Call CheckStatus
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown
    
    ' Sun added 2002-04-02
    If Me.WorkFrame.Enabled = False Then Exit Sub
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0

    Select Case KeyCode
    Case vbKeyDelete              '' Del
        frmMain.Shell_MenuItem_Delete
            
    Case vbKeyLeft
        If AltDown Or CtrlDown Then
            MoveSelectedNoeds 1, AltDown
        End If
        
    Case vbKeyUp
        If AltDown Or CtrlDown Then
            MoveSelectedNoeds 2, AltDown
        End If

    Case vbKeyRight
        If AltDown Or CtrlDown Then
            MoveSelectedNoeds 3, AltDown
        End If
    
    Case vbKeyDown
        If AltDown Or CtrlDown Then
            MoveSelectedNoeds 4, AltDown
        End If

    Case vbKeyA                    '' Select All Nodes in Current Page
        If CtrlDown Then
            m_objCallFlow.SetPageNodeSelect m_objCallFlow.CurrentPage, True
        End If
    
    End Select

End Sub

Private Sub Form_Load()
    
    ' Initialize drag process
    DragInit
    
    InitNewFlow
    
    ' Initialize Work frame and page
    WorkFrame.BackColor = gFrameBackColor
    WorkPage.BackColor = gPageBackColor
    WorkPage.Move 0, 0, gPageWidth, gPageHeight
    
    ' initialize HScroll and VScroll properties
    sysXscroll = GetSystemMetrics(SM_CXHSCROLL)
    sysYscroll = GetSystemMetrics(SM_CYHSCROLL)
    vbsysXscroll = sysXscroll * Screen.TwipsPerPixelX
    vbsysYscroll = sysYscroll * Screen.TwipsPerPixelY
    HScroll.Max = 100
    HScroll.LargeChange = 20
    HScroll.SmallChange = 5
    VScroll.Max = 1000
    VScroll.LargeChange = 200
    VScroll.SmallChange = 50
    VScroll.ZOrder
    HScroll.ZOrder
    
    cmdPage(0).Width = vbsysXscroll
    cmdPage(1).Width = vbsysXscroll
    cmdPage(0).Height = vbsysYscroll
    cmdPage(1).Height = vbsysYscroll
   
    '' Sun added 2001-10-07
    clpArrows.ClipHeight = 8
    clpArrows.ClipWidth = 8
    
    '' Sun added 2002-04-02
    gintSoundResourceID = 0
    Call SoundResourceIDChanged
    
    'Mike added @2008-8-29 for Support Mouse wheel
    Hook Me.hWnd
    
    '上句中的中文是以 buttons中的key出现的 所以不需要考虑
    LoadResStrings Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

    If Not m_objCallFlow.SavedMark Then
    
        Dim msgresult As VbMsgBoxResult
        
        msgresult = MsgBox(LoadNationalResString(1102) + Format(Str(m_objCallFlow.CallFlowID)) + "  ?", vbYesNoCancel + vbApplicationModal + vbQuestion)
        If msgresult = vbYes Then
            ' Add storing flow data to disk/database procedure code here...
            m_objCallFlow.UpdateIvrTable
        End If
        If msgresult = vbCancel Then Cancel = True
    End If

End Sub

Private Sub Form_Resize()
    
    WorkFrame.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim lv_cfID As Integer, lv_cfname As String
    lv_cfID = m_objCallFlow.CallFlowID
    lv_cfname = m_objCallFlow.CallFlowName
    
    frmMain.DeassignMDIForm m_objCallFlow.CallFlowID
    
    '清屏
    ClearFormContent
   
    Set pnt = Nothing
    
    'Destroy the font object created in
    'the form's window procedure.
    Call DeleteObject(MyFont)
    
    'Mike added @2008-8-29 for Support Mouse wheel
    UnHook Me.hWnd
    
    'Mike added @ 2008-7-8
    Call WriteLogMessage(0, enu_Information, "Callflow Closed!", "Callflow ID:" & lv_cfID & " Callflow Name:" & lv_cfname)
End Sub

'Tony Modified @ 2007-11-27
'添加滚动按钮
Private Sub WorkFrame_Resize()
    
    'VScroll.Move IIf(WorkFrame.ScaleWidth > vbsysXscroll, WorkFrame.ScaleWidth - vbsysXscroll, 0), 0, vbsysXscroll, IIf(WorkFrame.ScaleHeight > vbsysYscroll, WorkFrame.ScaleHeight - vbsysYscroll, 0)
    
    Dim lv_nLeft As Integer
    Dim lv_nBottom As Integer
    
'    cmdPage(0).Picture = LoadPicture(App.path & "..\bmp\arr_Up.BMP")
'    cmdPage(1).Picture = LoadPicture(App.path & "..\bmp\arr_Down.BMP")
    
    lv_nLeft = IIf(WorkFrame.ScaleWidth > vbsysXscroll, WorkFrame.ScaleWidth - vbsysXscroll, 0)
    lv_nBottom = IIf(WorkFrame.ScaleHeight > 3 * vbsysYscroll, WorkFrame.ScaleHeight - 3 * vbsysYscroll, 0)
    cmdPage(0).Left = lv_nLeft
    cmdPage(0).Top = 0
    cmdPage(1).Left = lv_nLeft
    cmdPage(1).Top = lv_nBottom + vbsysYscroll
    VScroll.Move lv_nLeft, vbsysYscroll, vbsysXscroll, lv_nBottom
    'Debug.Print gCallFlow.CurrentPage, gCallFlow.PageCount
    HScroll.Move 0, WorkFrame.ScaleHeight - vbsysYscroll, IIf(WorkFrame.ScaleWidth > vbsysXscroll, WorkFrame.ScaleWidth - vbsysXscroll, vbsysXscroll), vbsysYscroll
    Sizer.Move WorkFrame.ScaleWidth - vbsysXscroll, WorkFrame.ScaleHeight - vbsysYscroll, vbsysXscroll, vbsysYscroll
End Sub

Private Sub HScroll_Change()
    WorkPage.Left = (-HScroll.value / 100) * (Abs(WorkFrame.ScaleWidth - WorkPage.ScaleWidth) + vbsysXscroll + 180) ' + PageLeftMargin
End Sub

Private Sub HScroll_Scroll()
    WorkPage.Left = (-HScroll.value / 100) * (Abs(WorkFrame.ScaleWidth - WorkPage.ScaleWidth) + vbsysXscroll + 180) '+ PageLeftMargin
End Sub

Private Sub VScroll_Change()
    Dim lv_nPageHeight As Integer
    Dim lv_nPageStart As Integer
    Dim lv_nPageNo As Integer
    
    If gCallFlow.PageCount > 0 Then
        
        ''Michael Modified @ 2007-12-5
        '' Page No
        'lv_nPageHeight = Int((VScroll.Max - VScroll.Min + 1) / gCallFlow.PageCount + 0.5)
        lv_nPageHeight = Int((VScroll.Max - VScroll.Min + 1) / 0.5)
        
        If lv_nPageHeight > 0 Then
        
            lv_nPageNo = Int((VScroll.value - VScroll.Min) / lv_nPageHeight) + 1
            lv_nPageStart = (lv_nPageNo - 1) * lv_nPageHeight - 3
            
             'Michael Commented @ 2007-12-5
'            If lv_nPageNo <> gCallFlow.CurrentPage Then
'                Mdlfunction.GotoAnotherPage lv_nPageNo
'            End If
            
            'Michael Modified @ 2007-12-5
            WorkPage.Top = (-VScroll.value / VScroll.Max) * (Abs(WorkFrame.ScaleHeight - WorkPage.ScaleHeight) + vbsysYscroll + 120) ' + PageTopMargin
            'WorkPage.Top = (-(VScroll.value - lv_nPageStart) / lv_nPageHeight) * (Abs(WorkFrame.ScaleHeight - WorkPage.ScaleHeight) + vbsysYscroll + 120)
    
        End If
        
    End If
    
    'Michael Added @ 2007-11-8
    'Michael Commented @ 2007-12-5
    'Call CheckStatus
End Sub

Private Sub VScroll_Scroll()
    VScroll_Change
End Sub

Private Sub UpdateCurrentPositionDisplay(ByVal f_nIndex As Integer)
    
    frmMain.StatusBar.Panels("keyPosition").Text = _
        Str(Int(m_objCallFlow.Node(f_nIndex).Left / Screen.TwipsPerPixelX)) _
        + " ," + Str(Int(m_objCallFlow.Node(f_nIndex).Top / Screen.TwipsPerPixelY))

End Sub

Private Sub MoveCurrentNode()
    Dim nWidth As Single, nHeight As Single
    Dim pt As POINTAPI

    'Save dimensions before modifying rectangle
    nWidth = m_DragRect.Right - m_DragRect.Left
    nHeight = m_DragRect.Bottom - m_DragRect.Top
    'Get current mouse position in screen coordinates
    GetCursorPos pt
    'Hide existing rectangle
    DrawDragRect
    'Update drag rectangle coordinates
    m_DragRect.Left = pt.x - m_DragPoint.x
    m_DragRect.Top = pt.y - m_DragPoint.y
    m_DragRect.Right = m_DragRect.Left + nWidth
    m_DragRect.Bottom = m_DragRect.Top + nHeight
    
    frmMain.StatusBar.Panels("keyPosition").Text = _
        Str(m_DragRect.Left) + " ," _
        + Str(m_DragRect.Top - WorkFrame.Top / Screen.TwipsPerPixelY - 10)
    
    'Draw new rectangle
    DrawDragRect
    
    'Mike Added @ 2008-7-7
    'Call WriteLogMessage(0, enu_Information, "Move Select node to a new location.")

End Sub

Private Sub ResizeCurrentNode()
    Dim nWidth As Single, nHeight As Single
    Dim pt As POINTAPI

    'Get current mouse position in screen coordinates
    GetCursorPos pt
    'Hide existing rectangle
    DrawDragRect
    'Action depends on handle being dragged
    Select Case m_DragHandle
        Case 0   'Top Left
            m_DragRect.Left = pt.x - WorkPage.Left / Screen.TwipsPerPixelX - 3
            m_DragRect.Top = pt.y - WorkPage.Top / Screen.TwipsPerPixelY
        Case 1   'Top center
            m_DragRect.Top = pt.y - WorkPage.Top / Screen.TwipsPerPixelY
        Case 2   'Top right
            m_DragRect.Right = pt.x - WorkPage.Left / Screen.TwipsPerPixelX - 3
            m_DragRect.Top = pt.y - WorkPage.Top / Screen.TwipsPerPixelY
        Case 3   'Center right
            m_DragRect.Right = pt.x - WorkPage.Left / Screen.TwipsPerPixelX - 3
        Case 4  'Bottom right
            m_DragRect.Right = pt.x - WorkPage.Left / Screen.TwipsPerPixelX - 3
            m_DragRect.Bottom = pt.y - WorkPage.Top / Screen.TwipsPerPixelY
        Case 5   'Bottom center
            m_DragRect.Bottom = pt.y - WorkPage.Top / Screen.TwipsPerPixelY
        Case 6   'Bottom left
            m_DragRect.Left = pt.x - WorkPage.Left / Screen.TwipsPerPixelX - 3
            m_DragRect.Bottom = pt.y - WorkPage.Top / Screen.TwipsPerPixelY
        Case 7  'Center left
            m_DragRect.Left = pt.x - WorkPage.Left / Screen.TwipsPerPixelX - 3
    End Select

    'Draw new rectangle
    DrawDragRect

End Sub

'To handle all mouse message anywhere on the form, we set the mouse
'capture to the form. Mouse movement is processed here
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Sun added 2002-04-02
    If Me.WorkFrame.Enabled = False Then Exit Sub
    
    If m_DragState = StateDragging Then
        MoveCurrentNode
    ElseIf m_DragState = StateSizing Then
        ResizeCurrentNode
    End If
        
    frmMain.StatusBar.Panels("keySize").Text = Str(m_DragRect.Right - m_DragRect.Left) + " x" + Str(m_DragRect.Bottom - m_DragRect.Top)
    
End Sub

'To handle all mouse message anywhere on the form, we set the mouse
'capture to the form. Mouse up is processed here
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Sun added 2002-04-02
    If Me.WorkFrame.Enabled = False Then Exit Sub
    
    If Button = vbLeftButton Then
        AutoDragEnd
    End If
End Sub

Public Sub AutoDragEnd()
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_loop As Integer
    Dim lv_Str As String
    Dim lv_nOldLeft As Long
    Dim lv_nOldTop As Long
    
    If m_DragState = StateDragging Or m_DragState = StateSizing Then
        
        'Hide drag rectangle
        DrawDragRect
        
        'Move control to new location
        gClipBoard.PushClipBoardStack
        lv_nOldLeft = m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Left
        lv_nOldTop = m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Top
        m_DragRect.ScreenToTwips m_CurrCtl
        m_DragRect.SetCtrlToRect m_CurrCtl
        m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Width = m_DragRect.Width
        m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Height = m_DragRect.Height
        
        If TypeOf m_CurrCtl Is Line Then
            
            '' Move Image with Line
            m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Line_X1 = m_CurrCtl.x1
            m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Line_Y1 = m_CurrCtl.y1
            m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Line_X2 = m_CurrCtl.x2
            m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Line_Y2 = m_CurrCtl.y2
            m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).MoveImageWithLine
            
        End If
        
        If m_DragState = StateSizing Or _
            Abs(lv_nOldLeft - m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Left) >= Move_Small_Step Or _
            Abs(lv_nOldTop - m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Top) >= Move_Small_Step Then
        
            'Sun added 2001-10-03
            If TypeOf m_CurrCtl Is Line Then
                
                '' Automatically connect line to node(s)
                Node255_Data2.StartNode = m_objCallFlow.FindNodeIDAtPoint(m_CurrCtl.x1, m_CurrCtl.y1)
                Node255_Data2.EndNode = m_objCallFlow.FindNodeIDAtPoint(m_CurrCtl.x2, m_CurrCtl.y2)
                
                '' Sun added 2002-04-01
                ''' For update the parent node property automatically
                If Node255_Data2.EndNode > 255 And Node255_Data2.StartNode > 255 Then
                    Call m_objCallFlow.SetNodeParentNote(Node255_Data2.EndNode, Node255_Data2.StartNode)
                End If
                                   
                '' Sun added 2002-04-05
                ''' For update the child node property automatically
                If Node255_Data2.StartNode > 255 And Node255_Data2.EndNode > 255 Then
                    lv_Str = m_objCallFlow.SetNodeChildNode(Node255_Data2.StartNode, Node255_Data2.EndNode, Node255_Data2.Index)
                    If lv_Str <> "" Then
                        m_objCallFlow.Node(m_objCallFlow.NodeSelectedID).Description = lv_Str
                    End If
                End If
                
                '' Sun added 2002-03-20
                F_NodeData m_objCallFlow.NodeSelectedID, 255

            Else
            
                MoveLinesOnNode m_objCallFlow.NodeSelectedID
                
            End If
        
            '' Sun added 2001-09-28
            ''' 保存节点位置数据
            m_objCallFlow.UpdateAnotherIVRRecord m_objCallFlow.NodeSelectedID
        
        End If
    
        'Restore sizing handles
        ShowHandles True
        'Free mouse movement
        ClipCursor ByVal 0&
        'Release mouse capture
        ReleaseCapture
        'Reset drag state
        m_DragState = StateNothing
        
    End If

End Sub

'Process MouseDown over handles
Private Sub picHandle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer

    'Handles should only be visible when a control is selected
    Debug.Assert (Not m_CurrCtl Is Nothing)
    'NOTE: m_DragPoint not used for sizing
    'Save control position in screen coordinates
    m_DragRect.SetRectToCtrl m_CurrCtl
    m_DragRect.TwipsToScreen m_CurrCtl
    'Track index handle
    m_DragHandle = Index
    'Hide sizing handles
    ShowHandles False
    'We need to force handles to hide themselves before drawing drag rectangle
    Refresh
    'Indicate sizing is under way
    m_DragState = StateSizing
    'Show sizing rectangle
    DrawDragRect
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    SetCapture hWnd
    'Limit cursor movement within form
    GetWindowRect hWnd, rc
    GetClientRect hWnd, rc1
    rc.Top = rc.Bottom - rc1.Bottom - rc1.Top
    ClipCursor rc
End Sub


'Because some lightweight controls do not have a MouseDown event,
'when we get a MouseDown event on a form, we do a scan of the
'Controls collection to see if any lightweight controls are under
'the mouse. Note that this code does not work for controls within
'containers. Also, if no control is under the mouse, then we
'remove the sizing handles and clear the current control.
Private Sub WorkPage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    If Button = vbLeftButton Then
        'Hit test over light-weight (non-windowed) controls
'        For i = 0 To (Controls.Count - 1)
'            'Check for visible, non-menu controls
'            '[Note 1]
'            'If any of the sizing handle controls are under the mouse
'            'pointer, then they must not be visible or else they would
'            'have already intercepted the MouseDown event.
'            '[Note 2]
'            'This code will fail if you have a control such as the
'            'Timer control which has no Visible property. You will
'            'either need to make sure your form has no such controls
'            'or add code to handle them.
'            If Not TypeOf Controls(i) Is Menu And Controls(i).Visible Then
'                'And Controls(i).Name <> "WorkFrame" And Controls(i).Name <> "WorkPage" And _
'                Controls(i).Name <> "HScroll" And Controls(i).Name <> "VScroll" And _
'                Controls(i).Name <> "Sizer" Then
'                m_DragRect.SetRectToCtrl Controls(i)
'                If m_DragRect.PtInRect(X, Y) Then
'                    DragBegin Controls(i)
'                    Exit Sub
'                End If
'            End If
'        Next i
        
        'No control is active
        Set m_CurrCtl = Nothing
        m_objCallFlow.NodeSelectedID = 0
        
        SetMouseClickPoint x, y
        shpSelectRegion.Width = 0
        shpSelectRegion.Height = 0
        shpSelectRegion.Visible = True
        
    End If
    
    If Button = vbRightButton Then PopupMenu frmMain.mEdit
End Sub

'========================== Dragging Code ================================

'Initialization -- Do not call more than once
Private Sub DragInit()
    Dim i As Integer, xHandle As Single, yHandle As Single

    'Use black Picture box controls for 8 sizing handles
    'Calculate size of each handle
    xHandle = 5 * Screen.TwipsPerPixelX
    yHandle = 5 * Screen.TwipsPerPixelY
    'Load array of handles until we have 8
    For i = 0 To 7
        If i <> 0 Then
            Load picHandle(i)
        End If
        picHandle(i).BackColor = gNodeHandColor
        picHandle(i).Width = xHandle
        picHandle(i).Height = yHandle
        'Must be in front of other controls
        picHandle(i).ZOrder
    Next i
    'Set mousepointers for each sizing handle
    picHandle(0).MousePointer = vbSizeNWSE
    picHandle(1).MousePointer = vbSizeNS
    picHandle(2).MousePointer = vbSizeNESW
    picHandle(3).MousePointer = vbSizeWE
    picHandle(4).MousePointer = vbSizeNWSE
    picHandle(5).MousePointer = vbSizeNS
    picHandle(6).MousePointer = vbSizeNESW
    picHandle(7).MousePointer = vbSizeWE
    'Initialize current control
    Set m_CurrCtl = Nothing
End Sub

'Drags the specified control
Public Sub DragBegin(ctl As Control)
    'Hide and visible handles
    ShowHandles False
    'Save reference to control being dragged
    Set m_CurrCtl = ctl
    'Store initial mouse position
    GetCursorPos m_DragPoint
    'Save control position (in screen coordinates)
    'Note: control might not have a window handle
    m_DragRect.SetRectToCtrl m_CurrCtl
    m_DragRect.TwipsToScreen m_CurrCtl
    'Make initial mouse position relative to control
    m_DragPoint.x = m_DragPoint.x - m_DragRect.Left
    m_DragPoint.y = m_DragPoint.y - m_DragRect.Top
    'Force redraw of form without sizing handles
    'before drawing dragging rectangle
    Refresh
    'Show dragging rectangle
    DrawDragRect
    'Indicate dragging under way
    m_DragState = StateDragging
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    ReleaseCapture  'This appears needed before calling SetCapture
    SetCapture hWnd
    'Limit cursor movement within form
    GetWindowRect hWnd, rc
    GetClientRect hWnd, rc1
    rc.Top = rc.Bottom - rc1.Bottom - rc1.Top
    ClipCursor rc
    'Modify:Scott Date:2001/08/31 Get Node's data
    
On Error Resume Next
       
       F_ExplainNodeData m_objCallFlow.NodeSelectedID

On Error GoTo 0

End Sub

'Clears any current drag mode and hides sizing handles
Private Sub DragEnd()
    'Hide and visible handles
    ShowHandles False
    Set m_CurrCtl = Nothing
    m_DragState = StateNothing
End Sub

' Drags all selected controls in current page
Public Sub DragBeginEx()
    
    'Store initial mouse position
    GetCursorPos m_DragPoint
    
    'Save control position (in screen coordinates)
    'Note: control might not have a window handle
    If Not gCallFlow.GetSelectItemRect(m_DragRect.Left, m_DragRect.Top, m_DragRect.Right, m_DragRect.Bottom) Then
        Exit Sub
    End If
    
    'Make initial mouse position relative to control
    m_DragPoint.x = m_DragPoint.x - m_DragRect.Left
    m_DragPoint.y = m_DragPoint.y - m_DragRect.Top
    
    'Force redraw of form without sizing handles
    'before drawing dragging rectangle
    Refresh
    'Show dragging rectangle
    DrawDragRect
    'Indicate dragging under way
    m_DragState = StateDragging
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    ReleaseCapture  'This appears needed before calling SetCapture
    SetCapture hWnd
    'Limit cursor movement within form
    GetWindowRect hWnd, rc
    GetClientRect hWnd, rc1
    rc.Top = rc.Bottom - rc1.Bottom - rc1.Top
    ClipCursor rc
    'Modify:Scott Date:2001/08/31 Get Node's data

End Sub

'Clears any current drag mode and hides sizing handles
Private Sub DragEndEx()
    m_DragState = StateNothing
End Sub

'Display or hide the sizing handles and arrange them for the current rectangld
Public Sub ShowHandles(Optional bShowHandles As Boolean = True)
    Dim i As Integer
    Dim xFudge As Long, yFudge As Long
    Dim nWidth As Long, nHeight As Long
    Dim lv_StartPT As Byte, lv_EndPT As Byte

    For i = 0 To 7
        picHandle(i).BackColor = &HFFFFFF               '' White
        picHandle(i).Visible = False
    Next i

    If bShowHandles And Not m_CurrCtl Is Nothing Then
    
        nWidth = (picHandle(0).Width \ 2)
        nHeight = (picHandle(0).Height \ 2)
        xFudge = (0.5 * Screen.TwipsPerPixelX)
        yFudge = (0.5 * Screen.TwipsPerPixelY)

        'Sun added 2001-10-03
        If TypeOf m_CurrCtl Is Line Then
        
            With m_DragRect
                
                ' Get the Start Point and End Point
                If m_CurrCtl.x2 - m_CurrCtl.x1 > 0 Then
                    If m_CurrCtl.y2 - m_CurrCtl.y1 > 0 Then
                        ' Top Left
                        lv_StartPT = 0
                        ' Bottom right
                        lv_EndPT = 4
                    Else
                        ' Bottom left
                        lv_StartPT = 6
                        ' Top right
                        lv_EndPT = 2
                    End If
                Else
                    If m_CurrCtl.y2 - m_CurrCtl.y1 > 0 Then
                        ' Top right
                        lv_StartPT = 2
                        ' Bottom left
                        lv_EndPT = 6
                    Else
                        ' Bottom right
                        lv_StartPT = 4
                        ' Top Left
                        lv_EndPT = 0
                    End If
                End If
                        
                
                'Top Left
                picHandle(0).Move (.Left - nWidth) + xFudge, (.Top - nHeight) + yFudge
                'Bottom right
                picHandle(4).Move (.Left + .Width) - nWidth - xFudge, .Top + .Height - nHeight - yFudge
                'Top center
                picHandle(1).Move .Left + (.Width / 2) - nWidth, .Top - nHeight + yFudge
                'Bottom center
                picHandle(5).Move .Left + (.Width / 2) - nWidth, .Top + .Height - nHeight - yFudge
                'Top right
                picHandle(2).Move .Left + .Width - nWidth - xFudge, .Top - nHeight + yFudge
                'Bottom left
                picHandle(6).Move .Left - nWidth + xFudge, .Top + .Height - nHeight - yFudge
                'Center right
                picHandle(3).Move .Left + .Width - nWidth - xFudge, .Top + (.Height / 2) - nHeight
                'Center left
                picHandle(7).Move .Left - nWidth + xFudge, .Top + (.Height / 2) - nHeight
                
                picHandle(lv_StartPT).Visible = True
                picHandle(lv_EndPT).Visible = True
                            
                If Node255_Data2.StartNode > 255 Then
                    picHandle(lv_StartPT).BackColor = &HFF00&                '' Green
                End If
                            
                If Node255_Data2.EndNode > 255 Then
                    picHandle(lv_EndPT).BackColor = &HFF00&                '' Green
                End If
                
            End With
            
        Else
        
            With m_DragRect
                'Top Left
                picHandle(0).Move (.Left - nWidth) + xFudge, (.Top - nHeight) + yFudge
                picHandle(0).Visible = True
                'Bottom right
                picHandle(4).Move (.Left + .Width) - nWidth - xFudge, .Top + .Height - nHeight - yFudge
                picHandle(4).Visible = True
                'Top center
                picHandle(1).Move .Left + (.Width / 2) - nWidth, .Top - nHeight + yFudge
                picHandle(1).Visible = True
                'Bottom center
                picHandle(5).Move .Left + (.Width / 2) - nWidth, .Top + .Height - nHeight - yFudge
                picHandle(5).Visible = True
                'Top right
                picHandle(2).Move .Left + .Width - nWidth - xFudge, .Top - nHeight + yFudge
                picHandle(2).Visible = True
                'Bottom left
                picHandle(6).Move .Left - nWidth + xFudge, .Top + .Height - nHeight - yFudge
                picHandle(6).Visible = True
                'Center right
                picHandle(3).Move .Left + .Width - nWidth - xFudge, .Top + (.Height / 2) - nHeight
                picHandle(3).Visible = True
                'Center left
                picHandle(7).Move .Left - nWidth + xFudge, .Top + (.Height / 2) - nHeight
                picHandle(7).Visible = True
            End With
        
        End If
        
    End If
    
End Sub

'Draw drag rectangle. The API is used for efficiency and also
'because drag rectangle must be drawn on the screen DC in
'order to appear on top of all controls
Private Sub DrawDragRect()
On Error Resume Next

    Dim hPen As Long, hOldPen As Long
    Dim hBrush As Long, hOldBrush As Long
    Dim hScreenDC As Long, nDrawMode As Long
    Dim lv_OldPoint As POINTAPI
    Dim lv_X1 As Long, lv_X2 As Long, lv_Y1 As Long, lv_Y2 As Long

    'Get DC of entire screen in order to
    'draw on top of all controls
    hScreenDC = GetDC(0)
    
    'Select GDI object
    hPen = CreatePen(PS_SOLID, 2, 0)
    hOldPen = SelectObject(hScreenDC, hPen)
    hBrush = GetStockObject(NULL_BRUSH)
    hOldBrush = SelectObject(hScreenDC, hBrush)
    nDrawMode = SetROP2(hScreenDC, R2_NOT)
    
    lv_X1 = m_DragRect.Left + WorkPage.Left / Screen.TwipsPerPixelX + 3
    lv_Y1 = m_DragRect.Top + WorkPage.Top / Screen.TwipsPerPixelY + 3
    lv_X2 = m_DragRect.Right + WorkPage.Left / Screen.TwipsPerPixelX + 3
    lv_Y2 = m_DragRect.Bottom + WorkPage.Top / Screen.TwipsPerPixelY + 3
    
    'Sun added 2002-04-03
    If TypeOf m_CurrCtl Is Line Then
        
        'Draw Line
        If Sgn(m_CurrCtl.x2 - m_CurrCtl.x1) = Sgn(m_CurrCtl.y2 - m_CurrCtl.y1) Then
        
            MoveToEx hScreenDC, lv_X1, lv_Y1, lv_OldPoint
            LineTo hScreenDC, lv_X2, lv_Y2
                
        Else
                        
            MoveToEx hScreenDC, lv_X2, lv_Y1, lv_OldPoint
            LineTo hScreenDC, lv_X1, lv_Y2

        End If
    
    Else
        
        'Draw rectangle
        Rectangle hScreenDC, lv_X1, lv_Y1, lv_X2, lv_Y2
    
    End If

    'Restore DC
    SetROP2 hScreenDC, nDrawMode
    SelectObject hScreenDC, hOldBrush
    SelectObject hScreenDC, hOldPen
    ReleaseDC 0, hScreenDC
    'Delete GDI objects
    DeleteObject hPen
End Sub

Public Sub ShowNodeProp(nNodeNo As Byte)
    Select Case nNodeNo
        Case 0   '全程转移规则
            frm_000.Show vbModal
        Case 1  'Buffer定义日志
            frm_001.Show vbModal
        Case 2  'Buffer定义变量
            frm_buffer.Show vbModal
        Case 6  '无条件转移
            Frm_006.Show 0, frmMain
        Case 7  '身份验证
            frm_007.Show 0, frmMain
        Case 8  '修改口令
            Frm_008.Show 0, frmMain
        Case 9  '时间分支
            frm_009.Show 0, frmMain
        Case 10  '工作日设定
            frm_010.Show 0, frmMain
        
        ''--------------------------------
        '' Sun added 2004-12-30
        Case 16  '条件分支
            frm_016.Show 0, frmMain
        ''--------------------------------
        
        Case 17 '选择服务语言
            frm_017.Show 0, frmMain
        Case 18  '发送数据
            frm_018.Show 0, frmMain
        Case 19  '无操作
            frm_019.Show 0, frmMain
        Case 255  '节点连线
            frm_Line.Show vbModal
        Case 20  '放音挂机
'            frm_020.Show 0, frmmain
            frm_020.Show 0, frmMain
        Case 21  '放音继续
            frm_021.Show 0, frmMain
        Case 22  '放音等待按键
            frm_022.Show 0, frmMain
        Case 23  '放音转移
            frm_023.Show 0, frmMain
            
        
        ''--------------------------------
        '' Sun added 2004-12-30
        Case 28  'TTS放音
            frm_028.Show 0, frmMain
        ''--------------------------------
        
        Case 40  '建立留言
            frm_040.Show 0, frmMain
        Case 41  '察看留言
            frm_041.Show 0, frmMain
        Case 50  '简单传真
            frm_050.Show 0, frmMain
        Case 51  'TTF传真
            frm_051.Show 0, frmMain
        
        ''--------------------------------
        '' Sun added 2006-12-31, V6.5.11
        Case 55  '传真接收
            frm_055.Show 0, frmMain
        ''--------------------------------
        
        Case 60  '转接座席
            frm_060.Show 0, frmMain
        Case 61  '转接座席组
            frm_061.Show 0, frmMain
        Case 62  '增强转接座席
            frm_062.Show 0, frmMain
        Case 63  '增强转接座席组
            frm_063.Show 0, frmMain
        Case 69  '转虚拟分机
            frm_069.Show 0, frmMain
        ''--------------------------------
        '' Sun added 2005-05-26
        Case 70  ' 查询路由点
            frm_070.Show 0, frmMain
        
        Case 71  ' 查询座席状态
            frm_071.Show 0, frmMain
        ''--------------------------------

        Case 90  '呼叫外线号码
            frm_090.Show 0, frmMain
        ''--------------------------------
        '' Sun added 2005-05-26
        Case 91  'Calling Card
            frm_091.Show 0, frmMain
        ''--------------------------------
        
        ''--------------------------------
        '' Sun added 2005-03-15
        Case 96  'TTS放音
            frm_096.Show 0, frmMain
        ''--------------------------------
        
        Case 100  '用户DLL
            frm_100.Show 0, frmMain
        Case 101  '用户COM
            frm_101.Show 0, frmMain
        Case 102  '102-记录变量
            frm_102.Show 0, frmMain
    End Select
End Sub

Private Sub InitNewFlow()
    
    'Create a buffer
    gSystem.strOSUser = String(100, Chr$(0))
    
    'Get the username
    GetUserName gSystem.strOSUser, 100
    
    'strip the rest of the buffer
    gSystem.strOSUser = Left$(gSystem.strOSUser, InStr(gSystem.strOSUser, Chr$(0)) - 1)
    
    ShowHandles False
    WorkPage.Cls
    
    Set m_objCallFlow = New clsIVRProgram
    Set gCallFlow = m_objCallFlow
    
    frmMain.mnuFile(2).Enabled = False
    frmMain.mnuFile(3).Enabled = False
    frmMain.mnuFile(5).Enabled = False
    frmMain.mnuFile(6).Enabled = False
    frmMain.mnuFile(10).Enabled = False
    
    ' Check Edit popup menu state
    frmMain.CheckEditPopMenu False
    
End Sub

' Print Work Page node by node
'
Public Sub PrintWorkPage(ByVal PageNo As Integer)
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    Dim lv_Index As Integer
    Dim lv_PageCount As Integer
    Dim lv_CurPage As Integer
    Dim lv_PageHeight As Long
    Dim lv_PageWidth As Long
    Dim lv_NodeLeft As Long
    Dim lv_NodeTop As Long
    Dim lv_FirstPageMargin As Integer
        
    '' Init Var
    lv_PageCount = m_objCallFlow.PageCount
    Call F_GetPrintPageScale(lv_PageWidth, lv_PageHeight)
    
    For lv_Index = 1 To m_objCallFlow.NewNodeID
        lv_CurPage = m_objCallFlow.Node(lv_Index).InPage
        If lv_CurPage > 0 And lv_CurPage <= lv_PageCount Then
        
            '' Sun added 2002-03-31
            If lv_CurPage = PageNo Or PageNo <= 0 Then
                
                '' Get Node Position
                lv_NodeLeft = m_objCallFlow.Node(lv_Index).Left
                'lv_NodeTop = m_objCallFlow.Node(lv_Index).Top + (lv_CurPage - 1) * lv_PageHeight
                lv_NodeTop = m_objCallFlow.Node(lv_Index).Top
                
                '' Skip title Region if first page
                If lv_CurPage = 1 Then
                    lv_FirstPageMargin = Def_TWIPS_PER_CM * 2
                Else
                    lv_FirstPageMargin = 0
                End If
                
                '' Sun added 2002-03-31 to avoid printing out of page
                If lv_NodeLeft <= lv_PageWidth And lv_NodeTop + lv_FirstPageMargin <= lv_PageHeight Then
                
                    '' Print Node
                    If m_objCallFlow.Node(lv_Index).NodeNo <> 255 Then
                        ''' General Node
                        Printer.PaintPicture m_objCallFlow.Node(lv_Index).Picture, lv_NodeLeft, lv_NodeTop + lv_FirstPageMargin
                    Else
                        ''' Line
                        Printer.Line (m_objCallFlow.Node(lv_Index).Line_X1, m_objCallFlow.Node(lv_Index).Line_Y1 + lv_FirstPageMargin)-(m_objCallFlow.Node(lv_Index).Line_X2, m_objCallFlow.Node(lv_Index).Line_Y2 + lv_FirstPageMargin)
                        ''' Arrow
                        Printer.PaintPicture m_objCallFlow.Node(lv_Index).Picture, lv_NodeLeft, lv_NodeTop + lv_FirstPageMargin
                    End If
                    
                    ''' Caption
                    If m_objCallFlow.Node(lv_Index).NodeCaptionVisible Then
                        Printer.CurrentX = m_objCallFlow.Node(lv_Index).NodeCaptionLeft
                        Printer.CurrentY = m_objCallFlow.Node(lv_Index).NodeCaptionTop + lv_FirstPageMargin
                        Printer.FontUnderline = True
                        Printer.Print Trim(m_objCallFlow.Node(lv_Index).NodeCaption)
                        Printer.FontUnderline = False
                    End If

                    '' Sun added 2008-01-18
                    ''' Tag
                    If m_objCallFlow.Node(lv_Index).NodeTagVisible Then
                        Printer.CurrentX = m_objCallFlow.Node(lv_Index).NodeTagLeft
                        Printer.CurrentY = m_objCallFlow.Node(lv_Index).NodeTagTop + lv_FirstPageMargin
                        Printer.FontUnderline = True
                        Printer.Print Trim(m_objCallFlow.Node(lv_Index).NodeTag)
                        Printer.FontUnderline = False
                    End If

                
                End If
                
            End If
            
        End If
    Next

    '' Sun added 2002-03-31
    ''' Print Page No.
    Printer.CurrentX = lv_PageWidth / 2 - Def_TWIPS_PER_CM * 2
    Printer.CurrentY = lv_PageHeight - Def_TWIPS_PER_CM * 2
    Printer.Print LoadNationalResString(1554) & Trim(Str(PageNo)) & LoadNationalResString(1132) & Trim(Str(m_objCallFlow.PageCount)) & LoadNationalResString(1133)

    '' Sun added 2002-03-29
    ShowMarginLine
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Sub

Public Sub CopyContentsOnWorkPage()
Dim lv_ptOld As POINTAPI
Dim lv_ptOld2 As POINTAPI
        
    gClipBoard.GetCopyedItem
    gClipBoard.CopyedItem.GetItemRect lv_ptOld.x, lv_ptOld.y, lv_ptOld2.x, lv_ptOld2.y
    SetMouseClickPoint lv_ptOld.x + 500, lv_ptOld.y + 500

End Sub

Public Sub PasteContentsOnWorkPage()
Dim lv_nItemNode As Byte
Dim lv_ptOld As POINTAPI
Dim lv_ptOld2 As POINTAPI
Dim lv_nNodeCount As Byte
Dim lv_nNewIndex() As Integer
Dim lv_nNewNodeID() As Integer
Dim lv_nOldNodeID() As Integer

If gClipBoard.MultiPushInitialize(gClipBoard.CopyedItem.NodeCount, DEF_OPERATION_NEW) Then
    
    gClipBoard.CopyedItem.GetItemRect lv_ptOld.x, lv_ptOld.y, lv_ptOld2.x, lv_ptOld2.y
''    m_objCallFlow.SetAllNodeSelect False
    
    '' Sun added 2006-02-06
    lv_nNodeCount = gClipBoard.CopyedItem.NodeCount
    ReDim lv_nNewIndex(lv_nNodeCount) As Integer
    ReDim lv_nNewNodeID(lv_nNodeCount) As Integer
    ReDim lv_nOldNodeID(lv_nNodeCount) As Integer
    
    For lv_nItemNode = 0 To lv_nNodeCount - 1
    
        m_objCallFlow.AddUserNode
        
        If m_objCallFlow.CreateNode(m_objCallFlow.NewNodeID) Then
        
            gClipBoard.RestoreNodeData gClipBoard.CopyedItem.Contents(lv_nItemNode), m_objCallFlow.Node(m_objCallFlow.NewNodeID), False, False
            
            '' Sun added 2006-02-06
            lv_nNewIndex(lv_nItemNode) = m_objCallFlow.NewNodeID
            lv_nNewNodeID(lv_nItemNode) = m_objCallFlow.UserNodeID
            lv_nOldNodeID(lv_nItemNode) = m_objCallFlow.Node(m_objCallFlow.NewNodeID).NodeID
            
            m_objCallFlow.Node(m_objCallFlow.NewNodeID).FlowID = m_objCallFlow.CallFlowID
            m_objCallFlow.Node(m_objCallFlow.NewNodeID).NodeID = m_objCallFlow.UserNodeID
            MovePasteNodeToRightPoint m_objCallFlow.Node(m_objCallFlow.NewNodeID), lv_ptOld.x, lv_ptOld.y
                        
            If m_objCallFlow.Node(m_objCallFlow.NewNodeID).NodeNo = 255 Then
                Call m_objCallFlow.Node(m_objCallFlow.NewNodeID).AddLine
                m_objCallFlow.Node(m_objCallFlow.NewNodeID).MoveImageWithLine
            End If
            m_objCallFlow.Node(m_objCallFlow.NewNodeID).InPage = m_objCallFlow.CurrentPage
            m_objCallFlow.Node(m_objCallFlow.NewNodeID).Visible = True
            m_objCallFlow.Node(m_objCallFlow.NewNodeID).NodeCaptionVisible = gblnShowNodeCaption
            
            '' Sun added 2008-01-18
            m_objCallFlow.Node(m_objCallFlow.NewNodeID).NodeTagVisible = gblnShowNodeTag
            
            m_objCallFlow.AddNewIvrRecord m_objCallFlow.NewNodeID
            
        Else
            lv_nNewIndex(lv_nItemNode) = 0
            lv_nNewNodeID(lv_nItemNode) = 0
            lv_nOldNodeID(lv_nItemNode) = 0
            m_objCallFlow.NewNodeID = m_objCallFlow.NewNodeID - 1
        End If
    
    Next
    
    SetMouseClickPoint m_ptMouseClick.x + 500, m_ptMouseClick.y + 500
    
    '' Sun added 2006-02-06
    ''' Automatically change node id
    For lv_nItemNode = 0 To lv_nNodeCount - 1
        If lv_nNewIndex(lv_nItemNode) > 0 Then
            m_objCallFlow.UpdateNodeIDProperties lv_nNewIndex(lv_nItemNode), lv_nNewNodeID(), lv_nOldNodeID()
            gClipBoard.MultiPushClipBoardStack lv_nNewIndex(lv_nItemNode)
            m_objCallFlow.Node(lv_nNewIndex(lv_nItemNode)).IsSelected = True
        End If
    Next
    
End If

End Sub

' Show Page Margin
'
Private Sub ShowMarginLine()
On Error Resume Next

    Dim lv_Index As Integer
    Dim lv_PageHeight As Long
    Dim lv_PageWidth As Long
    Dim lv_ShowPageCount As Integer
    
    If Not F_GetPrintPageScale(lv_PageWidth, lv_PageHeight) Then
        lnHMargin.Visible = False
        lnVMargin(0).Visible = False
        Exit Sub
    End If
    
    lnHMargin.x1 = lv_PageWidth
    lnHMargin.y1 = 1
    lnHMargin.x2 = lv_PageWidth
    lnHMargin.y2 = WorkPage.Height
    lnHMargin.Visible = True
    'Debug.Print "HLine: (" & Str(lnHMargin.X1) & ","; Str(lnHMargin.Y1) & "-(" & Str(lnHMargin.X2) & ","; Str(lnHMargin.Y2) & ")"
    
    '' Dynamic Add Lines
    lv_ShowPageCount = Int(WorkPage.Height / lv_PageHeight)
    If lv_ShowPageCount <= 0 Then lv_ShowPageCount = 1
    For lv_Index = lnVMargin.Count To lv_ShowPageCount - 1
        'Me.Controls.Add "VB.line", "lnVMargin(" & Str(lv_Index) & ")", WorkPage
        Load lnVMargin(lv_Index)
    Next
    For lv_Index = lnVMargin.LBound To lnVMargin.UBound
        If lv_Index >= lv_ShowPageCount Then
            Unload lnVMargin(lv_Index)
        Else
            
            lnVMargin(lv_Index).x1 = 1
            lnVMargin(lv_Index).y1 = lv_PageHeight * (lv_Index + 1)
            lnVMargin(lv_Index).x2 = WorkPage.Width
            lnVMargin(lv_Index).y2 = lv_PageHeight * (lv_Index + 1)
            lnVMargin(lv_Index).Visible = True
            
        End If
    Next
                
End Sub

Public Sub SetMouseClickPoint(ByVal x As Single, ByVal y As Single)
    m_ptMouseClick.x = x
    m_ptMouseClick.y = y
End Sub

Public Sub MovePasteNodeToRightPoint(f_Node As CNode, f_OldX As Long, f_OldY As Long)
    Dim lv_nRelativeX As Long
    Dim lv_nRelativeY As Long
    
    lv_nRelativeX = m_ptMouseClick.x - f_OldX
    lv_nRelativeY = m_ptMouseClick.y - f_OldY
    
    f_Node.Left = f_Node.Left + lv_nRelativeX
    f_Node.Top = f_Node.Top + lv_nRelativeY

End Sub

Private Sub WorkPage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        
        '' Draw a rectangle with dashed border
        Dim lv_RectBox As New CRect
        
        lv_RectBox.Left = m_ptMouseClick.x
        lv_RectBox.Top = m_ptMouseClick.y
        lv_RectBox.Right = x
        lv_RectBox.Bottom = y
        lv_RectBox.SetCtrlToRect shpSelectRegion
        
    End If
    
End Sub

Private Sub WorkPage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        
        '' Select Nodes in rectangle region
        Dim lv_RectBox As New CRect
        
        lv_RectBox.SetRectToCtrl shpSelectRegion
        m_objCallFlow.SetRegionNodeSelect lv_RectBox, True
        
        shpSelectRegion.Visible = False
        
    End If
    
End Sub

' f_Direction:
' 1 - Left
' 2 - Up
' 3 - Right
' 4 - Down
Private Sub MoveSelectedNoeds(ByVal f_Direction As Byte, Optional ByVal f_blnSmallStep As Boolean = False)
    
    Dim lv_X_Step As Integer, lv_Y_Step As Integer
    Dim lv_nNodeCount As Integer
    Dim lv_nIndex As Integer
    
    Select Case f_Direction
    Case 1                  '' Left
        lv_X_Step = -IIf(f_blnSmallStep, Move_Small_Step, Move_Big_Step)
        lv_Y_Step = 0
    Case 2                  '' Top
        lv_X_Step = 0
        lv_Y_Step = -IIf(f_blnSmallStep, Move_Small_Step, Move_Big_Step)
    Case 3                  '' Right
        lv_X_Step = IIf(f_blnSmallStep, Move_Small_Step, Move_Big_Step)
        lv_Y_Step = 0
    Case 4                  '' Bottom
        lv_X_Step = 0
        lv_Y_Step = IIf(f_blnSmallStep, Move_Small_Step, Move_Big_Step)
    Case Else
        lv_X_Step = 0
        lv_Y_Step = 0
    End Select
    
    lv_nNodeCount = m_objCallFlow.SelectedCount
    If lv_nNodeCount > 0 Then
        
        If gClipBoard.MultiPushInitialize(lv_nNodeCount, DEF_OPERATION_MODIFY) Then
            
            For lv_nIndex = 1 To m_objCallFlow.NewNodeID
                
                If m_objCallFlow.Node(lv_nIndex).IsSelected Then
                                            
                    gClipBoard.MultiPushClipBoardStack lv_nIndex
                   
                    If m_objCallFlow.Node(lv_nIndex).NodeNo = 255 Then
                        '' Move Image with Line
                        m_objCallFlow.Node(lv_nIndex).Line_X1 = m_objCallFlow.Node(lv_nIndex).Line_X1 + lv_X_Step
                        m_objCallFlow.Node(lv_nIndex).Line_Y1 = m_objCallFlow.Node(lv_nIndex).Line_Y1 + lv_Y_Step
                        m_objCallFlow.Node(lv_nIndex).Line_X2 = m_objCallFlow.Node(lv_nIndex).Line_X2 + lv_X_Step
                        m_objCallFlow.Node(lv_nIndex).Line_Y2 = m_objCallFlow.Node(lv_nIndex).Line_Y2 + lv_Y_Step
                        m_objCallFlow.Node(lv_nIndex).MoveImageWithLine
                    Else
                        m_objCallFlow.Node(lv_nIndex).Left = m_objCallFlow.Node(lv_nIndex).Left + lv_X_Step
                        m_objCallFlow.Node(lv_nIndex).Top = m_objCallFlow.Node(lv_nIndex).Top + lv_Y_Step
                    End If
                        
                    Call UpdateCurrentPositionDisplay(lv_nIndex)
    
                    '' Move Line(s) with Node
                    MoveLinesOnNode lv_nIndex
                    
                    ''' 保存节点位置数据
                    m_objCallFlow.UpdateAnotherIVRRecord lv_nIndex
                    
                End If
            
            Next
        
            RefreshHandlesPosition
        
        End If
        
    End If

End Sub

Public Sub RefreshHandlesPosition()
    If Not m_CurrCtl Is Nothing Then
        m_DragRect.SetRectToCtrl m_CurrCtl
        ShowHandles True
    End If
End Sub

Public Sub MouseSelectNode(ByVal f_nIndex As Integer)
    m_objCallFlow.NodeSelectedID = f_nIndex
End Sub

Public Sub ChangeAllSelectedNodesStatus(ByVal f_blnSelected As Boolean)
    m_objCallFlow.SetAllNodeSelect f_blnSelected
End Sub

Public Sub MoveLinesOnNode(ByVal f_nIndex As Integer)
    Dim lv_Position As Integer, lv_NodeID As Integer
    
    If m_objCallFlow.Node(f_nIndex).NodeNo = 255 Then Exit Sub
    
    '' Move Line(s) with Node
    lv_Position = 1
    Do
        lv_NodeID = m_objCallFlow.FindLinesOnNode(m_objCallFlow.Node(f_nIndex), lv_Position)
        If lv_Position < 0 Then Exit Do
        lv_Position = lv_Position + 1
    Loop While True

End Sub

'Michael Added @ 2007-11-28 for check vscoll button status
Public Sub CheckStatus()
'    If gCallFlow.PageCount <= 1 Then
'        cmdPage(0).Enabled = False
'        cmdPage(1).Enabled = False
'    Else
        'If gCallFlow.CurrentPage = 1 And gCallFlow.PageCount <> 0 Then
        If gCallFlow.CurrentPage = 1 Then
            cmdPage(0).Enabled = False
            cmdPage(1).Enabled = True
'        ElseIf gCallFlow.CurrentPage = 1 And gCallFlow.PageCount = 0 Then
'            cmdPage(0).Enabled = False
'            cmdPage(1).Enabled = False
        ElseIf gCallFlow.CurrentPage = gCallFlow.PageCount Then
            cmdPage(0).Enabled = True
            cmdPage(1).Enabled = False
        Else
            cmdPage(0).Enabled = True
            cmdPage(1).Enabled = True
        End If
'    End If
End Sub
