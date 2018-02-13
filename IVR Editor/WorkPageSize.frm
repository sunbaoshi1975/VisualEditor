VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form WorkPageProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "画布属性"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2430
   Icon            =   "WorkPageSize.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1475"
   Begin VB.Frame Frame1 
      Caption         =   "工作页面"
      Height          =   2295
      Left            =   143
      TabIndex        =   2
      Tag             =   "1476"
      Top             =   90
      Width           =   2145
      Begin ComCtl2.UpDown UpdPageCount 
         Height          =   285
         Left            =   1006
         TabIndex        =   13
         Top             =   1830
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPageCount"
         BuddyDispid     =   196611
         OrigLeft        =   1260
         OrigTop         =   1830
         OrigRight       =   1500
         OrigBottom      =   2115
         Max             =   20
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPageCount 
         Height          =   285
         Left            =   480
         TabIndex        =   12
         Top             =   1830
         Width           =   525
      End
      Begin VB.CommandButton custwpbkcolor 
         Caption         =   "..."
         Height          =   315
         Left            =   1590
         TabIndex        =   9
         ToolTipText     =   "1520"
         Top             =   270
         Width           =   345
      End
      Begin VB.PictureBox wpbkcolor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         ScaleHeight     =   255
         ScaleWidth      =   465
         TabIndex        =   8
         Top             =   300
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Caption         =   "区域"
         Height          =   1035
         Left            =   150
         TabIndex        =   3
         Tag             =   "1441"
         Top             =   660
         Width           =   1845
         Begin VB.TextBox txtHeight 
            Height          =   285
            Left            =   720
            TabIndex        =   5
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtWidth 
            Height          =   285
            Left            =   720
            TabIndex        =   4
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "高:"
            Height          =   180
            Left            =   150
            TabIndex        =   7
            Tag             =   "1443"
            Top             =   630
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "宽:"
            Height          =   180
            Left            =   150
            TabIndex        =   6
            Tag             =   "1442"
            Top             =   270
            Width           =   270
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "页"
         Height          =   180
         Index           =   2
         Left            =   1380
         TabIndex        =   14
         Tag             =   "1133"
         Top             =   1860
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "共"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Tag             =   "1444"
         Top             =   1860
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "画布颜色:"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Tag             =   "1477"
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出"
      Height          =   315
      Index           =   1
      Left            =   1245
      TabIndex        =   1
      Tag             =   "1144"
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Tag             =   "1372"
      Top             =   2520
      Width           =   945
   End
End
Attribute VB_Name = "WorkPageProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Color As CHOOSECOLOR

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
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
    
    gCallFlow.PageCount = Val(txtPageCount)
    gPageWidth = Val(txtWidth.Text) * Screen.TwipsPerPixelX
    gPageHeight = Val(txtHeight.Text) * Screen.TwipsPerPixelY
    gCallFlow.SetWorkPageScale 0, 0, gPageWidth, gPageHeight
        
    gPageBackColor = wpbkcolor.BackColor
    gCallFlow.SetWorkPageBackColor gPageBackColor
    
    'Michael Added @ 2007-11-27
    'Michael commented @ 2007-12-5
    'Call CFlowWorks.CheckStatus
        
    ' Write to INI
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_HE, Str(gPageHeight), gSystem.strINI_File)
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_WD, Str(gPageWidth), gSystem.strINI_File)
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_BG, Str(gPageBackColor), gSystem.strINI_File)
    
    frmMain.StatusBar.Panels("Page").Text = LoadNationalResString(1554) & Trim(Str(gCallFlow.CurrentPage)) & LoadNationalResString(1132) & Trim(Str(gCallFlow.PageCount)) & LoadNationalResString(1133)
    
    Unload Me
End Sub

Private Sub custwpbkcolor_Click()
    F_CustomColor gPageBackColor, wpbkcolor
End Sub

Private Sub Form_Load()
On Error Resume Next

    txtWidth.Text = Format(Str(gPageWidth / Screen.TwipsPerPixelX))
    txtHeight.Text = Format(Str(gPageHeight / Screen.TwipsPerPixelY))
    wpbkcolor.BackColor = gPageBackColor
    
    '' Sun added
    UpdPageCount.Min = gCallFlow.PageCount
    If gCallFlow.PageCount < 1 Then gCallFlow.PageCount = 1
    txtPageCount = gCallFlow.PageCount
    LoadResStrings Me
End Sub

Private Sub txtPageCount_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

