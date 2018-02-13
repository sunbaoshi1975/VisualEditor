VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmPrintPreview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "◊ ‘¥¥Ú”°‘§¿¿"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14775
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgListArrow 
      Left            =   5040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":030A
            Key             =   "RIGHT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":041C
            Key             =   "RIGHTDis"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":052E
            Key             =   "LEFT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":0640
            Key             =   "LEFTDis"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":0752
            Key             =   "ZOOM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":0AD4
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":0BE6
            Key             =   "CLOSE"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":0CF8
            Key             =   "SETUP"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":0E0A
            Key             =   "LANDSCAPE"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintPreview.frx":1884
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   495
      Left            =   -120
      TabIndex        =   1
      Top             =   120
      Width           =   14775
      _Version        =   458752
      _ExtentX        =   26061
      _ExtentY        =   873
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
      SpreadDesigner  =   "frmPrintPreview.frx":2B3B
   End
   Begin FPSpreadADO.fpSpreadPreview fpSpreadPreview1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   14535
      _Version        =   458752
      _ExtentX        =   25638
      _ExtentY        =   15690
      _StockProps     =   96
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************************
'File Name   : frmPrintPreview
'Create Date : Sep,7,07
'Modify Date : Sep,7,07
'Author      : Michael
'Contect     :
'*******************************************************************************************************************

Option Explicit

Private Sub Form_Activate()
 
    'Attach preview control to Spread
    'spreadpreview.fpSpreadPreview1.hWndSpread = mainFrm.fpSpread1.hWnd
    frmPrintPreview.fpSpreadPreview1.hWndSpread = frmResPrintDoc.spdResource.hWnd
   
    'Update page count listing
    UpdatePageCount
End Sub

Private Sub Form_Load()
   
    SetupToolbar
    
    'Disable Previous button
    DisableButton 4, "LEFT"
        
    'Get the zoom display
    GetZoom zoomindex
    
    'Set PrintFooter and PrintHead
    headerfooter.SaveConfiguration
    
    'Set up page numbering
    If frmResPrintDoc.spdResource.PrintPageCount = 1 Then
        'Disable Next button if only one page
        DisableButton 2, "LEFT"
    End If
           
End Sub
Sub SetupToolbar()
Dim i As Integer

    'Specify whether Edit Mode is to remain on when switching between cells
    fpSpread1.EditModePermanent = True

    fpSpread1.Col = -1
    fpSpread1.Row = -1
    fpSpread1.Lock = True
    
    'Set the number of rows in the spreadsheet
    fpSpread1.MaxRows = 1
 
    'Set the height of a selected row
    fpSpread1.RowHeight(1) = 15
   
    'Set the number of columns in the spreadsheet
    fpSpread1.MaxCols = 19
 
    'Set the column widths
    For i = 1 To fpSpread1.MaxCols Step 2
        fpSpread1.ColWidth(i) = 0.3
    Next i
   
    'Resize wide column
    fpSpread1.ColWidth(14) = 15
    
    'Show or hide the column headers
    fpSpread1.DisplayColHeaders = False
    fpSpread1.DisplayRowHeaders = False
    
    'Turn off scroll bars
    fpSpread1.ScrollBars = ScrollBarsNone
    
    'Turn off border
    fpSpread1.BorderStyle = BorderStyleNone
      
    'Select row(s)
    fpSpread1.Row = 1
    fpSpread1.Col = -1

    'Determine the color of background, foreground and border color
    fpSpread1.ForeColor = RGB(0, 0, 0)
    fpSpread1.BackColor = RGB(192, 192, 192)
    fpSpread1.fontname = "MS Sans Serif"
    fpSpread1.FontSize = 8
    fpSpread1.FontBold = False
    
    'Select a single cell
    fpSpread1.Col = 2
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Next"
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\RIGHT.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\RIGHT.BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages("RIGHT").Picture
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_LEFT
    
    'Select a single cell
    fpSpread1.Col = 4
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Previous"
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\LEFT.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\LEFT.BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages("LEFT").Picture
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    'Select a single cell
    fpSpread1.Col = 6
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Zoom"
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\ZOOM.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\ZOOM.BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages("ZOOM").Picture
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    'Select a single cell
    fpSpread1.Col = 8
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Print"
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\PRINT.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\PRINT.BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages("PRINT").Picture
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    'Select a single cell
    fpSpread1.Col = 10
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Setup"
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\SETUP.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\SETUP.BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages("SETUP").Picture
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    
    
    'Select a single cell
    fpSpread1.Col = 16
    fpSpread1.Row = 1

    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "Close"
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\CLOSE.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\CLOSE.BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages("CLOSE").Picture
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    fpSpread1.TextTip = TextTipFloating
    Dim bRet As Boolean
    bRet = fpSpread1.SetTextTipAppearance("MS Sans Serif", 8, 0, 0, &HC0FFFF, &H0)
    fpSpread1.CursorType = CursorTypeLockedCell
    fpSpread1.CursorStyle = CursorStyleArrow
    fpSpread1.NoBeep = True
    
    '****************************
    fpSpread1.Col = 18
    fpSpread1.Row = 1
    
    'Define cells as type BUTTON
    fpSpread1.CellType = SS_CELL_TYPE_BUTTON
    fpSpread1.Lock = False
    fpSpread1.TypeButtonText = "About"
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\SETUP.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\info.ico")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages("INFO").Picture
    fpSpread1.TypeButtonAlign = SS_CELL_BUTTON_ALIGN_RIGHT
    fpSpread1.TextTip = TextTipFloating
    bRet = fpSpread1.SetTextTipAppearance("MS Sans Serif", 8, 0, 0, &HC0FFFF, &H0)
    fpSpread1.CursorType = CursorTypeLockedCell
    fpSpread1.CursorStyle = CursorStyleArrow
    fpSpread1.NoBeep = True
    
End Sub
Sub DisableButton(Col As Long, bitmapdirection As String)
'Disable specified button
    fpSpread1.ReDraw = False
    
    fpSpread1.Row = 1
    fpSpread1.Col = Col
    
    fpSpread1.Lock = True
    fpSpread1.TypeButtonTextColor = RGB(128, 128, 128)
    fpSpread1.Protect = True
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\" & bitmapdirection & "DIS.BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\" & bitmapdirection & "DIS.BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages(Trim(bitmapdirection) & "Dis").Picture
    
    fpSpread1.ReDraw = True
End Sub
Sub EnableButton(Col As Long, bitmapdirection As String)
'Enable specified button
    fpSpread1.ReDraw = False
    
    fpSpread1.Row = 1
    fpSpread1.Col = Col
    
    fpSpread1.Lock = False
    fpSpread1.TypeButtonTextColor = RGB(0, 0, 0)
    fpSpread1.Protect = False
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\..\..\..\files\" & bitmapdirection & ".BMP")
    'Set fpSpread1.TypeButtonPicture = LoadPicture(App.path & "..\bmp\" & bitmapdirection & ".BMP")
    Set fpSpread1.TypeButtonPicture = imgListArrow.ListImages(Trim(bitmapdirection)).Picture
    
    fpSpread1.ReDraw = True
End Sub

Private Sub Form_Resize()
    fpSpread1.Move 0, 0, ScaleWidth, fpSpread1.Height
    fpSpreadPreview1.Move 0, fpSpread1.Height, ScaleWidth, ScaleHeight - fpSpread1.Height
End Sub

Private Sub fpSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    fpSpread1.Col = Col
    fpSpread1.Row = Row
    
    If fpSpread1.CellType = CellTypeButton Then
        Select Case Col
            Case 2  'Next
                If fpSpreadPreview1.PageCurrent < frmResPrintDoc.spdResource.PrintPageCount Then
                    fpSpreadPreview1.PageCurrent = fpSpreadPreview1.PageCurrent + fpSpreadPreview1.PagesPerScreen
                    EnableButton Col, "RIGHT"
                    'Enable Previous button
                    EnableButton 4, "LEFT"
                   'Update page count listing
'                    UpdatePageCount
                End If
                
                 'If at last page, disable button
                    If fpSpreadPreview1.PageCurrent >= frmResPrintDoc.spdResource.PrintPageCount Then
                        DisableButton Col, "RIGHT"
                    End If
            Case 4  'Previous
                If fpSpreadPreview1.PageCurrent > 1 Then
                    fpSpreadPreview1.PageCurrent = fpSpreadPreview1.PageCurrent - fpSpreadPreview1.PagesPerScreen
                    EnableButton Col, "LEFT"
                    EnableButton 2, "RIGHT"
                    'Update page count listing
'                    UpdatePageCount
                End If
                
                'If at first page, disable button
                If fpSpreadPreview1.PageCurrent = 1 Then
                    DisableButton Col, "LEFT"
                End If
                
            Case 6  'Zoom
                fpSpreadPreview1.ZoomState = 3
                
            Case 8  'Print
                PrintDlg.Show 1
                                 
            Case 10 'Setup
                pagesetup.Show 1
             
            Case 16 'Close
                Unload Me
                
            Case 18 'info
                PrnINFO.Show 1
                
        End Select
    End If
End Sub
Sub UpdatePageCount()
 'Page Count
    fpSpread1.Row = 1
    fpSpread1.Col = 14
    fpSpread1.Text = "Page " & fpSpreadPreview1.PageCurrent & " of " & frmResPrintDoc.spdResource.PrintPageCount
    
End Sub


Private Sub fpSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    With fpSpread1
        .Col = Col
        .Row = Row
        If .CellType = CellTypeButton And Not .Lock Then
            ShowTip = True
            TipText = .TypeButtonText
        ElseIf .CellType = CellTypeEdit And .Text <> "" Then
            ShowTip = True
            TipText = .Text
        End If
    End With

End Sub

Private Sub fpSpreadPreview1_PageChange(ByVal Page As Long)
    UpdatePageCount
End Sub
