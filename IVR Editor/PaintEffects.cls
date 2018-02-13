VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaintEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Halftone created for default palette use
Private m_hpalHalftone As Long

'************************************************************************
'     Michael Add this note
' GDI32.dll是Windows GDI图形用户界面相关程序，用于辅助创建组件,Graphical Device Interface
' User32.dll是Windows用户界面相关应用程序接口，用于包括Windows处理，基本用户界面等特性
' Oleaut32.dll是对象链接与嵌入OLE相关文件 (Microsoft OLE DLL)
'************************************************************************

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, _
     ByVal X As Long, ByVal Y As Long, _
     ByVal hIcon As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" _
     (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, _
     ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, _
     ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "GDI32" (ByVal hdc As Long, _
     ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" _
     (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" _
     (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
     ByVal hdc As Long) As Integer
Private Declare Function CreateCompatibleBitmap Lib "GDI32" _
     (ByVal hdc As Long, ByVal nWidth As Long, _
     ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, _
     ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "GDI32" (ByVal hdc As Long, _
     ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal nPlanes As Long, _
     ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "GDI32" _
     (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "GDI32" _
     (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hdc As Long, _
     ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" _
     (ByVal hdc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "GDI32" _
     (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" _
     (ByVal lOleColor As Long, ByVal lHPalette As Long, _
     lColorRef As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
     lpRect As RECT, ByVal hBrush As Long) As Integer

'Raster Operation Codes
Private Const DSna = &H220326 '0x00220326


Public Sub PaintTransparentDC(ByVal hdcDest As Long, _
               ByVal xDest As Long, _
               ByVal yDest As Long, _
               ByVal Width As Long, _
               ByVal Height As Long, _
               ByVal hdcSrc As Long, _
               ByVal XSrc As Long, _
               ByVal YSrc As Long, _
               ByVal clrMask As OLE_COLOR, _
               Optional ByVal hPal As Long = 0)

    Dim hdcMask As Long     'HDC of the created mask image
    Dim hdcColor As Long    'HDC of the created color image
    Dim hbmMask As Long     'Bitmap handle to the mask image
    Dim hbmColor As Long    'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long 'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    
    hdcScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor

    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the
    'destination when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcDest, xDest, _
        yDest, vbSrcCopy
    
    'Create a (color) bitmap for the cover (can't use
    'CompatibleBitmap with hdcSrc, because this will create a
    'DIB section if the original bitmap is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this
    'first and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome
    'bitmap does a nearest-color selection rather than painting
    'based on the backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set
    'the destination foreground/background colors according to
    'those currently set in hdcSrc (because Windows will
    'associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, XSrc, _
        YSrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent
    'color from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.
    'All other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, _
        vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color,
    'and the original colors everywhere else.  To do this,
    'we first paint the original onto the cover (which we
    'already did), then we AND the inverse of the mask onto
    'that using the DSna ternary raster operation
    '(0x00220326 - see Win32 SDK reference, Appendix,
    '"Raster Operation Codes", "Ternary
    'Raster Operations", or search in MSDN for 00220326).
    'DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows
    'transforms all white bits (1) to the background color
    'of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, _
        0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, _
        0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, _
        0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hdcDest, xDest, yDest, Width, Height, _
        hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
End Sub


Public Sub PaintTransparentStdPic(ByVal hdcDest As Long, _
                ByVal xDest As Long, _
                ByVal yDest As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByVal picSource As Picture, _
                ByVal XSrc As Long, _
                ByVal YSrc As Long, _
                ByVal clrMask As OLE_COLOR, _
                Optional ByVal hPal As Long = 0)

    Dim hdcSrc As Long 'HDC for source bitmap
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long
    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo _
        PaintTransparentStdPic_InvalidParam
    
    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.Handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hdcDest, xDest, yDest, _
                Width, Height, hdcSrc, XSrc, YSrc, clrMask, hPal

            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, _
                Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.Handle
            'Draw Transparent image
            PaintTransparentDC hdcDest, xDest, yDest, Width, _
                Height, hdcSrc, 0, 0, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub
PaintTransparentStdPic_InvalidParam:
    Exit Sub
End Sub
'--end block--'
