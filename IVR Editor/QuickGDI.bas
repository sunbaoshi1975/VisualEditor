Attribute VB_Name = "QuickGDI"
Option Explicit

Dim m_hDC As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function CreateSolidBrush Lib "GDI32" _
     (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "GDI32" _
     (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "GDI32" _
     (ByVal hObject As Long) As Integer

Declare Function GetSysColor Lib "user32" _
     (ByVal nIndex As ColConst) As Long
'Color constants for GetSysColor
Public Enum ColConst
    COLOR_ACTIVEBORDER = 10
    COLOR_ACTIVECAPTION = 2
    COLOR_ADJ_MAX = 100
    COLOR_ADJ_MIN = -100
    COLOR_APPWORKSPACE = 12
    COLOR_BACKGROUND = 1
    COLOR_BTNFACE = 15
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_GRAYTEXT = 17
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_INACTIVEBORDER = 11
    COLOR_INACTIVECAPTION = 3
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_MENU = 4
    COLOR_MENUTEXT = 7
    COLOR_SCROLLBAR = 0
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_WINDOWTEXT = 8
End Enum

Private Declare Function GetTextColor Lib "GDI32" _
     (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "GDI32" _
     (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "GDI32" _
     (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Const NEWTRANSPARENT = 3 'use with SetBkMode()

Private Declare Function CreatePen Lib "GDI32" _
     (ByVal nPenStyle As Long, ByVal nWidth As Long, _
     ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" _
     (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
     lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "GDI32" _
     (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "GDI32" _
     (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20

Public Sub DrawRect(ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long)
     If m_hDC = 0 Then Exit Sub
     Call Rectangle(m_hDC, X1, Y1, X2, Y2)
End Sub

Public Function GetPen(ByVal nWidth As Long, _
ByVal Clr As Long) As Long
     GetPen = CreatePen(0, nWidth, Clr)
End Function

Public Function hPrint(rc As RECT, ByVal hStr As String, ByVal Clr As Long) As Long
     If m_hDC = 0 Then Exit Function
     'Equivalent to setting a form's property
     'FontTransparent = True
     SetBkMode m_hDC, NEWTRANSPARENT

     Dim OT As Long
     OT = GetTextColor(m_hDC)
     SetTextColor m_hDC, Clr
     'Print the text
     hPrint = DrawText(m_hDC, hStr, LenB(StrConv(hStr, vbFromUnicode)), rc, DT_VCENTER + DT_SINGLELINE)
     'Restore old text color
     SetTextColor m_hDC, OT
End Function

Public Property Get TargethDC() As Long
     TargethDC = m_hDC
End Property
Public Property Let TargethDC(ByVal vNewValue As Long)
     'The hDC to draw to when performing operations
     'from this module's subroutines.
     m_hDC = vNewValue
End Property

Public Sub ThreedBox(ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long, _
Optional Sunken As Boolean = False)
     'Draw a raised box around the specified
     'coordinates.

     If m_hDC = 0 Then Exit Sub

     Dim CurPen As Long, OldPen As Long
     Dim dm As POINTAPI

     If Sunken = False Then
         CurPen = GetPen(1, GetSysColor(COLOR_BTNHIGHLIGHT))
     Else
          CurPen = GetPen(1, GetSysColor(COLOR_BTNSHADOW))
     End If
     OldPen = SelectObject(m_hDC, CurPen)
     'FirstLightLine
     MoveToEx m_hDC, X1, Y2, dm
     LineTo m_hDC, X1, Y1
     'SecondLightLine
     LineTo m_hDC, X2, Y1

     SelectObject m_hDC, OldPen
     DeleteObject CurPen
     If Sunken = False Then
          CurPen = GetPen(1, GetSysColor(COLOR_BTNSHADOW))
     Else
          CurPen = GetPen(1, GetSysColor(COLOR_BTNHIGHLIGHT))
     End If
     OldPen = SelectObject(m_hDC, CurPen)
     'FirstDarkLine
     MoveToEx m_hDC, X2, Y1, dm
     LineTo m_hDC, X2, Y2
     'SecondDarkLine
     LineTo m_hDC, X1, Y2

     SelectObject m_hDC, OldPen
     DeleteObject CurPen
End Sub

'--end block--'
