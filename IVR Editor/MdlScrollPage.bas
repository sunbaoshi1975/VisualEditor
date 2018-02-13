Attribute VB_Name = "MdlScrollPage"
Option Explicit

'/* Michael added this module for support mouse wheel
' * @2008-8-29
' */

Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'' GWL_WNDPROC already defined in module 'MenuItems'
'Public Const GWL_WNDPROC = -4
Public Const SPI_GETWHEELSCROLLLINES = 104
Public Const WM_MOUSEWHEEL = &H20A
Public WHEEL_SCROLL_LINES As Long
Global lProc As Long

'Mouse Coordinate
Public Type Coordinate
    x As Long
    y As Long
End Type
  
Public Sub Hook(ByVal hWnd As Long)
    lProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
  
Public Sub UnHook(ByVal hWnd As Long)
    Dim r As Long
    r = SetWindowLong(hWnd, GWL_WNDPROC, lProc)
End Sub
  
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MP As Coordinate
    If uMsg = WM_MOUSEWHEEL Then
        
'       'get mouse location
        Dim MW_FB As Integer
        MW_FB = (wParam And &HFFFF0000) \ &H10000
'        MP.x = (lParam And &HFFFF&)
'        MP.y = (lParam And &HFFFF0000) \ &H10000

        Dim k As Integer
        k = (frmMain.ActiveForm.VScroll.value - Sgn(MW_FB) * frmMain.ActiveForm.VScroll.LargeChange)
        If k > frmMain.ActiveForm.VScroll.Max Then k = frmMain.ActiveForm.VScroll.Max
        If k < frmMain.ActiveForm.VScroll.Min Then k = frmMain.ActiveForm.VScroll.Min
        frmMain.ActiveForm.VScroll.value = k
    End If
    WindowProc = CallWindowProc(lProc, hw, uMsg, wParam, lParam)
End Function
