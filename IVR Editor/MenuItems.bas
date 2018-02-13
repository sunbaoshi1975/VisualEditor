Attribute VB_Name = "MenuItems"
Option Explicit

Dim hMenu As Long
Dim hSubMenu As Long
Dim mnuID As Long

Dim m_Form As frmMain

'Subclassing stuff we'll need...
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
     lParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias _
     "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
     ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
     (ByVal hWnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)
'Messages to use in the wndproc
Public Const WM_SYSCOMMAND = &H112
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUSELECT = &H11F
Public Const WM_COMMAND = &H111
Public Const WM_GETFONT = &H31

Type MENUITEMINFO
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long
     cch As Long
End Type
Public Const MIIM_TYPE = &H10

Type MEASUREITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemWidth As Long
     itemHeight As Long
     ItemData As Long
End Type
Type DRAWITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemAction As Long
     itemState As Long
     hwndItem As Long
     hdc As Long
     rcItem As RECT
     ItemData As Long
End Type

Declare Function GetMenu Lib "user32" _
     (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" _
     (ByVal hMenu As Long, ByVal nPos As Long) _
     As Long
Declare Function GetMenuItemID Lib "user32" _
     (ByVal hMenu As Long, ByVal nPos As Long) _
     As Long

Declare Function ModifyMenu Lib "user32" Alias _
     "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, _
     ByVal wFlags As Long, ByVal wIDNewItem As Long, _
     ByVal lpString As Long) As Long

Declare Function GetMenuItemInfo Lib "user32" Alias _
     "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, _
     ByVal ByPosition As Long, lpMenuItemInfo As MENUITEMINFO) _
     As Boolean

Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400
Private Const MF_OWNERDRAW = &H100
Private Const MF_SEPARATOR = &H800
Public Const MFT_SEPARATOR = MF_SEPARATOR

Public Const ODS_SELECTED = &H101
Public Const ODS_DISABLED = &H4
Public Const ODS_GRAYED = &H2

Public Property Get MenuForm() As frmMain
     Set MenuForm = m_Form
End Property
Public Property Let MenuForm(ByVal vNewValue As frmMain)
     Set m_Form = vNewValue
     hMenu = GetMenu(m_Form.hWnd)
End Property

Public Property Get MenuID() As Long
     MenuID = mnuID
End Property
Public Property Let MenuID(ByVal vNewValue As Long)
     mnuID = GetMenuItemID(hSubMenu, vNewValue)
End Property

Public Sub OwnerDrawMenu(ByVal ItemData As Long)
     'Change the menu's style to owner-draw. You must
     'now subclass the form that this menu is on so
     'you can respond to the WM_MEASUREITEM and WM_DRAWITEM
     'messages.
     Dim mii As MENUITEMINFO
     mii.cbSize = Len(mii)
     mii.fMask = MIIM_TYPE
     GetMenuItemInfo hSubMenu, MenuID, False, mii
     If ((mii.fType And MF_SEPARATOR) = MF_SEPARATOR) Then
          '*Preserve* separator style...
          Call ModifyMenu(hSubMenu, MenuID, _
               MF_BYCOMMAND Or MF_OWNERDRAW Or MF_SEPARATOR, _
               MenuID, ItemData)
     Else
          Call ModifyMenu(hSubMenu, MenuID, _
               MF_BYCOMMAND Or MF_OWNERDRAW, MenuID, ItemData)
     End If
End Sub

Public Function OwnMenuProc(ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long
     OwnMenuProc = frmMain.MsgProc(hWnd, wMsg, wParam, lParam)
End Function

Public Sub SetTopMenu(NewMnu As Long)
     hMenu = NewMnu
End Sub

Public Property Get SubMenu() As Long
     SubMenu = hSubMenu
End Property
Public Property Let SubMenu(ByVal vNewValue As Long)
     hSubMenu = GetSubMenu(hMenu, vNewValue)
End Property
'--end block--'
