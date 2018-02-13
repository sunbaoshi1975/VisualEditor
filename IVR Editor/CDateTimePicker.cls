VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDateTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class       : CDateTimePicker
' Description : This class implements the DateTime Picker Windows common control.
' Source      : JustACoder
'

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function API_InitCommonControlsEx Lib "comctl32.dll" Alias "InitCommonControlsEx" (inex As InitCommonControlsEx) As Long

Private Const WS_CHILD = &H40000000
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000   '  WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VSCROLL = &H200000
Private Const WS_VISIBLE = &H10000000

Private Type InitCommonControlsEx
  dwSize As Long
  dwICC As Long
End Type

Private Const WM_USER = &H400
Private Const DATETIMEPICK_CLASS = "SysDateTimePick32"
Private Const ICC_DATE_CLASSES = &H100&
Private Const DTS_UPDOWN = &H1
Private Const DTS_SHOWNONE = &H2
Private Const DTS_SHORTDATEFORMAT = &H0
Private Const DTS_LONGDATEFORMAT = &H4
Private Const DTS_TIMEFORMAT = &H9
Private Const DTS_APPCANPARSE = &H10
Private Const DTS_RIGHTALIGN = &H20
Private Const DTM_GETSYSTEMTIME = &H1001&
Private Const DTM_SETSYSTEMTIME = &H1002&
Private Const DTM_SETFORMAT = &H1005&

Private m_lnghWndDTP As Long
Private m_fDisplayTime As Boolean
Private m_fShowCalendar As Boolean
Private m_hWndParent As Long
Private m_lngLeft As Long
Private m_lngTop As Long
Private m_lngHeight As Long
Private m_lngWidth As Long
Private m_bVisible As Boolean

Private Sub Class_Terminate()
  ' Clean up the window
  ' Source: JustACoder

  Destroy

End Sub

Public Property Get DateTime() As Date
  ' Returns: The date time value of the control
  ' Source: JustACoder
  '
  Dim strDateTime As String
  Dim st As SYSTEMTIME
  
  On Error GoTo PROC_ERR
  
  ' Retrieve the date time from the control
  SendMessage m_lnghWndDTP, DTM_GETSYSTEMTIME, 0, st
  
  ' Format the date time string
  strDateTime = st.wMonth & "/" & st.wDay & "/" & st.wYear & " " & st.wHour _
    & ":" & st.wMinute & ":" & st.wSecond
  
  ' Return the date time
  DateTime = strDateTime

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, vbExclamation, _
    "DateTime"
  Resume PROC_EXIT

End Property

Public Property Let DateTime(ByVal dtValue As Date)
  ' dtValue: Set the current date time of the date time picker
  ' Source: JustACoder
  '
  Dim st As SYSTEMTIME

  On Error GoTo PROC_ERR

  ' Fill the SYSTEMTIME structure
  st.wYear = Year(dtValue)
  st.wMonth = Month(dtValue)
  st.wDay = Day(dtValue)
  st.wHour = Hour(dtValue)
  st.wMinute = Minute(dtValue)
  st.wSecond = Second(dtValue)

  ' Set the date time in the control
  SendMessage m_lnghWndDTP, DTM_SETSYSTEMTIME, 0, st

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, vbExclamation, _
    "DateTime"
  Resume PROC_EXIT

End Property

Public Property Get DisplayTime() As Boolean
  ' Returns: True if we are displaying time, False if we are displaying date
  ' Source: JustACoder
  
  DisplayTime = m_fDisplayTime
  
End Property

Public Property Let DisplayTime(ByVal fValue As Boolean)
  ' fValue: Set a flag indicating if we show the time picker
  ' Source: JustACoder
  
  m_fDisplayTime = fValue
  
End Property

Public Property Let Format(ByVal strValue As String)
  ' strValue: Set the display format of of the datetime control
  ' Source: JustACoder
  
  SendMessage m_lnghWndDTP, DTM_SETFORMAT, 0, ByVal strValue
  
End Property

Public Property Get Height() As Long
  ' Returns: The Height of the datetime control
  ' Source: JustACoder
  
  Height = m_lngHeight
  
End Property

Public Property Let Height(ByVal lngValue As Long)
  ' lngValue: Set the Height of the datetime control. This is measured in pixels
  ' Source: JustACoder
  
  m_lngHeight = lngValue
  
End Property

Public Property Get hwnd() As Long
  ' Returns: The handle of the datetime control
  ' Source: JustACoder
  
  hwnd = m_lnghWndDTP
  
End Property

Public Property Get hWndParent() As Long
  ' Returns: The handle to the parent window
  ' Source: JustACoder
  
  hWndParent = m_hWndParent
  
End Property

Public Property Let hWndParent(ByVal lngValue As Long)
  ' lngValue: Set the handle to the parent window
  ' Source: JustACoder
  
  m_hWndParent = lngValue
  
End Property

Public Property Get Left() As Long
  ' Returns: The left position of the datetime control
  ' Source: JustACoder
  
  Left = m_lngLeft
  
End Property

Public Property Let Left(ByVal lngValue As Long)
  ' lngValue: Set the left position of the datetime control. This is measured in
  ' pixels, and is relative to the parent window
  ' Source: JustACoder
  
  m_lngLeft = lngValue
  
End Property

Public Property Get ShowCalendar() As Boolean
  ' Returns: True if we display a drop down calendar, False if we are not
  ' Source: JustACoder
  
  ShowCalendar = m_fShowCalendar
  
End Property

Public Property Let ShowCalendar(ByVal fValue As Boolean)
  ' fValue: Set a flag indicating if we show the drop down calendar. Not
  ' applicable if the time picker is shown
  ' Source: JustACoder
  
  m_fShowCalendar = fValue
  
End Property

Public Property Get Top() As Long
  ' Returns: The top position of the datetime control
  ' Source: JustACoder
  
  Top = m_lngTop
  
End Property

Public Property Let Top(ByVal lngValue As Long)
  ' lngValue: Set the Top position of the datetime control. This is measured in
  ' pixels, and is relative to the parent window
  ' Source: JustACoder
  
  m_lngTop = lngValue
  
End Property

Public Property Get Width() As Long
  ' Returns: The Width of the datetime control
  ' Source: JustACoder
  
  Width = m_lngWidth
  
End Property

Public Property Let Width(ByVal lngValue As Long)
  ' lngValue: Set the Width of the datetime control. This is measured in pixels
  ' Source: JustACoder
  
  m_lngWidth = lngValue

End Property

Public Function Create() As Boolean
  ' Comments  : Create the DateTime picker control.
  ' Parameters: None
  ' Returns   : True if the control was created, False otherwise
  ' Source    : JustACoder
  '
  Dim icex As InitCommonControlsEx
  Dim lngFlags As Long
  Dim lnghInstanceLib As Long
  Dim lngProcAddress As Long
  
  On Error GoTo PROC_ERR
  
    ' Assume control creation fails
  Create = False
 
  lngFlags = WS_BORDER Or WS_CHILD Or WS_TABSTOP Or WS_GROUP
  
  ' If we are not showing the calendar, then set the updown flag
  If Not m_fShowCalendar Then
    lngFlags = lngFlags Or DTS_UPDOWN
  End If
  
  ' If we are displaying time, then set the timeformat flag
  If m_fDisplayTime Then
    lngFlags = lngFlags Or DTS_TIMEFORMAT
  End If
  
  ' Load common control 32 bit library
  lnghInstanceLib = LoadLibrary("comctl32.dll")
  
  ' If the handle is valid, try to get the function address.
  If (lnghInstanceLib <> 0) Then
    
    lngProcAddress = GetProcAddress(lnghInstanceLib, "InitCommonControlsEx")
    'If we were able to get the function address, then go ahead with API call
    If lngProcAddress <> 0 Then
      'Initialize the common control library
      icex.dwSize = LenB(icex)
      icex.dwICC = ICC_DATE_CLASSES
      API_InitCommonControlsEx icex
      
      ' Create the control
      m_lnghWndDTP = CreateWindowEx(0, DATETIMEPICK_CLASS, "DateTime", lngFlags, _
        m_lngLeft, m_lngTop, m_lngWidth, m_lngHeight, m_hWndParent, 0, _
        App.hInstance, 0)
      
      ' Return True if the control was created
      If m_lnghWndDTP <> 0 Then Create = True
    End If
    
    'Release the library
    FreeLibrary lnghInstanceLib
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, vbExclamation, _
    "Create"
  Resume PROC_EXIT

End Function

Public Sub Destroy()
  ' Comments  : Destroy the window that we created
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : JustACoder
  '
  On Error GoTo PROC_ERR

  DestroyWindow (m_lnghWndDTP)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, vbExclamation, _
    "Destroy"
  Resume PROC_EXIT

End Sub

Public Sub UpdatePosition()
  ' Comments  : This procedure updates the controls position to the left, top
  '             width and height properties. This is useful if you want to
  '             change the size or position of the control after it has
  '             been created
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : JustACoder
  '
  On Error GoTo PROC_ERR

  ' Move the control to the new position
  MoveWindow m_lnghWndDTP, m_lngLeft, m_lngTop, m_lngWidth, m_lngHeight, True
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, vbExclamation, _
    "UpdatePosition"
  Resume PROC_EXIT

End Sub


Public Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    m_bVisible = vNewValue
    ShowWindow m_lnghWndDTP, IIf(m_bVisible, 1, 0)
End Property
