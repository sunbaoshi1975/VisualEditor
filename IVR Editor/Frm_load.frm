VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Frm_load 
   Caption         =   "流程加载"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2250
   Icon            =   "Frm_load.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2250
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1393"
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   -60
   End
   Begin MSWinsockLib.Winsock Winskget 
      Left            =   1230
      Top             =   -60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7777
   End
   Begin VB.CommandButton Comm_exit 
      Caption         =   "&Q退出"
      Height          =   315
      Left            =   1260
      TabIndex        =   4
      Tag             =   "1144"
      Top             =   930
      Width           =   885
   End
   Begin VB.CommandButton Comm_load 
      Caption         =   "&L加载"
      Default         =   -1  'True
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Tag             =   "1395"
      Top             =   930
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "加载流程"
      ForeColor       =   &H00FF8080&
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Tag             =   "1394"
      Top             =   120
      Width           =   2025
      Begin VB.TextBox Txt_Group 
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
         Left            =   720
         TabIndex        =   5
         Top             =   210
         Width           =   495
      End
      Begin VB.TextBox Txt_pid 
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
         Left            =   1200
         TabIndex        =   2
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "流程ID"
         Height          =   225
         Left            =   60
         TabIndex        =   1
         Tag             =   "1363"
         Top             =   270
         Width           =   615
      End
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   0
      Top             =   -60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   7777
      LocalPort       =   6666
   End
End
Attribute VB_Name = "Frm_load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prop As Integer
Dim mFormTitle As String

Private Sub Comm_exit_Click()
    sckMain.Close
    Winskget.Close
    Unload Me
End Sub

Private Sub Comm_load_Click()
  prop = 2
' Call Mdlsend.tcpSendLoadMsg
  While sckMain.State <> sckClosed
        sckMain.Close
  Wend
       sckMain.Connect
' Comm_load = False
End Sub

Private Sub Form_Load()
    prop = 1
    mFormTitle = Me.Caption
    sckMain.Close
    sckMain.RemoteHost = gSystem.strServerIP
    sckMain.RemotePort = gSystem.intServerPort
    sckMain.Connect
    Winskget.LocalPort = 1001
    Winskget.Listen
    LoadResStrings Me
End Sub

Private Sub sckMain_Connect()
Select Case prop
Case 1
  Me.Caption = LoadNationalResString(1396)
  Mdlsend.tcpLogonLoadMsg
Case 2
  Debug.Print Winskget.State
  Me.Caption = LoadNationalResString(1397)
  Mdlsend.tcpSendLoadMsg
End Select

End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Me.Caption = mFormTitle & " - " & Description
End Sub

Private Sub sckMain_SendComplete()
    Comm_load.Enabled = True
End Sub

'Private Sub sckMain_Close()
'    Comm_load.Enabled = True
'    While sckMain.State <> sckClosed
'        sckMain.Close
'    Wend
'
'    If Comm_exit.Enabled = False Then
'        sckMain.Connect
'
'    End If
'
'End Sub

Private Sub Winskget_ConnectionRequest(ByVal requestID As Long)
If Winskget.State <> 0 Then Winskget.Close
    Winskget.Accept requestID
End Sub
Private Sub Winskget_DataArrival(ByVal bytesTotal As Long)
On Error GoTo BackDoor
Dim num As Integer
Dim strdata As Variant
Dim c As Byte
Dim d As Integer
Dim s As Byte
Dim lv_Msg(Def_PKGLENGTH) As Byte
    Dim lv_Str As String
    Dim lv_Cmd As Byte
    Dim lv_PKG As SCtiMsi_Package
    Dim lv_Len As Byte
    Dim cnstate As Integer
    Dim lv_lvp As Integer
 Timer1.Enabled = False
 Winskget.GetData lv_Str, vbString

 While Winskget.State <> sckClosed
       Winskget.Close
 Wend
 Winskget.Listen
 'Debug.Print winskget.State
   lv_Len = IIf(bytesTotal > Def_PKGLENGTH, Def_PKGLENGTH, bytesTotal)
    
    If lv_Len > Len(lv_Str) Then
        lv_Len = Len(lv_Str)
    End If
    For lv_lvp = 0 To lv_Len - 1
        lv_Msg(lv_lvp) = AscB(Mid(lv_Str, lv_lvp + 1, 1))
        Debug.Print lv_Msg(lv_lvp)
    Next
    CopyMemory lv_PKG.command, lv_Msg(6), 1
    CopyMemory lv_PKG.bytData, lv_Msg(11), 1
    CopyMemory c, lv_Msg(6), 1
    CopyMemory s, lv_Msg(10), 1
    CopyMemory d, lv_Msg(11), 2
    Select Case c
 Case 1:  'LONON RESULT
   If s = "1" Then   '成功=1
   cnstate = 1
   Me.Caption = LoadNationalResString(1398)

   Else
   cnstate = 0
   Me.Caption = LoadNationalResString(1399)
   End If
 Case 2:  'DATA RESULT
    If d = "0" Then    '成功=1
   ' CnState = 1
    Me.Caption = LoadNationalResString(1400)
 
    Else
    'CnState = 0
    num = (Int(Txt_Group) + 1) * 10 + 100 - 5
''    M_Cn.Execute "update tbparameter set dvalue='" & Txt_pid.Text & "' where id=" & num
    Me.Caption = LoadNationalResString(1401)
    End If
 

'   CnState = 2
'   Me.Caption = gFormCaption & " -" & "  无法连接到消息机"

 Case Else
    'Debug.Print "unknown msg!"
    MsgBox "unknown msg!"
 End Select
 Exit Sub
BackDoor:
    On Error GoTo 0
    While Winskget.State <> sckListening
        While Winskget.State <> sckClosed
            Winskget.Close
        Wend
        Winskget.Listen
    Wend

End Sub
