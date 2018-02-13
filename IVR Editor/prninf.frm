VERSION 5.00
Begin VB.Form PrnINFO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   3075
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4290
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "prninf.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   2808
      TabIndex        =   0
      Top             =   2550
      Width           =   1452
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   372
      Left            =   1332
      TabIndex        =   1
      Top             =   2550
      Width           =   1452
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   600
      Picture         =   "prninf.frx":030A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3105
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4320
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "home: http://www.pcsg.net"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   2115
      Width           =   3900
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 1999-2007 PCSG"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   1470
      Width           =   3900
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IVR Workflow Editor Print Preview Sub System V1.0"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   1680
      Width           =   3900
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail: michael.shi@pcsg.net"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   1890
      Width           =   3900
   End
   Begin VB.Label vInf 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1296
      TabIndex        =   2
      ToolTipText     =   "versione esterna"
      Top             =   5652
      Visible         =   0   'False
      Width           =   3288
   End
End
Attribute VB_Name = "PrnINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opzioni di protezione per la chiave del registro di configurazione...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipi di primo livello per la chiave del registro di configurazione...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Stringa Unicode con terminazione Null
Const REG_DWORD = 4                      ' Numero a 32 bit

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSysInfo_Click()
    
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then

    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Convalida l'esistenza di una versione del file a 32 bit conosciuta
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Errore - Impossibile trovare il file...
        Else
            GoTo SysInfoErr
        End If
    ' Errore - Impossibile trovare la voce del registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Non installed or not available", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contatore per il ciclo
    Dim rc As Long                                          ' Codice restituito
    Dim hKey As Long                                        ' Handle a una chiave del registro di configurazione aperta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo di dati di una chiave del registro di configurazione
    Dim tmpVal As String                                    ' Variabile per la memorizzazione temporanea del valore di una chiave del registro di configurazione
    Dim KeyValSize As Long                                  ' Dimensioni della variabile per la chiave del registro di configurazione
    '------------------------------------------------------------------
    ' Apre la chiave del registro sotto KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Apre la chiave del registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestisce gli errori...
    
    tmpVal = String$(1024, 0)                             ' Assegna lo spazio per la variabile
    KeyValSize = 1024                                       ' Definisce le dimensioni della variabile
    
    '---------------------------------------------------------------
    ' Recupera il valore della chiave del registro di configurazione...
    '---------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Recupera/crea il valore della chiave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestisce gli errori
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 aggiunge una stringa con terminazione Null...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Trova Null, estrae dalla stringa
    Else                                                    ' WinNT non aggiunge la terminazione Null alle stringhe...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Non trova Null, estrae solo la stringa
    End If
    '----------------------------------------------------------------
    ' Determina il tipo del valore della chiave per la conversione...
    '----------------------------------------------------------------
    Select Case KeyValType                                  ' Esamina i tipi di dati...
    Case REG_SZ                                             ' Tipo di dati String per la chiave del registro
        KeyVal = tmpVal                                     ' Copia il valore String
    Case REG_DWORD                                          ' Tipo di dati Double Word per la chiave del registro
        For i = Len(tmpVal) To 1 Step -1                    ' Converte ogni bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Crea il valore carattere per carattere.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Converte Double Word in String
    End Select
    
    GetKeyValue = True                                      ' Operazione riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro
    Exit Function                                           ' Esce
    
GetKeyError:      ' Svuota in seguito a un errore...
    KeyVal = ""                                             ' Imposta su una stringa vuota il valore restituito
    GetKeyValue = False                                     ' Operazione non riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

'BackDoor -> a joke
Private Sub Image1_Click()
    Dim iTop  As Integer
    Dim iLeft As Integer
    
    Dim iNewTop  As Integer
    Dim iNewLeft As Integer
    
    
    iTop = Me.Height - Image1.Height
    iLeft = Me.Width - Image1.Width
    
    iNewTop = iTop * Rnd(1)
    iNewLeft = iLeft * Rnd(1)

    If iNewTop < 0 Or iNewTop > iTop Then
        Do Until iNewTop >= 0 And iNewTop <= iTop
            iNewTop = iTop * Rnd(1)
        Loop
    End If

    If iNewLeft < 0 Or iNewLeft > iLeft Then
        Do Until iNewLeft >= 0 And iNewLeft <= iLeft
            iNewLeft = iNewLeft * Rnd(1)
        Loop
    End If

    Image1.Move iNewLeft, iNewTop
    
End Sub
