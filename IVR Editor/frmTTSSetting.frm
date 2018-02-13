VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTTSSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TTS参数设置"
   ClientHeight    =   4485
   ClientLeft      =   3345
   ClientTop       =   3360
   ClientWidth     =   7605
   Icon            =   "frmTTSSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7605
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1913"
   Begin VB.CommandButton ApplyBtn 
      Caption         =   "应用设置"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5760
      TabIndex        =   24
      Tag             =   "1924"
      Top             =   2760
      Width           =   1125
   End
   Begin VB.CheckBox chkShowEvents 
      Caption         =   "详细信息"
      Height          =   195
      Left            =   4320
      TabIndex        =   17
      Tag             =   "1923"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton StopBtn 
      Caption         =   "停止"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4800
      TabIndex        =   2
      Tag             =   "1915"
      Top             =   495
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "扩展属性"
      Height          =   1335
      Left            =   4320
      TabIndex        =   15
      Tag             =   "1925"
      Top             =   1200
      Width           =   3135
      Begin VB.CheckBox chkSpFlagPersistXML 
         Caption         =   "PersistXML"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Width           =   1215
      End
      Begin VB.CheckBox chkSpFlagIsFilename 
         Caption         =   "IsFilename"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkSpFlagAync 
         Caption         =   "FlagsAsync"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkSpFlagPurgeBeforeSpeak 
         Caption         =   "PurgeBeforeSpeak"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   540
         Width           =   1695
      End
      Begin VB.CheckBox chkSpFlagNLPSpeakPunc 
         Caption         =   "NLPSpeakPunc"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkSpFlagIsXML 
         Caption         =   "IsXML"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox AudioOutputCB 
      Height          =   315
      ItemData        =   "frmTTSSetting.frx":030A
      Left            =   840
      List            =   "frmTTSSetting.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2805
      Width           =   3300
   End
   Begin VB.TextBox MainTxtBox 
      Height          =   1095
      HideSelection   =   0   'False
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "1926"
      Top             =   60
      Width           =   4575
   End
   Begin VB.CommandButton ResetBtn 
      Caption         =   "恢复默认"
      Height          =   350
      Left            =   6240
      TabIndex        =   4
      Tag             =   "1917"
      Top             =   495
      Width           =   1125
   End
   Begin VB.TextBox DebugTxtBox 
      BackColor       =   &H80000000&
      Height          =   1080
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   18
      Top             =   3285
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   7080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox FormatCB 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2385
      Width           =   3300
   End
   Begin MSComctlLib.Slider RateSldr 
      Height          =   315
      Left            =   840
      TabIndex        =   8
      ToolTipText     =   "Changes voice playback rate"
      Top             =   1575
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   556
      _Version        =   393216
      LargeChange     =   1
      Min             =   -10
      TickStyle       =   3
   End
   Begin VB.ComboBox VoiceCB 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1275
      Width           =   3300
   End
   Begin VB.CommandButton PauseBtn 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6240
      MaskColor       =   &H00808080&
      TabIndex        =   3
      Tag             =   "1916"
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton SpeakBtn 
      Caption         =   "朗读"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4800
      TabIndex        =   1
      Tag             =   "1914"
      Top             =   60
      Width           =   1125
   End
   Begin MSComctlLib.Slider VolumeSldr 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      ToolTipText     =   "Changes voice playback volume"
      Top             =   1980
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   556
      _Version        =   393216
      Max             =   100
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.Label Label1 
      Caption         =   "音频  输出"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Tag             =   "1922"
      Top             =   2775
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "码率"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Tag             =   "1921"
      Top             =   2415
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "音量"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Tag             =   "1920"
      Top             =   2010
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "速率"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Tag             =   "1919"
      Top             =   1605
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "声音"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Tag             =   "1918"
      Top             =   1305
      Width           =   495
   End
End
Attribute VB_Name = "frmTTSSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare the main SAPI object we are using in this sample
Dim WithEvents Voice As SpVoice
Attribute Voice.VB_VarHelpID = -1

' Speak flags is a combination of bit flags.
Dim m_speakFlags As SpeechVoiceSpeakFlags

' This is the default format we will use.
Dim DefaultFmt As String

' We will disable the output combo box and show this if there's no audio output.
Const NoAudioOutput = "No audio ouput object available"

' m_speaking indicates whether a speak task is in progress
' m_paused indicates whether Voice.Pause is called
Private m_bSpeaking As Boolean
Private m_bPaused As Boolean
Private m_bChanged As Boolean

Private Sub ApplyBtn_Click()
    ' save the setting to ini file
    Dim lv_strTTSFormat As String
    If FormatCB.ListIndex = 0 Then
        lv_strTTSFormat = "SAFT8kHz8BitMono"
    Else
        lv_strTTSFormat = "SAFT8kHz16BitMono"
    End If
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSFormat, lv_strTTSFormat, gSystem.strINI_File)
    
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSRATE, RateSldr.value, gSystem.strINI_File)
    
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVOICE, VoiceCB.Text, gSystem.strINI_File)
    
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVolume, VolumeSldr.value, gSystem.strINI_File)
    
    gSystem.strTTSFormat = lv_strTTSFormat
    gSystem.intTTSRate = RateSldr.value
    gSystem.strTTSVoice = VoiceCB.Text
    gSystem.intTTSVolume = VolumeSldr.value
    
    ApplyBtn.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    MainTxtBox.Text = LoadNationalResString(1926)
    
    ' Creates the voice object first
    Set Voice = New SpVoice
    
    ' Load the voices combo box
    Dim Token As ISpeechObjectToken

    For Each Token In Voice.GetVoices
        VoiceCB.AddItem (Token.GetDescription())
    Next
    
    Dim lv_strVoice As String
    If Trim(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVOICE, gSystem.strINI_File)) <> "" Then
        lv_strVoice = Trim(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVOICE, gSystem.strINI_File))
    End If
    
    If TTSVoiceSetting(lv_strVoice) = False Then
        Message "M151"
        VoiceCB.ListIndex = 0
    End If
    
    'load the format combo box
    AddItemToFmtCB
    
    'set the default format
    If Trim(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSFormat, gSystem.strINI_File)) <> "" Then
        DefaultFmt = Trim(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSFormat, gSystem.strINI_File))
    Else
        DefaultFmt = "SAFT8kHz16BitMono"
    End If
        
    FormatCB.Text = DefaultFmt
    
    ' set rate and volume to the same as the Voice
    Dim lv_iRate As Integer
    If Trim(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSRATE, gSystem.strINI_File)) <> "" Then
        lv_iRate = CInt(Trim(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSRATE, gSystem.strINI_File)))
    Else
        lv_iRate = 0
    End If
    RateSldr.value = lv_iRate
    
    Dim lv_iVolume As Byte
    If Trim(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVolume, gSystem.strINI_File)) <> "" Then
        lv_iVolume = CByte(Val(GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVolume, gSystem.strINI_File)) Mod 256)
    Else
        lv_iVolume = 100
    End If
    VolumeSldr.value = lv_iVolume
    
    ' Load the audio output combo box
    If Voice.GetAudioOutputs.Count > 0 Then
        For Each Token In Voice.GetAudioOutputs
            AudioOutputCB.AddItem (Token.GetDescription)
        Next
    Else
        AudioOutputCB.AddItem NoAudioOutput
        AudioOutputCB.Enabled = False
    End If
    AudioOutputCB.ListIndex = 0
    
    ' init speak flags and sync flag check boxes
    m_speakFlags = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsXML
    chkSpFlagAync.value = Checked
    chkSpFlagPurgeBeforeSpeak.value = Checked
    chkSpFlagIsXML.value = Checked
    
    SetSpeakingState False, False
    
    ApplyBtn.Enabled = False
    
    LoadResStrings Me
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error in TTS initialization: " & vbCrLf & vbCrLf & Err.Description & _
        vbCrLf & vbCrLf & "Shutting down.", vbOKOnly, "TTSApp"
    'Mike added @ 2008-7-8
    Call WriteLogMessage(Err.Number, enu_Error, "Error in TTS initialization !", Err.Description)
    Set Voice = Nothing
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Voice = Nothing
    
    'Michael Added @ 2007-12-5
    If ApplyBtn.Enabled = True Then
        If Message("Q006") = vbYes Then
            Call ApplyBtn_Click
        End If
    End If
End Sub

Private Sub AudioOutputCB_Click()
    On Error GoTo ErrHandler
    
    ' change the output to the selected one
    Set Voice.AudioOutput = Voice.GetAudioOutputs().Item(AudioOutputCB.ListIndex)
    FormatCB_Click
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Set audio output error: ", Err.Description
    'Mike added @ 2008-7-8
    Call WriteLogMessage(Err.Number, enu_Error, "Set TTS audio output error !", Err.Description)
End Sub

Private Sub FormatCB_Click()
    On Error GoTo ErrHandler
    
    m_bChanged = True
    ApplyBtn.Enabled = True
    ' Note: AllowAudioOutputFormatChangesOnNextSet is a hidden property
    Voice.AllowAudioOutputFormatChangesOnNextSet = False
    
    ' The format Type is associated with the selected list item as a long.
    Voice.AudioOutputStream.Format.Type = FormatCB.ItemData(FormatCB.ListIndex)
    
    ' Currently you have to call this so that SAPI picks up the new format.
    Set Voice.AudioOutputStream = Voice.AudioOutputStream
    
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Set format error: ", Err.Description
    'Mike added @ 2008-7-8
    Call WriteLogMessage(Err.Number, enu_Error, "Set TTS output format error !", Err.Description)
End Sub

Private Sub PauseBtn_Click()
    Select Case PauseBtn.Caption
    'LoadResString (1917)
    Case "暂停"
        AddDebugInfo "Pause"
        Voice.Pause
        SetSpeakingState m_bSpeaking, True
    Case "Pause"
        AddDebugInfo "Pause"
        Voice.Pause
        SetSpeakingState m_bSpeaking, True
    'LoadResString(1926)
    Case "继续"
        AddDebugInfo "Resume"
        Voice.Resume
        SetSpeakingState m_bSpeaking, False
    Case "Resume"
        AddDebugInfo "Resume"
        Voice.Resume
        SetSpeakingState m_bSpeaking, False
    End Select
End Sub

Private Sub RateSldr_Scroll()
    m_bChanged = True
    ApplyBtn.Enabled = m_bChanged
    Voice.Rate = RateSldr.value
End Sub

Private Sub ResetBtn_Click()
    'set output to default
    AudioOutputCB.ListIndex = 0
    Set Voice.AudioOutput = Nothing
    
    'use default voice
    VoiceCB.ListIndex = 0
    
    'Format to default
    FormatCB.Text = DefaultFmt
    
    'reset main text field
    MainTxtBox.Text = LoadNationalResString(1926)
    
    'reset volume and rate
    VolumeSldr.value = 100
    VolumeSldr_Scroll
    
    RateSldr.value = 0
    RateSldr_Scroll
    
    ' reset speak flags
    m_speakFlags = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsXML
    chkSpFlagAync.value = Checked
    chkSpFlagPurgeBeforeSpeak.value = Checked
    chkSpFlagIsXML.value = Checked
    chkSpFlagIsFilename.value = Unchecked
    chkSpFlagNLPSpeakPunc.value = Unchecked
    chkSpFlagPersistXML.value = Unchecked
    
    'reset DebugTxtbox text
    DebugTxtBox.Text = Empty
    
    ' if it's paused, call Resume to reset state
    If m_bPaused Then Voice.Resume

    SetSpeakingState False, False
End Sub

Private Sub SpeakBtn_Click()
    On Error GoTo ErrHandler
    AddDebugInfo ("Speak")
    
    ' exit if there's nothing to speak
    If MainTxtBox.Text = "" Then
        Exit Sub
    End If
    
    If Not (m_bPaused And m_bSpeaking) Then
        ' just speak the text with the given flags
        Voice.Speak MainTxtBox.Text, m_speakFlags
    End If
    
    ' Resume if Voice is paused
    If m_bPaused Then Voice.Resume
    
    ' set the state of menu items and buttons
    SetSpeakingState True, False
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Speak Error: ", Err.Description
    'Mike added @ 2008-7-8
    Call WriteLogMessage(Err.Number, enu_Error, "TTS Speak Error !", Err.Description)
    SetSpeakingState False, m_bPaused
End Sub

Private Sub StopBtn_Click()
    On Error GoTo ErrHandler
    AddDebugInfo ("Stop")
    
    Voice.Speak vbNullString, SVSFPurgeBeforeSpeak
    If m_bPaused Then Voice.Resume
    
    SetSpeakingState False, False
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Speak Error: ", Err.Description
    'Mike added @ 2008-7-8
    Call WriteLogMessage(Err.Number, enu_Error, "TTS Speak Error !", Err.Description)
End Sub

Private Sub Voice_AudioLevel(ByVal StreamNumber As Long, _
                             ByVal StreamPosition As Variant, _
                             ByVal AudioLevel As Long)
    ShowEvent "AudioLevel", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "AudioLevel=" & AudioLevel
End Sub

Private Sub Voice_Bookmark(ByVal StreamNumber As Long, _
                           ByVal StreamPosition As Variant, _
                           ByVal Bookmark As String, _
                           ByVal BookmarkId As Long)
    ShowEvent "BookMark", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "Bookmark=" & Bookmark, _
            "BookmarkId=" & BookmarkId
End Sub

Private Sub Voice_EndStream(ByVal StreamNum As Long, ByVal StreamPos As Variant)
    ShowEvent "EndStream", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos

    ' select all text to indicate that we are done
    HighLightSpokenWords 0, Len(MainTxtBox.Text)
    
    ' reset the state of buttons, checkboxes and menu items
    SetSpeakingState False, m_bPaused
End Sub

Private Sub Voice_EnginePrivate(ByVal StreamNumber As Long, _
                                ByVal StreamPosition As Long, _
                                ByVal lParam As Variant)
    ShowEvent "EnginePrivate", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "lParam=" & lParam
End Sub

Private Sub Voice_Phoneme(ByVal StreamNumber As Long, _
                          ByVal StreamPosition As Variant, _
                          ByVal Duration As Long, _
                          ByVal NextPhoneId As Integer, _
                          ByVal Feature As SpeechLib.SpeechVisemeFeature, _
                          ByVal CurrentPhoneId As Integer)
    ShowEvent "Phoneme", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "NextPhoneId=" & NextPhoneId, _
            "Feature=" & Feature, "CurrentPhoneId=" & CurrentPhoneId
End Sub

Private Sub Voice_Sentence(ByVal StreamNum As Long, _
                           ByVal StreamPos As Variant, _
                           ByVal Pos As Long, _
                           ByVal Length As Long)
    ShowEvent "Sentence", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Pos=" & Pos, "Length=" & Length
End Sub

Private Sub Voice_StartStream(ByVal StreamNum As Long, ByVal StreamPos As Variant)
    ShowEvent "StartStream", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos
    
    ' reset the state of buttons, checkboxes and menu items
    SetSpeakingState True, m_bPaused
End Sub

Private Sub Voice_Viseme(ByVal StreamNum As Long, _
                         ByVal StreamPos As Variant, _
                         ByVal Duration As Long, _
                         ByVal VisemeType As SpeechVisemeType, _
                         ByVal Feature As SpeechVisemeFeature, _
                         ByVal VisemeId As Long)
    
    ShowEvent "Viseme", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Duration=" & Duration, "VisemeType=" & VisemeType, _
            "Feature=" & Feature, "VisemeId=" & VisemeId
    
    If VisemeId = 0 Then
        VisemeId = VisemeId + 1
    End If

End Sub

Private Sub Voice_VoiceChange(ByVal StreamNum As Long, _
                              ByVal StreamPos As Variant, _
                              ByVal Token As SpeechLib.ISpeechObjectToken)
    
    ShowEvent "VoiceChange", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Token=" & Token.GetDescription
    
    ' Let's sync up the combo box with the new value
    Dim i As Long
    For i = 0 To VoiceCB.ListCount - 1
        If VoiceCB.List(i) = Token.GetDescription() Then
            VoiceCB.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Sub Voice_Word(ByVal StreamNum As Long, _
                       ByVal StreamPos As Variant, _
                       ByVal Pos As Long, _
                       ByVal Length As Long)
                       
    ShowEvent "Word", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Pos=" & Pos, "Length=" & Length
    
    Debug.Print Pos, Length, MainTxtBox.SelStart, MainTxtBox.SelLength
    
    ' Select the word that's currently being spoken.
    HighLightSpokenWords Pos, Length
End Sub

Private Sub VoiceCB_Click()
    ' change the voice to the selected one
    m_bChanged = True
    ApplyBtn.Enabled = m_bChanged
    Set Voice.Voice = Voice.GetVoices().Item(VoiceCB.ListIndex)
End Sub

Private Sub VolumeSldr_Scroll()
    m_bChanged = True
    ApplyBtn.Enabled = m_bChanged
    Voice.Volume = VolumeSldr.value
End Sub

' The following functions are simply to sync up the speak flags.
Private Sub chkSpFlagAync_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagAync.value, m_speakFlags, SVSFlagsAsync)
End Sub

Private Sub chkSpFlagIsFilename_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagIsFilename.value, m_speakFlags, SVSFIsFilename)
End Sub

Private Sub chkSpFlagIsXML_Click()
    If chkSpFlagIsXML.value = 0 Then
        ' clear SVSFIsXML bit and set SVSFIsNotXML bit
        m_speakFlags = m_speakFlags And Not SVSFIsXML
        m_speakFlags = m_speakFlags Or SVSFIsNotXML
    Else
        ' clear SVSFIsNotXML bit and set SVSFIsXML bit
        m_speakFlags = m_speakFlags And Not SVSFIsNotXML
        m_speakFlags = m_speakFlags Or SVSFIsXML
    End If
End Sub

Private Sub chkSpFlagNLPSpeakPunc_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagNLPSpeakPunc.value, m_speakFlags, SVSFNLPSpeakPunc)
End Sub

Private Sub chkSpFlagPersistXML_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagPersistXML.value, m_speakFlags, SVSFPersistXML)
End Sub

Private Sub chkSpFlagPurgeBeforeSpeak_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagPurgeBeforeSpeak.value, m_speakFlags, SVSFPurgeBeforeSpeak)
End Sub

Private Sub AddFmts(ByRef name As String, ByVal fmt As SpeechAudioFormatType)
    Dim Index As String
    
    ' get the count of existing list so that we are adding to the bottom of the list
    Index = FormatCB.ListCount
    
    ' add the name to the list box and associate the format type with the item
    FormatCB.AddItem name, Index
    FormatCB.ItemData(Index) = fmt
End Sub

Private Sub AddItemToFmtCB()
    AddFmts "SAFT8kHz8BitMono", SAFT8kHz8BitMono
    AddFmts "SAFT8kHz16BitMono", SAFT8kHz16BitMono
End Sub

Private Sub AddDebugInfo(DebugStr As String, Optional Error As String = Empty)
    ' This function adds debug string to the info window.
    If Len(DebugTxtBox.Text) > 64000 Then
        Debug.Print "Too much stuff in the debug window. Remove first 10K chars"
        DebugTxtBox.SelStart = 0
        DebugTxtBox.SelLength = 10240
        DebugTxtBox.SelText = ""
    End If
    
    ' append the string to the DebugTxtBox text box and add a newline
    DebugTxtBox.SelStart = Len(DebugTxtBox.Text)
    DebugTxtBox.SelText = DebugStr & Error & vbCrLf
End Sub

Private Sub ShowEvent(ParamArray strArray())
    ' we will only show the events if the ShowEvents box is checked
    If chkShowEvents.value = Checked Then
        Dim strText As String
        strText = Join(strArray, ", ")
        AddDebugInfo "  Event: " & strText
    End If
End Sub

Private Sub HighLightSpokenWords(ByVal Pos As Long, ByVal Length As Long)
    On Error GoTo ErrHandler
    If chkSpFlagIsFilename.value = Unchecked Then
        MainTxtBox.SelStart = Pos
        MainTxtBox.SelLength = Length
    End If
    
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Failed to high light words. This may be caused by too many charaters in the main text box."
    'Mike Added @ 2008-7-8
    Call WriteLogMessage(Err.Number, enu_Warnning, "Failed to high light words. This may be caused by too many charaters in the main text box.", Err.Description)
End Sub

Private Function SetOrClearFlag(ByVal cond As Long, _
                                ByVal base As Long, _
                                ByVal flag As Long) As Long
    
    If cond = 0 Then
        ' the condition is false, clear the flag
        SetOrClearFlag = base And Not flag
    Else
        ' the condition is false, set the flag
        SetOrClearFlag = base Or flag
    End If
End Function

Private Sub SetSpeakingState(ByVal bSpeaking As Boolean, ByVal bPaused As Boolean)
    SpeakBtn.Enabled = True
    
    StopBtn.Enabled = bSpeaking
    PauseBtn.Enabled = bSpeaking
    
    If bPaused Then
        PauseBtn.Caption = LoadNationalResString(1927)
    Else
        PauseBtn.Caption = LoadNationalResString(1916)
    End If
    
    m_bSpeaking = bSpeaking
    m_bPaused = bPaused
End Sub

Private Function TTSVoiceSetting(strVoice As String) As Boolean
On Error GoTo ErrorHandle
        VoiceCB.Text = strVoice
        TTSVoiceSetting = True
        Exit Function
ErrorHandle:
    TTSVoiceSetting = False
    On Error GoTo 0
End Function
