VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmResourceItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "资源编辑"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "frmResourceItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1742"
   Begin VB.CommandButton cmdBrowse 
      Height          =   345
      Index           =   1
      Left            =   8190
      Picture         =   "frmResourceItem.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "1521"
      Top             =   540
      Width           =   345
   End
   Begin VB.ComboBox cboLType 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1635
   End
   Begin VB.ComboBox cboRType 
      Height          =   315
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtModifyTime 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6180
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Width           =   2355
   End
   Begin VB.TextBox txtCreateTime 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   2355
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      MaxLength       =   200
      TabIndex        =   5
      Top             =   930
      Width           =   7335
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   3
      Top             =   540
      Width           =   6945
   End
   Begin VB.TextBox txtRID 
      BackColor       =   &H008080FF&
      Height          =   315
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   150
      Width           =   1155
   End
   Begin VB.TextBox txtNote 
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   1200
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1290
      Width           =   7335
   End
   Begin VB.CommandButton exit_Command 
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   7380
      TabIndex        =   11
      Tag             =   "1144"
      Top             =   2940
      Width           =   1065
   End
   Begin VB.CommandButton CommandCreate 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   315
      Left            =   6180
      TabIndex        =   10
      Tag             =   "1372"
      Top             =   2940
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar tlbNavigation 
      Height          =   390
      Left            =   1200
      TabIndex        =   9
      Top             =   2880
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "keyFirst"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "keyPrev"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "keyNext"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "keyLast"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyListen"
            Object.ToolTipText     =   "1523"
            ImageKey        =   "Play"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyStop"
            Object.ToolTipText     =   "1101"
            ImageKey        =   "Stop"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyRecord"
            Object.ToolTipText     =   "1717"
            ImageKey        =   "Record2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyTTS"
            Object.ToolTipText     =   "1723"
            ImageKey        =   "Record"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyTools"
            Object.ToolTipText     =   "1722"
            ImageKey        =   "WAVE2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":074C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":0A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":0B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":0CBC
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":0E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":1130
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":144A
            Key             =   "TTS"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":2BDC
            Key             =   "Record"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":48E6
            Key             =   "keyTools"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":60A8
            Key             =   "Record2"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":63C2
            Key             =   "WAVE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceItem.frx":66DC
            Key             =   "WAVE2"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "资源类型："
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5430
      TabIndex        =   19
      Tag             =   "1459"
      Top             =   150
      Width           =   900
   End
   Begin VB.Label lblLanguage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "资源语言："
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2565
      TabIndex        =   18
      Tag             =   "1445"
      Top             =   150
      Width           =   900
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "修改时间："
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   106
      Left            =   5055
      TabIndex        =   17
      Tag             =   "1370"
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "创建时间："
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   105
      Left            =   90
      TabIndex        =   16
      Tag             =   "1369"
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "资源描叙："
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   104
      Left            =   90
      TabIndex        =   15
      Tag             =   "1450"
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "资源路径："
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   103
      Left            =   90
      TabIndex        =   14
      Tag             =   "1451"
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "资源编号："
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   100
      Left            =   90
      TabIndex        =   13
      Tag             =   "1449"
      Top             =   180
      Width           =   900
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "资源备注："
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   107
      Left            =   90
      TabIndex        =   12
      Tag             =   "1959"
      Top             =   1320
      Width           =   900
   End
End
Attribute VB_Name = "frmResourceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 数据是否改变
Public m_blnDataChanged As Boolean

Public strNamePath As String
Public tempNamePath As String

Private Sub cboLType_Click()
    m_blnDataChanged = True
End Sub

Private Sub cboRType_Click()
    m_blnDataChanged = True
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
    
    ' 查看全路径
    ' 检查路径是否存在？
    ' Michael Note : StatusLogFolder have not define !!! ...
'    If StatusLogFolder(gSystem.strPath_SysVox) = False Then
'        'MsgBox LoadResString(377), vbOKOnly + vbInformation, LoadResString(371)
'        MsgBox LoadResString(377), vbOKOnly + vbInformation, MDIMain.Caption
'        Exit Sub
'    End If
    
    frmSetPath.Show 1
    
End Sub
' 保存资源条目
'
Private Sub CommandCreate_Click()
'On Error Resume Next
On Error GoTo Catch

    Dim lv_nLID As Integer
    Dim lv_nRType As Integer
    Dim itmX As ListItem                    ' ListItem 变量
    Dim lv_AbsPos As Integer
    Dim strConSQL As String
    
    'Michael Modified @ 2007-12-5
    '资源编号不能为0或空
    If CLng(txtRID.Text) = 0 Then
       Message ("E138")
       txtRID.SetFocus
       Exit Sub
    End If
    
    If m_blnDataChanged Then
    
        '' Gother Data
        If cboLType.ListIndex < 0 Then
            lv_nLID = 0
        Else
            lv_nLID = CByte(cboLType.ItemData(cboLType.ListIndex))
        End If
        
        If cboRType.ListIndex < 0 Then
            lv_nRType = 0
        Else
            lv_nRType = CByte(cboRType.ItemData(cboRType.ListIndex))
        End If
        
        If txtRID.Enabled Then
            '' Insert
             '检查资源编号不会发生重复
             'If Not CheckResID Then Exit Sub
             
'             strConSQL = "Insert into tbResource (P_ID,R_ID,L_ID,R_Type,R_Path,R_Description,R_Note," & _
'                         "CreateTime,ModifyTime) values ('" & frmResourceList.txtResID & "','" & txtRID & "','" & lv_nLID & "'," & _
'                         "'" & lv_nRType & "','" & txtPath & "','" & txtDescription & "','" & txtNote & "','" & txtCreateTime & "'" & _
'                         ",'" & txtModifyTime & "')"
'             frmResourceList.FillRSListView strConSQL
'             frmResourceList.FillRSListView gstrSQL
            
            'Mike modified @2008-6-27, use recordset.new instead of SQL "Insert..."
            lv_AbsPos = 0
            If gCallFlow.AddResourceItem(lv_AbsPos, Val(txtRID), lv_nLID, lv_nRType, txtPath, txtDescription, txtNote) Then
                Set itmX = frmResourceList.lstResource.ListItems.Add(, , Right("00000" & Trim(txtRID), 5))
                itmX.Key = "k" & CStr(lv_AbsPos)
                itmX.Tag = lv_nRType
                itmX.SubItems(1) = Trim(Str(lv_nLID))
                itmX.SubItems(2) = gRID(lv_nRType).Caption
                itmX.SubItems(3) = Trim(txtDescription)
                itmX.SubItems(4) = Trim(txtPath)
                itmX.SubItems(5) = Format(Now, "yyyy-mm-dd hh:nn:ss")
                itmX.SubItems(6) = Format(Now, "yyyy-mm-dd hh:nn:ss")
                itmX.SubItems(7) = Trim(txtNote)
                itmX.Selected = True
                itmX.EnsureVisible
            Else
                Message ("E139")
                txtRID.SetFocus
                Exit Sub
            End If

        Else
            '' Update
'            strConSQL = "Update tbResource SET R_Path='" & txtPath & "',R_Description='" & txtDescription & "',R_Note='" & txtNote & "'," & _
'                    "Createtime='" & txtCreateTime & "',ModifyTime='" & txtModifyTime & "',L_ID='" & lv_nLID & "',R_Type='" & lv_nRType & "'" & _
'                    "where P_ID='" & frmResourceList.txtResID & "' and R_ID='" & txtRID & "'"
'
'            frmResourceList.FillRSListView strConSQL
'            frmResourceList.FillRSListView gstrSQL
            
            'Mike modified @ 2008-6-27, use recordset instead of SQL "Update ..."
            If gCallFlow.UpdateResourceItem(frmResourceList.lstResource.SelectedItem.Key, _
                    Val(txtRID), lv_nLID, lv_nRType, txtPath, txtDescription, txtNote) Then
                frmResourceList.lstResource.SelectedItem.SubItems(1) = Trim(Str(lv_nLID))
                frmResourceList.lstResource.SelectedItem.SubItems(2) = gRID(lv_nRType).Caption
                frmResourceList.lstResource.SelectedItem.Tag = lv_nRType
                frmResourceList.lstResource.SelectedItem.SubItems(3) = Trim(txtDescription)
                frmResourceList.lstResource.SelectedItem.SubItems(4) = Trim(txtPath)
                frmResourceList.lstResource.SelectedItem.SubItems(6) = Format(Now, "yyyy-mm-dd hh:nn:ss")
                frmResourceList.lstResource.SelectedItem.SubItems(7) = Trim(txtNote)
            End If
           
        End If
        m_blnDataChanged = False
    End If
    
    '' Sun added 2008-02-04
    'gSystem.crlCurItem.Text = txtRID
    
    Unload Me
Exit Sub

Catch:
    'Call WriteLogMsg(Err.Description, Err.Number, Me.name & ".CommandCreate_Click()")
    On Error Resume Next
End Sub

Private Sub exit_Command_Click()
    'frmResourceList.strSelID = ""
    Unload Me
End Sub

Private Sub Form_Load()

    InitLanguageList
    InitResourceTypeList
    
    cboLType.Enabled = (gSystem.intCurStep < 0)
    cboRType.Enabled = (gSystem.intCurStep < 0)
    
    LoadResStrings Me
    m_blnDataChanged = False
    
End Sub

Private Sub InitLanguageList()
    Dim lv_loop As Integer

    cboLType.Clear
    'Michael Added @ Jul,20,07, Language must bigger than 0
    If Node0_Data1.Languages <= 0 Then Node0_Data1.Languages = 1
    
    If Node0_Data1.Languages > 0 Then
        For lv_loop = 0 To Node0_Data1.Languages - 1
            cboLType.AddItem LoadNationalResString(1445) & Str(lv_loop)
            cboLType.ItemData(cboLType.ListCount - 1) = lv_loop
        Next
    End If
End Sub

Private Sub InitResourceTypeList()
    Dim lv_loop As Integer

    cboRType.Clear
    For lv_loop = 0 To 4
        cboRType.AddItem gRID(lv_loop).Caption
        cboRType.ItemData(cboRType.ListCount - 1) = lv_loop
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_blnDataChanged Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub txtDescription_Change()
    m_blnDataChanged = True
End Sub

Private Sub txtDescription_GotFocus()
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription)
End Sub

Private Sub txtDescription_LostFocus()
    Call ToolbarSetting
End Sub

Private Sub txtNote_Change()
    m_blnDataChanged = True
End Sub

Private Sub txtNote_GotFocus()
    txtNote.SelStart = 0
    txtNote.SelLength = Len(txtNote)
End Sub

Private Sub txtPath_Change()
    m_blnDataChanged = True
End Sub

Private Sub txtPath_GotFocus()
    txtPath.SelStart = 0
    txtPath.SelLength = Len(txtPath)
End Sub

Private Sub txtPath_LostFocus()
    Call ToolbarSetting
End Sub

Private Sub txtRID_Change()
    m_blnDataChanged = True
End Sub

Private Sub txtRID_GotFocus()
    txtRID.SelStart = 0
    txtRID.SelLength = Len(txtRID)
End Sub

'**** Michael Commented : Use gcallflow.IsNewResourceItem insteand @ 2008-6-30
'Michael Added @ Sep,4,07
'Michael Modified @ 2007-11-01 -> 修改资源编号
Private Function CheckResID() As Boolean
    Dim iLoop As Integer
    Dim iItemCount As Integer
    iItemCount = frmResourceList.lstResource.ListItems.Count
    CheckResID = True
    
    If iItemCount <> 0 Then
        For iLoop = 1 To iItemCount
            If txtRID.Text = frmResourceList.lstResource.ListItems(iLoop).Text Then
                CheckResID = False
                'Michael Modified @ 2007-11-26
                'MsgBox "资源编号重复", vbOKOnly, "资源冲突"
                Message ("E139")
                txtRID.SetFocus
                Exit Function
            End If
        Next iLoop
    End If
End Function

'Michael Added @ 2007-11-01
'启用工具栏按钮
Private Sub tlbNavigation_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo 0
    Dim lv_strFName As String
    Dim lv_intLID As Integer
    Dim lv_Item As MSComctlLib.ListItem
    Dim lv_Index As Integer

    If Button.Key = "keyListen" Then
        If txtRID.Text <> "" Then
            lv_strFName = Trim(txtRID.Text)

            If LCase(Right(txtPath.Text, 3)) = "vox" Then
                frmMain.VOXPlayer.PlayVOXFile gSystem.strPath_SysVox & Trim(txtPath.Text)
            ElseIf LCase(Right(txtPath, 3)) = "wav" Then
                'Play wav file
                Dim strWAVPath As String
                strWAVPath = gSystem.strPath_SysVox & Trim(txtPath.Text)
                sndPlaySound 0, 0
                sndPlaySound strWAVPath, &H1 Or &H2
            End If
        End If

    ElseIf Button.Key = "keyStop" Then
        StopSound

    ElseIf Button.Key = "keyRecord" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If txtPath <> "" And (LCase(Right(txtPath, 3)) = "vox" Or LCase(Right(txtPath, 3)) = "wav") Then
            frmRecordVoice.txtScript = txtDescription.Text
            frmRecordVoice.txtFilePath = gSystem.strPath_SysVox & txtPath.Text
            'check exists file
            If fso.FileExists(frmRecordVoice.txtFilePath) = True Then
                If MsgBox(LoadResString(1928) & txtPath & LoadResString(1929), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) = vbYes Then
                    frmRecordVoice.Show vbModal
                End If
            Else
                frmRecordVoice.Show vbModal
            End If
        End If
        Set fso = Nothing

    ElseIf Button.Key = "keyTTS" Then
        Call TTSConvert

    ElseIf Button.Key = "keyTools" Then
        On Error Resume Next
        If txtPath.Text <> "" Then
            lv_strFName = gSystem.strPath_SysVox & Trim(txtPath.Text)
            Call ShellExecute(0, "open", gSystem.strVoiceEditorPath, lv_strFName, vbNullString, 1)
        End If
    End If
End Sub

'Michael Added @ 2007-11-01
'设置工具栏按钮状态
Public Sub ToolbarSetting()
    'Disable Record & TTS
    If txtDescription = "" Or (LCase(Right(txtPath.Text, 3)) <> "vox" And LCase(Right(txtPath.Text, 3)) <> "wav") Then
        tlbNavigation.Buttons(6).Enabled = False
        tlbNavigation.Buttons(7).Enabled = False
        tlbNavigation.Buttons(10).Enabled = False
        tlbNavigation.Buttons(11).Enabled = False
        tlbNavigation.Buttons(13).Enabled = False
    Else
        tlbNavigation.Buttons(6).Enabled = True
        tlbNavigation.Buttons(7).Enabled = True
        tlbNavigation.Buttons(10).Enabled = True
        tlbNavigation.Buttons(11).Enabled = True
        tlbNavigation.Buttons(13).Enabled = True
    End If
End Sub

'TTS convert the text to speech
Private Sub TTSConvert()
On Error GoTo VoiceDebug

    Dim Voice As SpVoice
    Set Voice = New SpVoice
    
    Dim Token As SpeechLib.SpObjectToken
    Set Token = New SpeechLib.SpObjectToken
    
    'Create a new wave stream
    Dim cpFileStream As New SpFileStream
    Dim lv_iCount As Integer
    Dim lv_iTTSCount As Integer
    lv_iTTSCount = 0
    
    For Each Token In Voice.GetVoices
        If gSystem.strTTSVoice = Token.GetDescription Then
            gSystem.intTTSVoice = lv_iTTSCount
            Exit For
        Else
            lv_iTTSCount = lv_iTTSCount + 1
        End If
    Next
    
    If txtPath <> "" And txtDescription <> "" Then
        If txtRID.Enabled Then
            For lv_iCount = 1 To frmResourceList.lstResource.ListItems.Count
                'deal with same resource description
                If frmResourceList.lstResource.ListItems(lv_iCount).SubItems(3) = txtDescription.Text Then
                    If MsgBox("" & frmResourceList.lstResource.ListItems(lv_iCount).Text & LoadResString(1931) & txtRID.Text & LoadResString(1933), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) <> vbYes Then
                        Exit Sub
                    End If
                End If
                
                'deal with same resource name
                If frmResourceList.lstResource.ListItems(lv_iCount).SubItems(4) = txtPath.Text Then
                    If MsgBox("" & frmResourceList.lstResource.ListItems(lv_iCount).Text & LoadResString(1932) & txtRID.Text & LoadResString(1933), vbYesNo + vbQuestion + vbDefaultButton2, App.Title & LoadResString(1934) & frmResourceList.lstResource.ListItems(lv_iCount).Text & LoadResString(1935)) <> vbYes Then
                        Exit Sub
                    End If
                End If
            Next lv_iCount
        End If
                        
        strNamePath = gSystem.strPath_SysVox & Trim(txtPath)
        
        Dim fsoTTS As Object
        Set fsoTTS = CreateObject("Scripting.FileSystemObject")
        'check exists file
        If fsoTTS.FileExists(strNamePath) Then
            If MsgBox(LoadResString(1928) & txtPath & LoadResString(1929), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
        Set fsoTTS = Nothing
        
        'Voice File Type - VOX
        If LCase(Right(strNamePath, 3)) = "vox" Then
            tempNamePath = Left(strNamePath, Len(strNamePath) - 3) & "wav"
             'Set audio format -> Michael Modified @ 2007-11-29
            Set Voice.Voice = Voice.GetVoices().Item(gSystem.intTTSVoice)
            Voice.Volume = gSystem.intTTSVolume
            Voice.Rate = gSystem.intTTSRate
            If gSystem.strTTSFormat = "SAFT8kHz8BitMono" Then
                cpFileStream.Format.Type = SAFT8kHz8BitMono
            ElseIf gSystem.strTTSFormat = "SAFT8kHz16BitMono" Then
                cpFileStream.Format.Type = SAFT8kHz16BitMono
            End If
            
            cpFileStream.Open tempNamePath, SSFMCreateForWrite, False
            Set Voice.AudioOutputStream = cpFileStream
            Voice.Speak txtDescription.Text, SVSFDefault
            'Close the file
            cpFileStream.Close
            Set cpFileStream = Nothing
            'Reset the Voice object's output to 'Nothing'.
            Set Voice.AudioOutputStream = Nothing
            'convert 2 vox file
            Call frmResourceList.ctlConvert.WaveFile2VOX(ByVal tempNamePath, ByVal strNamePath)
            
        'Voice File Type - WAV
        ElseIf LCase(Right(strNamePath, 3)) = "wav" Then
            'Set audio format
            'Michael Modified @ 2007-11-28
            Set Voice.Voice = Voice.GetVoices().Item(gSystem.intTTSVoice)
            Voice.Volume = gSystem.intTTSVolume
            Voice.Rate = gSystem.intTTSRate
            If gSystem.strTTSFormat = "SAFT8kHz8BitMono" Then
                cpFileStream.Format.Type = SAFT8kHz8BitMono
            ElseIf gSystem.strTTSFormat = "SAFT8kHz16BitMono" Then
                cpFileStream.Format.Type = SAFT8kHz16BitMono
            End If
            '******************************************
            cpFileStream.Open strNamePath, SSFMCreateForWrite, False
            Set Voice.AudioOutputStream = cpFileStream
            Voice.Speak txtDescription.Text, SVSFDefault
            'Close the file
            cpFileStream.Close
            Set cpFileStream = Nothing
            'Reset the Voice object's output to 'Nothing'.
            Set Voice.AudioOutputStream = Nothing
            strNamePath = ""
        'Other file Type or invaid path and file name ...
        Else
            'Michael Modified @ 2007-11-26
            Message ("E140")
            Exit Sub
        End If
    End If
    Set Voice = Nothing
    
Exit Sub

VoiceDebug:
    MsgBox "TTS Error" & Err.Description, vbOKOnly + vbExclamation, App.Title
    Call WriteLogMessage(Err.Number, enu_Error, "TTS Convert error, FileName:" & strNamePath, Err.Description)
    'On Error GoTo 0
End Sub

