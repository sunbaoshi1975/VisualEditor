VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FB712C5B-9403-4F42-B022-73ABA0F01A72}#1.0#0"; "VOXExpCtrl.ocx"
Begin VB.Form frmResourceList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "资源列表"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "frmResourceList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1446"
   Begin VB.Timer TimerConvert 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   5400
   End
   Begin VOXEXPCTRLLib.VOXExpCtrl ctlConvert 
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin VB.CommandButton exit_Command 
      Caption         =   "退出(&E)"
      Height          =   345
      Left            =   9540
      TabIndex        =   7
      Tag             =   "1144"
      Top             =   5430
      Width           =   1065
   End
   Begin VB.CommandButton CommandCreate 
      Caption         =   "确定(&C)"
      Default         =   -1  'True
      Height          =   345
      Left            =   8400
      TabIndex        =   6
      Tag             =   "1372"
      Top             =   5430
      Width           =   1065
   End
   Begin MSComctlLib.ListView lstResource 
      Height          =   4845
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   8546
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ID"
         Object.Tag             =   "1449"
         Text            =   "资源编号"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "LID"
         Object.Tag             =   "1744"
         Text            =   "语言"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Type"
         Object.Tag             =   "1745"
         Text            =   "类型"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Description"
         Object.Tag             =   "1450"
         Text            =   "资源描述"
         Object.Width           =   11465
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Path"
         Object.Tag             =   "1451"
         Text            =   "资源路径"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "CreateTime"
         Object.Tag             =   "1369"
         Text            =   "创建时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "ModifyTime"
         Object.Tag             =   "1370"
         Text            =   "修改时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Notes"
         Object.Tag             =   "1746"
         Text            =   "备注"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtResName 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   7035
   End
   Begin VB.TextBox txtResID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   2850
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
            Picture         =   "frmResourceList.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":28FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":2A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":2BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":2D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":3186
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":34A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":4C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":693C
            Key             =   "keyTools"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":80FE
            Key             =   "Record2"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":8418
            Key             =   "WAVE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":8732
            Key             =   "WAVE2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbNavigation 
      Height          =   390
      Left            =   90
      TabIndex        =   5
      Top             =   5400
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
            Key             =   "keyFirst"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyPrev"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyNext"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyLast"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyListen"
            Object.ToolTipText     =   "1523"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
            Enabled         =   0   'False
            Key             =   "keyRecord"
            Object.ToolTipText     =   "1717"
            ImageKey        =   "Record2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyTTS"
            Object.ToolTipText     =   "1723"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "keyTools"
            Object.ToolTipText     =   "1722"
            ImageKey        =   "WAVE2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFile 
      Left            =   0
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":8A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":8BA8
            Key             =   "Reset"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":8FBC
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":9556
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":AAE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":B082
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":B61C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":BBB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":BF41
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":C2E7
            Key             =   "Submit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":C881
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":CE1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":D17C
            Key             =   "Update"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":D716
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":DCB0
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":E42A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":E586
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":EB22
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":EC7E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":F018
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":F5B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":FA90
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResourceList.frx":FBEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   390
      Left            =   4050
      TabIndex        =   8
      Top             =   5400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "imgFile"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyCopy"
            Object.ToolTipText     =   "1724"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyInsert"
            Object.ToolTipText     =   "1725"
            ImageKey        =   "Insert"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyUpdate"
            Object.ToolTipText     =   "1726"
            ImageKey        =   "Update"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyDelete"
            Object.ToolTipText     =   "1727"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keySearch"
            Object.ToolTipText     =   "1728"
            ImageKey        =   "Search"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KeyReflash"
            Object.ToolTipText     =   "1819"
            ImageKey        =   "Reset"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "资源项目名称"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Tag             =   "1448"
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "资源项目ID"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Tag             =   "1447"
      Top             =   150
      Width           =   885
   End
End
Attribute VB_Name = "frmResourceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////
'Modify Date: July -9 -07
'Modify Version : V1.1
'Modified By: Michael
'Contect : ...

'///////////////////////////////////////////////////////////////////////////////

Option Explicit
Dim strNamePath As String, tempNamePath As String

Private Sub CommandCreate_Click()
    Dim lv_NewID As Integer
    
    If Not lstResource.SelectedItem Is Nothing Then
        If Not gSystem.crlCurItem Is Nothing Then
            lv_NewID = lstResource.SelectedItem.Text
            If lv_NewID <> Val(gSystem.crlCurItem.Text) Then
                gSystem.crlCurItem.Text = Trim(Str(lv_NewID))
            End If
        End If
    End If

    Unload Me
End Sub

Private Sub exit_Command_Click()
    gbCallFromPro = 0
    gbSearchFlag = 0
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Dim lv_strName As String
    If gSystem.intCurStep < 2 Then
        tlbNavigation.Buttons("keyListen").Enabled = True
        tlbNavigation.Buttons("keyStop").Enabled = True
        tlbNavigation.Buttons("keyRecord").Enabled = True
        tlbNavigation.Buttons("keyTTS").Enabled = True
        tlbNavigation.Buttons("keyTools").Enabled = (gSystem.strVoiceEditorPath <> "")
    Else
        tlbNavigation.Buttons("keyListen").Enabled = False
        tlbNavigation.Buttons("keyStop").Enabled = False
        tlbNavigation.Buttons("keyRecord").Enabled = False
        tlbNavigation.Buttons("keyTTS").Enabled = False
        tlbNavigation.Buttons("keyTools").Enabled = False
    End If
    
    '' Title
    If gSystem.intCurStep >= 0 Then
        Me.Caption = LoadNationalResString(1663) & gRID(gSystem.intCurStep).Caption
    Else
        Me.Caption = LoadNationalResString(1663)
    End If
    
    '' Resource Project
    'Michael Modify @ Jul,18,07 for call from Project List
    If gbCallFromPro = 1 Then
        txtResID = frmProjectList.lstProject.SelectedItem.Text
        txtResName = frmProjectList.lstProject.SelectedItem.SubItems(1)
    Else
        txtResID = Trim(Str(gCallFlow.ResourceID))
        txtResName = gCallFlow.ResourceName
    End If
    
    '' Add Items
    'Michael Modified @ Jul,18,07 for call from Project List
    If gbCallFromPro = 1 Then
        FillRSListView gstrSQL
    Else
        RefreshRSListView
    End If
    
    mdlcommon.ChangeMousePointer 0, True
'    Set ncvVoxPlay = New AudioConvert
    LoadResStrings Me
On Error GoTo 0
End Sub

Private Sub lstResource_DblClick()
    If gSystem.intCurStep >= 0 Then
        CommandCreate_Click
    Else
        If tlbNavigation.Buttons("keyListen").Enabled Then
            tlbNavigation_ButtonClick tlbNavigation.Buttons("keyListen")
        End If
    End If
End Sub
'
'Private Sub ncvVoxPlay_ConvertBlock(ByVal Percent As Long)
'    DoEvents
'End Sub
'
'Private Sub ncvVoxPlay_Error(ByVal ErrorCode As NCTAUDIOCONVERTLib.ErrorConstants)
'    Debug.Print "Convert Error: " & Str(ErrorCode)
'End Sub

'刷新资源浏览列表
Public Sub RefreshRSListView()
On Error GoTo BackDoor

    Dim itmX As ListItem                    ' ListItem 变量
    Dim intCount As Integer                 ' 计数器变量
    Dim blnShowItem As Boolean
    Dim lngItemID As Long
    Dim lv_intItemType As Integer
    Dim lv_intLoop As Integer

    '' 删除现有内容
    lstResource.ListItems.Clear
    
    '' 根据数据添加新内容
    With gCallFlow.RS_Resource
    
        'Michael Added Aug,6,07 for Reflash Ado recordset
        .Requery
            
        '' Sun added 2007-04-16
        If gSystem.intCurStep >= 0 Then
            gCallFlow.RS_Resource.Filter = "L_ID = " & Trim(Str(gCallFlow.LanguageID))
        Else
            gCallFlow.RS_Resource.Filter = ""
        End If
        
        If .RecordCount > 0 Then
            
            .MoveFirst
    
            While Not .EOF
                            
                blnShowItem = False
                lngItemID = .Fields("R_ID")
                If IsNull(.Fields("R_Type")) Then
                    lv_intItemType = 0
                    For lv_intLoop = 0 To 4
                        If lngItemID > gRID(lv_intLoop).LBound And lngItemID < gRID(lv_intLoop).UBound Then
                            lv_intItemType = lv_intLoop
                            Exit For
                        End If
                    Next
                Else
                    lv_intItemType = Val(.Fields("R_Type"))
                    If lv_intItemType < 0 Or lv_intItemType > 4 Then
                        lv_intItemType = 0
                    End If
                End If
                
'                Select Case gSystem.intCurStep
'                Case 0, 1
'                    If (lngItemID > gRID(0).LBound And lngItemID < gRID(0).UBound) _
'                       Or (lngItemID > gRID(1).LBound And lngItemID < gRID(1).UBound) Then
'                        blnShowItem = True
'                    End If
'                Case 2
'                    If lngItemID > gRID(2).LBound And lngItemID < gRID(2).UBound Then
'                        blnShowItem = True
'                    End If
'                Case 3, 4
'                    If (lngItemID > gRID(3).LBound And lngItemID < gRID(3).UBound) _
'                       Or (lngItemID > gRID(4).LBound And lngItemID < gRID(4).UBound) Then
'                        blnShowItem = True
'                    End If
'                Case Else
'                    blnShowItem = True
'                End Select

                Select Case gSystem.intCurStep
                Case 0, 1
                    If lv_intItemType = 0 Or lv_intItemType = 1 Then
                        blnShowItem = True
                    End If
                Case 2
                    If lv_intItemType = 2 Then
                        blnShowItem = True
                    End If
                Case 3, 4
                    If lv_intItemType = 3 Or lv_intItemType = 4 Then
                        blnShowItem = True
                    End If
                Case Else
                    blnShowItem = True
                End Select
    
                If blnShowItem Then
                
                    ''' 使用 Add 方法添加新的 ListItem 并为新引用设置对象。
                    ''' 使用引用设置属性。
                    Set itmX = lstResource.ListItems.Add(, , Right("00000" & Trim(Str(lngItemID)), 5))
                    intCount = intCount + 1                 'Tag 属性计数器递增。
                    'Michael Note : Run time Error, .Key accept a string type
                    itmX.Key = "k" & CStr(.AbsolutePosition)
                    itmX.Tag = lv_intItemType
                    
                    '若 L_ID 字段不为空，则设置 subitem 1 为此字段。
                    If Not IsNull(.Fields("L_ID")) Then
                        itmX.SubItems(1) = Trim(.Fields("L_ID"))
                    End If
                    
                    '若 R_Type 字段不为空，则设置 subitem 2 为此字段。
                    itmX.SubItems(2) = gRID(lv_intItemType).Caption
                    
                    '若 Description 字段不为空，则设置 subitem 3 为此字段。
                    If Not IsNull(.Fields("R_Description")) Then
                        itmX.SubItems(3) = Trim(.Fields("R_Description"))
                    End If
                      
                    '若 PATH 字段不为空，则设置 subitem 4 为此字段。
                    If Not IsNull(.Fields("R_Path")) Then
                        itmX.SubItems(4) = Trim(.Fields("R_Path"))
                    End If
                    
                    '若 CreateTime 字段不为空，则设置 subitem 5 为此字段。
                    If Not IsNull(.Fields("CreateTime")) Then
                        itmX.SubItems(5) = Format(.Fields("CreateTime"), "yyyy-mm-dd hh:nn:ss")
                    End If
                    
                    '若 ModifyTime 字段不为空，则设置 subitem 6 为此字段。
                    If Not IsNull(.Fields("ModifyTime")) Then
                        itmX.SubItems(6) = Format(.Fields("ModifyTime"), "yyyy-mm-dd hh:nn:ss")
                    End If
                    
                    '若 R_Note 字段不为空，则设置 subitem 7 为此字段。
                    If Not IsNull(.Fields("R_Note")) Then
                        itmX.SubItems(7) = Trim(.Fields("R_Note"))
                    End If
                    
                    'Select Item
                    If Not gSystem.crlCurItem Is Nothing Then
                        If lngItemID = Val(gSystem.crlCurItem.Text) Then
                            itmX.Selected = True
                            itmX.EnsureVisible
                        Else
                            itmX.Selected = False
                        End If
                    End If
                    
                End If
                ' Next Item
                .MoveNext
            Wend
        End If
    End With
    lstResource.Refresh
Exit Sub

BackDoor:
    Debug.Print Err.Description
End Sub

Private Sub lstResource_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Tag < 2 Then
        tlbNavigation.Buttons("keyListen").Enabled = True
        tlbNavigation.Buttons("keyStop").Enabled = True
        tlbNavigation.Buttons("keyRecord").Enabled = True
        tlbNavigation.Buttons("keyTTS").Enabled = True
        tlbNavigation.Buttons("keyTools").Enabled = (gSystem.strVoiceEditorPath <> "")
    Else
        tlbNavigation.Buttons("keyListen").Enabled = False
        tlbNavigation.Buttons("keyStop").Enabled = False
        tlbNavigation.Buttons("keyRecord").Enabled = False
        tlbNavigation.Buttons("keyTTS").Enabled = False
        tlbNavigation.Buttons("keyTools").Enabled = False
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'记录漫游
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tlbNavigation_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Dim lv_strFName As String
    Dim lv_intLID As Integer
    Dim lv_Item As MSComctlLib.ListItem
    Dim lv_Index As Integer
        
    '' Play VOX
    If Button.Key = "keyListen" Then
        '' 拨放语音
        If Not lstResource.SelectedItem Is Nothing Then
            lv_strFName = lstResource.SelectedItem.Text
            
'*********************************************************************
'   Note : DO NOT DELETE,MAY BE USEFUL ...
'*********************************************************************
            '' Sun added 2007-04-16
            'Michael Commented : Jul,23,07
'            If gSystem.intCurStep >= 0 Then
'                Call F_PlayVoxFile(CLng(lv_strFName))
'            Else
'                lv_intLID = Val(lstResource.SelectedItem.SubItems(1))
'                If lstResource.SelectedItem.Tag < 2 Then
'                    Call F_PlayVoxFile(CLng(lv_strFName), lv_intLID)
'                End If
'            End If
            
            If LCase(Right(lstResource.SelectedItem.SubItems(4), 3)) = "vox" Then
                frmMain.VOXPlayer.PlayVOXFile gSystem.strPath_SysVox & lstResource.SelectedItem.SubItems(4)
            ElseIf LCase(Right(lstResource.SelectedItem.SubItems(4), 3)) = "wav" Then
                'Play wav file
                Dim strWAVPath As String
                strWAVPath = gSystem.strPath_SysVox & lstResource.SelectedItem.SubItems(4)
                sndPlaySound 0, 0
                sndPlaySound strWAVPath, &H1 Or &H2
            End If
        End If
        
    ElseIf Button.Key = "keyStop" Then
        '' Stop Play
        StopSound
    
    ElseIf Button.Key = "keyRecord" Then
        '' Record
        'Michael Added @ Aug,7,07
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If Not lstResource.SelectedItem Is Nothing Then
            If lstResource.SelectedItem.Tag < 2 Then
                frmRecordVoice.txtScript = lstResource.SelectedItem.SubItems(3)
                frmRecordVoice.txtFilePath = gSystem.strPath_SysVox & lstResource.SelectedItem.SubItems(4)
                'check exists file
                If fso.FileExists(frmRecordVoice.txtFilePath) = True Then
                    If MsgBox(LoadResString(1928) & lstResource.SelectedItem.SubItems(4) & LoadResString(1929), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) = vbYes Then
                        frmRecordVoice.Show vbModal
                    End If
                Else
                    frmRecordVoice.Show vbModal
                End If
            End If
        End If
        Set fso = Nothing
    
    ElseIf Button.Key = "keyTTS" Then
        ' TTS Convert
        ' Michael Complated this function @ Jul,10,07
        If Not lstResource.SelectedItem Is Nothing Then
            If lstResource.SelectedItem.Tag < 2 Then
                TTSConvert
            End If
        End If
        '----  To Here  -----------
        
    ElseIf Button.Key = "keyTools" Then
        ' Edit Tools
        On Error Resume Next
        If Not lstResource.SelectedItem Is Nothing Then
            If lstResource.SelectedItem.Tag < 2 Then
                lv_strFName = gSystem.strPath_SysVox & lstResource.SelectedItem.SubItems(4)
                Call ShellExecute(0, "open", gSystem.strVoiceEditorPath, lv_strFName, vbNullString, 1)
            End If
        End If
    Else
        If lstResource.ListItems.Count > 0 Then
            
            Select Case Button.Key
            Case "keyFirst"
                Set lv_Item = lstResource.ListItems(1)
                        
            Case "keyPrev"
                lv_Index = lstResource.SelectedItem.Index
                lv_Index = lv_Index - 1
                If lv_Index < 1 Then
                    lv_Index = 1
                End If
                Set lv_Item = lstResource.ListItems(lv_Index)
                
            Case "keyNext"
                lv_Index = lstResource.SelectedItem.Index
                lv_Index = lv_Index + 1
                If lv_Index > lstResource.ListItems.Count Then
                    lv_Index = lstResource.ListItems.Count
                End If
                Set lv_Item = lstResource.ListItems(lv_Index)
    
            Case "keyLast"
                Set lv_Item = lstResource.ListItems(lstResource.ListItems.Count)
                        
            End Select
            
            lv_Item.Selected = True
            lv_Item.EnsureVisible
            
        End If
    End If
End Sub

' 记录维护
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lv_iLoop   As Integer
    Dim iDelSel As Integer
    Dim iRowSel As Integer  '被选中的删除行,循环变量
    
    If Button.Key = "keyCopy" Then
        '' 是否选中复制条目
        If lstResource.SelectedItem Is Nothing Then
            Message "M148"
            lstResource.SetFocus
            Exit Sub
        Else
            iRowSel = lstResource.SelectedItem.Index
        End If
        
        '' Copy
        frmResourceItem.txtRID = ""
        frmResourceItem.txtRID.Enabled = True
        frmResourceItem.txtPath = lstResource.SelectedItem.SubItems(4)
        frmResourceItem.txtDescription = lstResource.SelectedItem.SubItems(3)
        frmResourceItem.txtCreateTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmResourceItem.txtModifyTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmResourceItem.txtNote = lstResource.SelectedItem.SubItems(7)
        frmResourceItem.cboLType.ListIndex = SearchItemDataIndex(frmResourceItem.cboLType, CLng(Val(lstResource.SelectedItem.SubItems(1))), 0)
        frmResourceItem.cboRType.ListIndex = SearchItemDataIndex(frmResourceItem.cboRType, CLng(Val(lstResource.SelectedItem.Tag)), 0)
        frmResourceItem.m_blnDataChanged = False
        Call frmResourceItem.ToolbarSetting
        frmResourceItem.Show vbModal
        
    ElseIf Button.Key = "keyInsert" Then
        '' Add New
        frmResourceItem.txtRID = ""
        frmResourceItem.txtRID.Enabled = True
        frmResourceItem.txtPath = ""
        frmResourceItem.txtDescription = ""
        frmResourceItem.txtCreateTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmResourceItem.txtModifyTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmResourceItem.txtNote = ""
        frmResourceItem.cboLType.Tag = gCallFlow.LanguageID
        frmResourceItem.cboLType.ListIndex = SearchItemDataIndex(frmResourceItem.cboLType, CLng(Val(Node0_Data1.Languages)), 0)
        frmResourceItem.cboRType.ListIndex = SearchItemDataIndex(frmResourceItem.cboRType, CLng(Val(gSystem.intCurStep)), 0)
        frmResourceItem.m_blnDataChanged = False
        Call frmResourceItem.ToolbarSetting
        frmResourceItem.Show vbModal

       
    ElseIf Button.Key = "keyUpdate" Then
        '' Update
        If lstResource.SelectedItem Is Nothing Then
            Message "M149"
            Exit Sub
        Else
            iRowSel = lstResource.SelectedItem.Index
        End If
        frmResourceItem.txtRID = lstResource.SelectedItem.Text
        frmResourceItem.txtRID.Enabled = False
        frmResourceItem.txtPath = lstResource.SelectedItem.SubItems(4)
        frmResourceItem.txtDescription = lstResource.SelectedItem.SubItems(3)
        frmResourceItem.txtCreateTime = lstResource.SelectedItem.SubItems(5)
        frmResourceItem.txtModifyTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmResourceItem.txtNote = lstResource.SelectedItem.SubItems(7)
        frmResourceItem.cboLType.ListIndex = SearchItemDataIndex(frmResourceItem.cboLType, CLng(Val(lstResource.SelectedItem.SubItems(1))), 0)
        frmResourceItem.cboRType.ListIndex = SearchItemDataIndex(frmResourceItem.cboRType, CLng(Val(lstResource.SelectedItem.Tag)), 0)
        frmResourceItem.m_blnDataChanged = False
        Call frmResourceItem.ToolbarSetting
        frmResourceItem.Show vbModal
         
    'Delete, Michael added @ Jul,16,07
    ElseIf Button.Key = "keyDelete" Then
        If lstResource.SelectedItem Is Nothing Then
            Message "M149"
            Exit Sub
        End If
        
        If vbYes = MsgBox(LoadResString(1930), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) Then
        
            'Modified @ Sep,4,07 ,delete multi-items
            Dim lv_isRowSelNull As Boolean
            For iDelSel = 1 To lstResource.ListItems.Count
                If iDelSel > lstResource.ListItems.Count Then Exit For
                If lstResource.ListItems(iDelSel).Selected = True Then
                    'Get the first line of selected
                    If Not lv_isRowSelNull Then iRowSel = iDelSel: lv_isRowSelNull = True
                    Dim lv_RID As String, lv_LID As String, lv_intLID As Byte
                    lv_RID = lstResource.ListItems(iDelSel).Text
                    lv_LID = lstResource.ListItems(iDelSel).SubItems(1)
                    If lv_LID = "普通话" Or lv_LID = "0" Then
                        lv_intLID = 0
                    ElseIf lv_LID = "广东话" Or lv_LID = "1" Then
                        lv_intLID = 1
                    ElseIf lv_LID = "英语" Or lv_LID = "2" Then
                        lv_intLID = 2
                    ElseIf lv_LID = "日语" Or lv_LID = "3" Then
                        lv_intLID = 3
                    End If
                    
                    Call UpdateResource("DELETE from tbResource WHERE P_ID='" & txtResID & "' AND R_ID='" & lv_RID & "' AND L_ID='" & lv_intLID & "'")
                    'Mike Added @ 2008-7-7
                    Call WriteLogMessage(0, enu_Information, "Delete Resource item:" & lv_RID & " from Project:" & txtResID)
                End If
            Next iDelSel
            
            'Michael Modified @ 2007-11-17
            If gbCallFromPro = 1 Then
                FillRSListView gstrSQL
            Else
                RefreshRSListView
            End If
            
            'Mike Added @ 2008-6-27; After delete an item, move the course to the next item
            If iRowSel <= lstResource.ListItems.Count Then
                lstResource.ListItems.Item(iRowSel).Selected = True
                lstResource.SelectedItem.EnsureVisible
            ElseIf lstResource.ListItems.Count > 1 Then
                lstResource.ListItems.Item(iRowSel - 1).Selected = True
                lstResource.SelectedItem.EnsureVisible
            End If
            
            
        End If
        
    ElseIf Button.Key = "keySearch" Then
        'Search ,Michael Added @ Jul,16.07
        frmResourceSearch.Show vbModal
        
    'Michael Added @ Sep,4,07 For Clear Search Result
    ElseIf Button.Key = "KeyReflash" Then
        lstResource.ListItems.Clear
        'FillRSListView gstrSQL
        'Michael Modified @ 2007-11-17
        If gbCallFromPro = 1 Then
            FillRSListView gstrSQL
        Else
            RefreshRSListView
        End If
    
        '' Sun added 2008-02-04
        'lstResource.SelectedItem.EnsureVisible
    
    End If
End Sub

'Michael Added @ Jul,18,07
Public Sub FillRSListView(strSQL As String)
On Error GoTo BackDoor

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim itmX As ListItem                    ' ListItem 变量
    Dim intCount As Integer                 ' 计数器变量
    Dim blnShowItem As Boolean
    Dim lngItemID As Long
    Dim lv_intItemType As Integer
    Dim lv_intLoop As Integer
    
    '' 根据数据添加新内容
    With gCallFlow.RS_Resource
    'With rs
'        If .State = 1 Then
'            .Close
'        End If
        
        
'        .CursorType = adOpenStatic
'        .LockType = adLockPessimistic
'        .CursorLocation = adUseClient
'        .Open strSQL, gSystem.strConString
        
'        Set rs = Nothing
'        gCallFlow.OpenResourceTable strSQL
        
        .Requery
        
        If gbSearchFlag = 1 Then
            ' 删除现有内容
            lstResource.ListItems.Clear
            '' Sun added 2007-04-16
            If gSystem.intCurStep >= 0 Then
                .Filter = "L_ID = " & Trim(Str(gCallFlow.LanguageID))
            Else
                .Filter = ""
            End If
    
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    blnShowItem = False
                    lngItemID = .Fields("R_ID")
                    If IsNull(.Fields("R_Type")) Then
                        lv_intItemType = 0
                        For lv_intLoop = 0 To 4
                            If lngItemID > gRID(lv_intLoop).LBound And lngItemID < gRID(lv_intLoop).UBound Then
                                lv_intItemType = lv_intLoop
                                Exit For
                            End If
                        Next
                    Else
                        lv_intItemType = Val(.Fields("R_Type"))
                        If lv_intItemType < 0 Or lv_intItemType > 4 Then
                            lv_intItemType = 0
                        End If
                    End If
    
                    Select Case gSystem.intCurStep
                    Case 0, 1
                        If lv_intItemType = 0 Or lv_intItemType = 1 Then
                            blnShowItem = True
                        End If
                    Case 2
                        If lv_intItemType = 2 Then
                            blnShowItem = True
                        End If
                    Case 3, 4
                        If lv_intItemType = 3 Or lv_intItemType = 4 Then
                            blnShowItem = True
                        End If
                    Case Else
                        blnShowItem = True
                    End Select
    
                    If blnShowItem Then
    
                        ''' 使用 Add 方法添加新的 ListItem 并为新引用设置对象。
                        ''' 使用引用设置属性。
                        Set itmX = lstResource.ListItems.Add(, , Right("00000" & Trim(Str(lngItemID)), 5))
                        intCount = intCount + 1                 'Tag 属性计数器递增。
                        'Michael Note : Run time Error, .Key is an invaid key
                        'itmX.Key = .AbsolutePosition
                        itmX.Key = "k" & CStr(.AbsolutePosition)
                        itmX.Tag = lv_intItemType
    
                        '若 L_ID 字段不为空，则设置 subitem 1 为此字段。
                        If Not IsNull(.Fields("L_ID")) Then
                            itmX.SubItems(1) = Trim(.Fields("L_ID"))
                        End If
    
                        '若 R_Type 字段不为空，则设置 subitem 2 为此字段。
                        itmX.SubItems(2) = gRID(lv_intItemType).Caption
    
                        '若 Description 字段不为空，则设置 subitem 3 为此字段。
                        If Not IsNull(.Fields("R_Description")) Then
                            itmX.SubItems(3) = Trim(.Fields("R_Description"))
                        End If
    
                        '若 PATH 字段不为空，则设置 subitem 4 为此字段。
                        If Not IsNull(.Fields("R_Path")) Then
                            itmX.SubItems(4) = Trim(.Fields("R_Path"))
                        End If
    
                        '若 CreateTime 字段不为空，则设置 subitem 5 为此字段。
                        If Not IsNull(.Fields("CreateTime")) Then
                            itmX.SubItems(5) = Format(.Fields("CreateTime"), "yyyy-mm-dd hh:nn:ss")
                        End If
    
                        '若 ModifyTime 字段不为空，则设置 subitem 6 为此字段。
                        If Not IsNull(.Fields("ModifyTime")) Then
                            itmX.SubItems(6) = Format(.Fields("ModifyTime"), "yyyy-mm-dd hh:nn:ss")
                        End If
    
                        '若 R_Note 字段不为空，则设置 subitem 7 为此字段。
                        If Not IsNull(.Fields("R_Note")) Then
                            itmX.SubItems(7) = Trim(.Fields("R_Note"))
                        End If
    
                        ' Select Item
                        If Not gSystem.crlCurItem Is Nothing Then
                            If lngItemID = Val(gSystem.crlCurItem.Text) Then
                                itmX.Selected = True
                                itmX.EnsureVisible
                            Else
                                itmX.Selected = False
                            End If
                        End If
    
                    End If
    
                    '' Next Item
                    .MoveNext
                Wend
            End If
        End If
        
        End With
                
    lstResource.Refresh
Exit Sub

BackDoor:
   On Error GoTo 0
End Sub

'TTS convert the text to speech
Private Sub TTSConvert()
On Error GoTo VoiceDebug
    Dim Voice As SpVoice
    Set Voice = New SpVoice
    
    'Create a new wave stream
    Dim cpFileStream As New SpFileStream
    Dim lv_iCount As Integer
    
    For lv_iCount = 1 To lstResource.ListItems.Count
        If lstResource.ListItems(lv_iCount).Selected Then
            
            'deal with same resource description
            If lv_iCount > 1 Then
                Dim lv_loop As Integer
                For lv_loop = lv_iCount To 2 Step -1
                    If lstResource.ListItems(lv_iCount).SubItems(3) = lstResource.ListItems(lv_loop - 1).SubItems(3) Then
                        If MsgBox("" & lstResource.ListItems(lv_iCount).Text & LoadResString(1931) & lstResource.ListItems(lv_loop - 1).Text & LoadResString(1933), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) <> vbYes Then
                            Exit Sub
                        End If
                    End If
                Next lv_loop
            End If
            
            'deal with same resource name
            If lv_iCount > 1 Then
                For lv_loop = lv_iCount To 2 Step -1
                    If lstResource.ListItems(lv_iCount).SubItems(4) = lstResource.ListItems(lv_loop - 1).SubItems(4) Then
                        If MsgBox("" & lstResource.ListItems(lv_iCount).Text & LoadResString(1932) & lstResource.ListItems(lv_loop - 1).Text & LoadResString(1933), vbYesNo + vbQuestion + vbDefaultButton2, App.Title & LoadResString(1934) & lstResource.ListItems(lv_loop - 1).Text & LoadResString(1935)) <> vbYes Then
                            Exit Sub
                        End If
                    End If
                Next lv_loop
            End If
                        
            strNamePath = gSystem.strPath_SysVox & lstResource.ListItems(lv_iCount).SubItems(4)
            
            Dim fsoTTS As Object
            Set fsoTTS = CreateObject("Scripting.FileSystemObject")
            'check exists file
            If fsoTTS.FileExists(strNamePath) Then
                If MsgBox(LoadResString(1928) & lstResource.ListItems(lv_iCount).SubItems(4) & LoadResString(1929), vbYesNo + vbQuestion + vbDefaultButton2, App.Title) = vbNo Then
                    Exit For
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
                Voice.Speak lstResource.ListItems(lv_iCount).SubItems(3), SVSFDefault
                'Close the file
                cpFileStream.Close
                
                Set cpFileStream = Nothing
                'Reset the Voice object's output to 'Nothing'.
                Set Voice.AudioOutputStream = Nothing
                'convert 2 vox file
                Call ctlConvert.WaveFile2VOX(tempNamePath, strNamePath)

                
            'Voice File Type - WAV
            ElseIf LCase(Right(strNamePath, 3)) = "wav" Then
                'Set audio format -> Michael Modified @ 2007-11-29
                Set Voice.Voice = Voice.GetVoices().Item(gSystem.intTTSVoice)
                Voice.Volume = gSystem.intTTSVolume
                Voice.Rate = gSystem.intTTSRate
                If gSystem.strTTSFormat = "SAFT8kHz8BitMono" Then
                    cpFileStream.Format.Type = SAFT8kHz8BitMono
                ElseIf gSystem.strTTSFormat = "SAFT8kHz16BitMono" Then
                    cpFileStream.Format.Type = SAFT8kHz16BitMono
                End If
                                
                cpFileStream.Open strNamePath, SSFMCreateForWrite, False
                Set Voice.AudioOutputStream = cpFileStream
                Voice.Speak lstResource.ListItems(lv_iCount).SubItems(3), SVSFDefault
                'Close the file
                cpFileStream.Close
                Set cpFileStream = Nothing
                'Reset the Voice object's output to 'Nothing'.
                Set Voice.AudioOutputStream = Nothing
                strNamePath = ""
            'Other file Type or invaid path and file name ...
            Else
                Message ("E140")
                'Mike added @ 2008-7-7
                Call WriteLogMessage(Err.Number, enu_Warnning, "Exception Occured during Recording..." & "Filepath:" & strNamePath, Err.Description)
                Exit Sub
            End If

        End If
    Next lv_iCount
    Set Voice = Nothing
Exit Sub

VoiceDebug:
    MsgBox "TTS Error" & Err.Description, vbOKOnly + vbExclamation, App.Title
    Call WriteLogMessage(Err.Number, enu_Error, "TTS Convert error, Filepath:" & strNamePath, Err.Description)
    'On Error GoTo 0
End Sub

'Delete the temp wav files
Private Sub ctlConvert_Wave2VOXFinished()
'    If tempNamePath <> "" Then Call RemoveFile(tempNamePath)
'    strNamePath = ""
'    tempNamePath = ""
'Michael Modified @ 2007-11-01
    TimerConvert.Enabled = True
End Sub

'Michael Added @ 2007-11-01
Private Sub TimerConvert_Timer()
    TimerConvert.Enabled = False
    If tempNamePath <> "" Then
        Call RemoveFile(tempNamePath)
        strNamePath = ""
        tempNamePath = ""
    End If
    
    If frmResourceItem.tempNamePath <> "" Then
        Call RemoveFile(frmResourceItem.tempNamePath)
        frmResourceItem.strNamePath = ""
        frmResourceItem.tempNamePath = ""
    End If
    
End Sub

'Michael added @ Sep,4,07
'For select multi-rows when key shift pressed
Private Sub lstResource_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Or KeyCode = vbKeyShift Then
        lstResource.MultiSelect = True
    Else
        lstResource.MultiSelect = False
    End If
End Sub

Private Sub lstResource_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Or KeyCode = vbKeyShift Then
        lstResource.MultiSelect = False
    End If
End Sub

' Michael Added @ 2007-12-5
Public Sub UpdateResource(ResourceSQL As String)

    Dim cnn As ADODB.Connection
    Dim rsUpdate As ADODB.Recordset
    Set cnn = New ADODB.Connection
    Set rsUpdate = New ADODB.Recordset
    
    With rsUpdate
        If .State = 1 Then
            .Close
        End If
        
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open ResourceSQL, gSystem.strConString
        
        Set rsUpdate = Nothing
    End With

End Sub
