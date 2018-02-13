VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProjectList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "项目列表"
   ClientHeight    =   6210
   ClientLeft      =   3915
   ClientTop       =   4500
   ClientWidth     =   11070
   Icon            =   "frmProjectList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1936"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&Q)"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Tag             =   "1909"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Tag             =   "1372"
      Top             =   5760
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstProject 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9551
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "1937"
         Text            =   "项目编号"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "1938"
         Text            =   "项目名称"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "1942"
         Text            =   "项目描述"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "1939"
         Text            =   "项目版本"
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "1940"
         Text            =   "项目作者"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "1941"
         Text            =   "项目用户"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "1106"
         Text            =   "创建时间"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "1120"
         Text            =   "修改时间"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "1115"
         Text            =   "项目类型"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   390
      Left            =   3120
      TabIndex        =   1
      Top             =   5760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "imgFile"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyCopy"
            Object.ToolTipText     =   "1088"
            Object.Tag             =   "1088"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyInsert"
            Object.ToolTipText     =   "1725"
            Object.Tag             =   "1725"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyUpdate"
            Object.ToolTipText     =   "1726"
            Object.Tag             =   "1726"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyDelete"
            Object.ToolTipText     =   "1727"
            Object.Tag             =   "1727"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keySearch"
            Object.ToolTipText     =   "1656"
            Object.Tag             =   "1656"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbNavigation 
      Height          =   390
      Left            =   1080
      TabIndex        =   2
      Top             =   5760
      Width           =   1605
      _ExtentX        =   2831
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
            Object.Visible         =   0   'False
            Key             =   "keyListen"
            Object.ToolTipText     =   "1523"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "keyStop"
            Object.ToolTipText     =   "1101"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "keyRecord"
            Object.ToolTipText     =   "1717"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "keyTTS"
            Object.ToolTipText     =   "1723"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "keyTools"
            Object.ToolTipText     =   "1722"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":071E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFile 
      Left            =   10920
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":087A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":0E14
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":13AE
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":1754
            Key             =   "Update"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectList.frx":1CEE
            Key             =   "Insert"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//Author : Michael
'//Date   : July,12,07
'//Title  : Project List
'//Ver    : V 1.0
'//Modify : [1] -> 复制项目时没有提示"是否复制项目包含的资源列表" Oct,24,2K7
'////////////////////////////////////////////////////////////

Option Explicit
Public bCopyRes As Boolean  '是否复制资源开关
Public strResourceType As String
Public strFrontPath As String

' Quit the Project List
Private Sub cmdCancel_Click()
    Unload Me
End Sub

' The same as double click the select list
Private Sub cmdOK_Click()
    lstProject_DblClick
End Sub

Private Sub Form_Load()
On Error Resume Next
    ''init the tool bar and G-vars
     bCopyRes = False
     
    ''Title
    
    ''Add items
    FillProListView
    mdlcommon.ChangeMousePointer 0, True
    
    LoadResStrings Me
    
On Error GoTo 0
End Sub

'Double click Project Item then jump to Resource List
Private Sub lstProject_DblClick()

'    gCallFlow.ResourceID = lstProject.SelectedItem.Text
'    gCallFlow.ResourceName = lstProject.SelectedItem.SubItems(1)
    gbCallFromPro = 1
    gbSearchFlag = 1
    
    gstrSQL = " Select R_ID,(case when L_ID=0 then '普通话' when L_ID=1 then '广东话' " & _
             " when L_ID=2 then '英语' when L_ID=3 then '日语' end ) as L_ID, (case " & _
             " when R_Type='0' then '系统资源' when R_Type='1' then '用户语音资源' when " & _
             " R_Type='2' then '用户传真资源' when R_Type='3' then '用户COM/DLL资源' end )" & _
             " as R_Type,R_Path,R_Description,CreateTime,ModifyTime,R_Note,P_ID From " & _
             " tbResource Where P_ID= " & lstProject.SelectedItem.Text & " Order By P_ID,R_ID,L_ID "
             
    'gCallFlow.OpenResourceTable CByte(lstProject.SelectedItem.Text), gstrSQL
    gCallFlow.OpenResourceRecordSet CByte(lstProject.SelectedItem.Text), 0
             
    frmResourceList.Show vbModal
End Sub

Private Sub tlbNavigation_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lv_Item As MSComctlLib.ListItem
    Dim lv_Index As Integer
    
    If lstProject.ListItems.Count > 0 Then
        Select Case Button.Key
        
        Case "keyFirst"
            Set lv_Item = lstProject.ListItems(1)
            
        Case "keyPrev"
            lv_Index = lstProject.SelectedItem.Index
            lv_Index = lv_Index - 1
            If lv_Index < 1 Then
                lv_Index = 1
            End If
            Set lv_Item = lstProject.ListItems(lv_Index)
            
        Case "keyNext"
            lv_Index = lstProject.SelectedItem.Index
            lv_Index = lv_Index + 1
            If lv_Index > lstProject.ListItems.Count Then
                lv_Index = lstProject.ListItems.Count
            End If
            Set lv_Item = lstProject.ListItems(lv_Index)
            
        Case "keyLast"
            Set lv_Item = lstProject.ListItems(lstProject.ListItems.Count)
        
        End Select
        lv_Item.Selected = True
        lv_Item.EnsureVisible
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Key = "keyCopy" Then
        If lstProject.SelectedItem Is Nothing Then
            Message "will add to resurce ..."
            lstProject.SetFocus
            Exit Sub   'Michael Added @ Oct,24,2K7
        End If
        
        'Copy operation
        frmProjectItem.txtPID = ""
        frmProjectItem.txtPID.Enabled = True
        frmProjectItem.txtPVersion = lstProject.SelectedItem.SubItems(3)
        frmProjectItem.txtPName = lstProject.SelectedItem.SubItems(1)
        frmProjectItem.txtPAuthor = lstProject.SelectedItem.SubItems(4)
        frmProjectItem.txtPUser = lstProject.SelectedItem.SubItems(5)
        frmProjectItem.txtPCTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmProjectItem.txtPMTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmProjectItem.txtPDes = lstProject.SelectedItem.SubItems(2)
        
        'Michael Added @ Oct,24,2K7
        'Michael Modified @ 2007-11-27
        '扩展复制资源时的用户选项,包括资源后缀名和资源所属项目
        '判断是否是空项目,如果是空项目就直接复制
        Call GetProjectProperty(lstProject.SelectedItem.Text)
        
'        If vbYes = MsgBox("是否复制项目包含的资源?", vbYesNo, "IVR Editor") Then
'            bCopyRes = True
'        End If
'        frmProjectItem.Show vbModal
        
    ElseIf Button.Key = "keyInsert" Then
        frmProjectItem.txtPID = ""
        frmProjectItem.txtPID.Enabled = True
        frmProjectItem.txtPVersion = ""
        frmProjectItem.txtPName = ""
        frmProjectItem.txtPAuthor = ""
        frmProjectItem.txtPUser = ""
        frmProjectItem.txtPCTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmProjectItem.txtPMTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmProjectItem.txtPDes = ""
        frmProjectItem.Show vbModal
    
    ElseIf Button.Key = "keyUpdate" Then
        frmProjectItem.txtPID = lstProject.SelectedItem.Text
        frmProjectItem.txtPID.Enabled = False
        frmProjectItem.txtPVersion = lstProject.SelectedItem.SubItems(3)
        frmProjectItem.txtPName = lstProject.SelectedItem.SubItems(1)
        frmProjectItem.txtPAuthor = lstProject.SelectedItem.SubItems(4)
        frmProjectItem.txtPUser = lstProject.SelectedItem.SubItems(5)
        frmProjectItem.txtPCTime = lstProject.SelectedItem.SubItems(7)
        frmProjectItem.txtPMTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
        frmProjectItem.txtPDes = lstProject.SelectedItem.SubItems(2)
        frmProjectItem.Show vbModal
    
    ElseIf Button.Key = "keyDelete" Then
        If lstProject.SelectedItem Is Nothing Then
            Message "will add to resource later ..."
            lstProject.SetFocus
            Exit Sub
        End If
        
        If Message("Q020", vbDefaultButton2) = vbYes Then
            gCallFlow.RS_Project.Open "Delete from tbRefIVR Where P_ID = " & lstProject.SelectedItem.Text & " AND P_TYPE = 'R'"
            'Mike added @ 2008-7-7
            Call WriteLogMessage(0, enu_Information, "Delete Resource Project:" & lstProject.SelectedItem.Text)
            'Michael Added @ 2007-12-5 --> 删除项目下的资源项
            gCallFlow.RS_Project.Open "DELETE FROM tbResource WHERE P_ID = " & lstProject.SelectedItem.Text & " "
            FillProListView
        End If
        
    ElseIf Button.Key = "keySearch" Then
        frmProjectSearch.Show vbModal
    
    End If
    
End Sub

'Fill the Project List
Public Sub FillProListView()
On Error GoTo BackDoor

Dim itmX As ListItem
Dim intCount As Integer
    
'Clear the old informations
lstProject.ListItems.Clear

' Add the contect
With gCallFlow.RS_Project
    
    .CursorType = adOpenKeyset
    .LockType = adLockReadOnly
    .CursorLocation = adUseClient
    .Open "Select * from tbrefivr where P_Type = 'R' Order By P_ID", gSystem.strConString

    If .RecordCount > 0 Then
        .MoveFirst
        While Not .EOF
            Set itmX = lstProject.ListItems.Add(, , .Fields("P_ID"))
            
            '项目名称
            If Not IsNull(.Fields("P_Name")) Then
                itmX.SubItems(1) = Trim(.Fields("P_Name"))
            End If
            
            '项目描述
            If Not IsNull(.Fields("P_Description")) Then
                itmX.SubItems(2) = Trim(.Fields("P_Description"))
            End If
            
            '项目版本
            If Not IsNull(.Fields("P_Version")) Then
                itmX.SubItems(3) = Trim(.Fields("P_Version"))
            End If
            
            '项目作者
            If Not IsNull(.Fields("P_Auther")) Then
                itmX.SubItems(4) = Trim(.Fields("P_Auther"))
            End If
            
            '项目用户
            If Not IsNull(.Fields("P_User")) Then
                itmX.SubItems(5) = Trim(.Fields("P_User"))
            End If
            
            '创建时间
            If Not IsNull(.Fields("P_CreateTime")) Then
                itmX.SubItems(6) = Format(.Fields("P_CreateTime"), "yyyy-mm-dd hh:nn:ss")
            End If
            
            '修改时间
            If Not IsNull(.Fields("P_ModifyTime")) Then
                itmX.SubItems(7) = Format(.Fields("P_ModifyTime"), "yyyy-mm-dd hh:nn:ss")
            End If
            
            '项目类型
            If Not IsNull(.Fields("P_Type")) Then
                itmX.SubItems(8) = Trim(.Fields("P_Type"))
            End If
            
            .MoveNext
        Wend
    End If
    .Close
End With
lstProject.Refresh
Exit Sub

BackDoor:
    'Used for Debug,Delete when release
    'MsgBox (Err.Description)
    On Error GoTo 0
End Sub

'判断项目是否为空项目
Private Sub GetProjectProperty(strProjectID As String)
    Dim RS_Res As ADODB.Recordset
    Set RS_Res = New ADODB.Recordset
    Dim strCopyRes As String
    
    Dim strSQLRes As String
    strSQLRes = "SELECT P_ID, R_ID, R_Path FROM tbResource WHERE " & _
                "P_ID= '" & strProjectID & "' Order By R_ID"
    
    With RS_Res
        If .State = 1 Then
            .Close
        End If
        
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open strSQLRes, gSystem.strConString
        Set RS_Res = Nothing
        
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(Trim(.Fields("R_Path"))) Then
                strResourceType = LCase(Right(Trim(.Fields("R_Path")), 3))
                Dim lv_count As Integer, lv_car As String
                For lv_count = 1 To Len(Trim(.Fields("R_Path")))
                    lv_car = Mid(Trim(.Fields("R_Path")), lv_count, 1)
                    If lv_car = "\" Then
                        strFrontPath = Left(Trim(.Fields("R_Path")), lv_count)
                        Exit For
                    Else
                        strFrontPath = ""
                    End If
                Next lv_count
                
                frmProjectCopyDlg.Show vbModal
            End If
        Else
            frmProjectItem.Show vbModal
        End If
        
    End With

End Sub

