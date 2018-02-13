VERSION 5.00
Begin VB.Form frmResourceSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "资源查询"
   ClientHeight    =   870
   ClientLeft      =   6105
   ClientTop       =   5550
   ClientWidth     =   6315
   Icon            =   "frmResourceSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSchName 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox cboSchOpt 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtSchState 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblSearchName 
      AutoSize        =   -1  'True
      Caption         =   "字段名称"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblSearchOpt 
      AutoSize        =   -1  'True
      Caption         =   "运算符"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblSearchSta 
      AutoSize        =   -1  'True
      Caption         =   "表达式"
      Height          =   195
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmResourceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/File Version : V1.0
'/Author       : Michael
'/ Date        : Jul,16,07
'/////////////////////////////////////////////////////////////

Option Explicit
Dim strSQL As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cboSchName.Text = "" Then
        Message "E134"
        cboSchName.SetFocus
        Exit Sub
    'Added @ Sep,4,07
    Else
        Call WriteIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecName, Str(cboSchName.ListIndex), gSystem.strINI_File)
    End If
    
    If cboSchOpt.Text = "" Then
        Message "E135"
        cboSchOpt.SetFocus
        Exit Sub
    'Added @ Sep,4,07
    Else
        Call WriteIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecOP, Str(cboSchOpt.ListIndex), gSystem.strINI_File)
    End If
    
    If txtSchState.Text = "" Then
        Message "E136"
        txtSchState.SetFocus
        Exit Sub
    'Added @ Sep,4,07
    Else
        Call WriteIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecExp, txtSchState.Text, gSystem.strINI_File)
    End If
    
    If cboSchName.ListIndex = 0 Then
        strSQL = "Select * from tbResource where R_ID  " & cboSchOpt.Text & " " & txtSchState.Text & " and P_ID='" & frmResourceList.txtResID & "'  Order By R_ID"
    ElseIf cboSchName.ListIndex = 1 Then
        strSQL = "Select * from tbResource where L_ID  " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order By R_ID"
    ElseIf cboSchName.ListIndex = 2 Then
        strSQL = "Select * from tbResource where R_Type " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order By R_ID"
    ElseIf cboSchName.ListIndex = 3 Then
        strSQL = "Select * from tbResource where R_Path  " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order By R_ID"
    ElseIf cboSchName.ListIndex = 4 Then
        strSQL = "Select * from tbResource where R_Description  " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order By R_ID"
    ElseIf cboSchName.ListIndex = 5 Then
        strSQL = "Select * from tbResource where CreateTime  " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order By R_ID"
    ElseIf cboSchName.ListIndex = 6 Then
        strSQL = "Select * from tbResource where ModifyTime  " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order By R_ID"
    ElseIf cboSchName.ListIndex = 7 Then
        strSQL = "Select * from tbResource where R_Note " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order By R_ID"
    ElseIf cboSchName.ListIndex = 8 Then
        strSQL = "select * from tbResource where P_ID " & cboSchOpt.Text & " '" & txtSchState.Text & "' and P_ID='" & frmResourceList.txtResID & "' Order by R_ID   "
    End If
    ReFillProList
    Unload Me
End Sub

Private Sub Form_Load()
    ' Fill and init the combo box - 字段名称,运算符
    cboSchName.Clear
    cboSchName.AddItem "资源编号", 0
    cboSchName.AddItem "资源语言", 1
    cboSchName.AddItem "资源类型", 2
    cboSchName.AddItem "资源路径", 3
    cboSchName.AddItem "资源描述", 4
    cboSchName.AddItem "创建时间", 5
    cboSchName.AddItem "修改时间", 6
    cboSchName.AddItem "资源备注", 7
    cboSchName.AddItem "项目编号", 8
    
    cboSchOpt.Clear
    cboSchOpt.AddItem ">", 0
    cboSchOpt.AddItem ">=", 1
    cboSchOpt.AddItem "=", 2
    cboSchOpt.AddItem "<", 3
    cboSchOpt.AddItem "<=", 4
    cboSchOpt.AddItem "like", 5
    
    txtSchState.Text = ""
    
    'Michael Added @ Sep,4,07 For Save the Search condi to ini file
    If GetIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecName, gSystem.strINI_File) <> "" Then
        cboSchName.ListIndex = CInt(GetIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecName, gSystem.strINI_File))
    End If
    
    If GetIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecOP, gSystem.strINI_File) <> "" Then
        cboSchOpt.ListIndex = CInt(GetIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecOP, gSystem.strINI_File))
    End If
    
    If GetIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecExp, gSystem.strINI_File) <> "" Then
        txtSchState.Text = CStr(GetIniFileString(Def_INI_SEC_SearchCon, Def_INI_ENTRY_SecExp, gSystem.strINI_File))
    End If
    
End Sub

Public Sub ReFillProList()
On Error GoTo BackDoor

    Dim itmX As ListItem
    Dim intCount As Integer
    
    'Clear the old informations
    frmResourceList.lstResource.ListItems.Clear

' Add the contect
With gCallFlow.RS_Project

    .CursorType = adOpenKeyset
    .LockType = adLockReadOnly
    .CursorLocation = adUseClient
    .Open strSQL, gSystem.strConString

    If .RecordCount > 0 Then
        .MoveFirst
        While Not .EOF
            

            ''' 使用 Add 方法添加新的 ListItem 并为新引用设置对象。
            ''' 使用引用设置属性。
            Set itmX = frmResourceList.lstResource.ListItems.Add(, , .Fields("R_ID"))
            intCount = intCount + 1                 'Tag 属性计数器递增。
            'Michael Note : Run time Error, .Key is an invaid key
            'itmX.Key = .AbsolutePosition
            'itmX.Tag = frmResourceList.lv_intItemType

            '若 L_ID 字段不为空，则设置 subitem 1 为此字段。
            If Not IsNull(.Fields("L_ID")) Then
                itmX.SubItems(1) = Trim(.Fields("L_ID"))
            End If

            '若 R_Type 字段不为空，则设置 subitem 2 为此字段。
            'itmX.SubItems(2) = gRID(lv_intItemType).Caption

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

        .MoveNext

        Wend
    End If
    .Close
End With

frmResourceList.lstResource.Refresh

BackDoor:
    On Error GoTo 0
    'MsgBox Err.Description
End Sub
