VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSetPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择文件"
   ClientHeight    =   4425
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6120
   Icon            =   "frmSetPath.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1747"
   Begin MSComctlLib.TreeView tvwPath 
      Height          =   3975
      Left            =   2370
      TabIndex        =   9
      Top             =   360
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   7011
      _Version        =   393217
      Indentation     =   2
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5040
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetPath.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetPath.frx":23EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetPath.frx":40F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetPath.frx":68AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetPath.frx":8C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetPath.frx":A996
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListPath 
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   4710
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   2143
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cboFileType 
      Height          =   315
      Left            =   5190
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   885
   End
   Begin VB.FileListBox FileList 
      Height          =   3990
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   2265
   End
   Begin VB.DirListBox DirList 
      Height          =   1215
      Left            =   2850
      TabIndex        =   2
      Top             =   4740
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   5190
      TabIndex        =   1
      Tag             =   "1144"
      Top             =   570
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   5190
      TabIndex        =   0
      Tag             =   "1372"
      Top             =   150
      Width           =   885
   End
   Begin VB.Image imgHome 
      Height          =   480
      Left            =   3300
      Picture         =   "frmSetPath.frx":D148
      ToolTipText     =   "语音文件根路径"
      Top             =   -60
      Width           =   480
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "文件类型："
      Height          =   195
      Left            =   5190
      TabIndex        =   7
      Tag             =   "1751"
      Top             =   1020
      Width           =   885
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "目录列表："
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Tag             =   "1750"
      Top             =   90
      Width           =   900
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "文件列表："
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Tag             =   "1749"
      Top             =   90
      Width           =   900
   End
End
Attribute VB_Name = "frmSetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public curFullPath As String    '当前选中的全路径名称
Dim xpos As Long, ypos As Long

Private Sub cboFileType_Change()
    FileList.Pattern = cboFileType.Text '设置文件列表框的文件显示类型
End Sub

Private Sub cboFileType_Click()
    FileList.Pattern = cboFileType.Text '设置文件列表框的文件显示类型
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'On Error GoTo BadFilename
'    comdResult = -1
'    ' 文本被解析为文件名、路径和驱动器
'    FileList.FileName = Text1.Text
'    DirList.Path = FileList.Path  ' 设置目录路径
''    Drive1.Drive = DirList.Path    ' 设置驱动器
'    FileList.Pattern = cboFileType.Text '恢复文件显示类型
'    comdResult = 0
'    Exit Sub
'BadFilename: '如果没该文件，则显示错误信息
'    MsgBox "No such file !"
'    comdResult = 0
    If Right$(gSystem.strPath_SysVox, 1) <> "\" Then
        frmResourceItem.txtPath = Mid(curFullPath, Len(gSystem.strPath_SysVox & "\") + 1, Len(curFullPath) - Len(gSystem.strPath_SysVox & "\"))
    Else
        frmResourceItem.txtPath = Mid(curFullPath, Len(gSystem.strPath_SysVox) + 1, Len(curFullPath) - Len(gSystem.strPath_SysVox))
    End If
    gStrLastResPath = dirList.path
    Unload Me
End Sub

Private Sub dirList_Change()
    FileList.path = dirList.path
    curFullPath = ""
    cmdOK.Enabled = False
End Sub

Private Sub FileList_Click()
    
    If FileList.ListIndex >= 0 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    
    If Right$(FileList.path, 1) <> "\" Then
        curFullPath = FileList.path & "\" & FileList.FileName
    Else
        curFullPath = FileList.path & FileList.FileName
    End If
    
End Sub

Private Sub FileList_DblClick()
    cmdOK_Click
End Sub

Private Sub Form_Load()
    
    '装载文件类型
    cboFileType.AddItem "*.vox"
    cboFileType.AddItem "*.wav"
    cboFileType.AddItem "*.mp3"
    cboFileType.AddItem "*.*"
    'cboFileType.ListIndex = 0
    
    'Michael Modified : default file type
    cboFileType.ListIndex = gSystem.intVoiFileType
    
    '设置在运行时显示在 FileListBox 中的文件类型
    FileList.Pattern = cboFileType.Text
    
    dirList.path = gStrLastResPath
    cmdOK.Enabled = False
    
'    ChDrive DirList
'    ChDir DirList.Path
'    Call ChDir(gSystem.strPath_SysVox)
'    DirList.Path = CurDir
    
'    curDirec = DirList.Path
'    curName = DirList.Path
     
'    Dim NameOfFile As String
'    Dim clmX As ColumnHeader
'    Dim itmX As ListItem
'    Dim Counter As Long
'    Dim dname As String
'    Dim TempDname As String
'    Dim Counter2 As Integer
'    Dim CurrentDir As String

'    ListPath.ColumnHeaders.Add , , "FilePath", ListPath.Width - 100
'
'    ListPath.BorderStyle = ccFixedSingle ' Set BorderStyle property.
'
'    ' SmallIcons properties.
'    ListPath.Icons = ImageList1
'    ListPath.SmallIcons = ImageList2
'
'    'NameOfFile = Dir$(CurrentDir & "*.*", vbDirectory)
'    Dim Fname As String
'
''    'If we are in a subdirectory then do the following
''    If Right(DirList.Path, 1) <> "\" Then
''        CurrentDir = DirList.Path & "\"
''        dname = ".."
''        Set itmX = ListPath.ListItems.Add(, , dname)
''        itmX.Icon = 3           ' Set an icon from ImageList1.
''        itmX.SmallIcon = 3      ' Set an icon from ImageList2.
''    Else
''        'If not in a subdirectory then do the following
''        CurrentDir = DirList.Path
''    End If
'
'    'Get the Directory Names
'    For Counter = 0 To DirList.ListCount - 1
'        dname = DirList.List(Counter)
'        For Counter2 = Len(dname) To 1 Step -1
'            If Mid$(dname, Counter2, 1) = "\" Then
'                TempDname = Right(dname, Len(dname) - Counter2)
'                Exit For
'            End If
'        Next Counter2
'        Set itmX = ListPath.ListItems.Add(, , TempDname)
'        itmX.Icon = 1
'        itmX.SmallIcon = 1
'    Next Counter
'
'
'    ListPath.View = lvwReport
'    ListPath.Arrange = 0 'lvwNoArrange
'    ListPath.LabelWrap = False
'    ListPath.Sorted = True

    tvwPath.LineStyle = tvwRootLines
    LoadResStrings Me
    
    Call FillupDirTree
    
End Sub

Private Sub FillupDirTree()
    Dim strPointName As String
    Dim strPointKey As String
    Dim Counter As Long
    Dim Counter2 As Long

    tvwPath.nodes.Clear
    tvwPath.nodes.Add , tvwPrevious, dirList.path, dirList.path, 1
    
    'Get the Directory Names
    For Counter = 0 To dirList.ListCount - 1
    
        strPointKey = dirList.List(Counter)
        strPointName = dirList.List(Counter)
        
        For Counter2 = Len(strPointKey) To 1 Step -1
            If Mid$(strPointKey, Counter2, 1) = "\" Then
                strPointName = Right(strPointKey, Len(strPointKey) - Counter2)
                Exit For
            End If
        Next Counter2
        
        tvwPath.nodes.Add dirList.path, tvwChild, strPointKey, strPointName, 2, 3
            
    Next Counter
    
    tvwPath.nodes(1).Expanded = True
End Sub

Private Sub imgHome_Click()
    dirList.path = gSystem.strPath_SysVox
    FillupDirTree
End Sub

Private Sub ListPath_DblClick()
'    Dim Counter As Long
'    Dim itmX As ListItem
'    Dim NameOfFile As String
'    Dim dname As String
'    Dim TempDname As String
'    Dim Counter2 As Integer
'    Dim Item As ListItem
'    Dim CurrentDir As String
'
'    If ListPath.HitTest(xpos, ypos) Is Nothing Then
'        Exit Sub
'    Else
'        Set Item = ListPath.HitTest(xpos, ypos)
'    End If
'
'
'    'Set Item = ListPath.SelectedItem
'
'    'If you Click on a filename just exit this subroutine
'    If Right(DirList.Path, 1) <> "\" Then
'        CurrentDir = DirList.Path & "\"
'    Else
'        CurrentDir = DirList.Path
'    End If
'
'    If (GetAttr(CurrentDir & Item) And vbDirectory) <= 0 Then Exit Sub
'    ListPath.ListItems.Clear 'Clear Out Old Items
'
'    'Change to selected Directory - Let Visual Basic do the work
'    ChDir Item
'
'    'Change the Directory List Box to equal the new Current Directory
'    DirList.Path = CurDir
'
'    Dim Fname As String
'
'
'    If DirList.Path <> gSystem.strPath_SysVox Then
'        'If we are in a subdirectory then add the backup ".." directory name
'        If Right(DirList.Path, 1) <> "\" Then
'            CurrentDir = DirList.Path & "\"
'            dname = ".."
'            Set itmX = ListPath.ListItems.Add(, , dname)
'            itmX.Icon = 3           ' Set an icon from ImageList1.
'            itmX.SmallIcon = 3      ' Set an icon from ImageList2.
'        Else
'            'If we are not in a subdirectory then just set our temporary Directory variable
'            CurrentDir = DirList.Path
'        End If
'    End If
'
'    'Add directory names to ListView
'    For Counter = 0 To DirList.ListCount - 1
'        dname = DirList.List(Counter)
'        'Get the actual directory name, not the full path and directory
'        For Counter2 = Len(dname) To 1 Step -1
'            If Mid$(dname, Counter2, 1) = "\" Then
'                TempDname = Right(dname, Len(dname) - Counter2)
'                Exit For
'            End If
'        Next Counter2
'        Set itmX = ListPath.ListItems.Add(, , TempDname)
'        itmX.Icon = 1           ' Set an icon from ImageList1.
'        itmX.SmallIcon = 1      ' Set an icon from ImageList2.
'
'    Next Counter
    
End Sub

Private Sub ListPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xpos = X
    ypos = Y
End Sub

Private Sub tvwPath_NodeClick(ByVal Node As MSComctlLib.Node)
    
    dirList.path = Node.Key
    
    Dim Counter As Long
    Dim Counter2 As Integer
    
    Dim strPointName As String
    Dim strPointKey As String
    
    'Get the Directory Names
    For Counter = 0 To dirList.ListCount - 1
    
        strPointKey = dirList.List(Counter)
        strPointName = dirList.List(Counter)
        
        For Counter2 = Len(strPointKey) To 1 Step -1
            If Mid$(strPointKey, Counter2, 1) = "\" Then
                strPointName = Right(strPointKey, Len(strPointKey) - Counter2)
                Exit For
            End If
        Next Counter2
        
        If CheckTreeKey(strPointKey) = False Then
            tvwPath.nodes.Add dirList.path, tvwChild, strPointKey, strPointName, 2, 3
        End If
        
    Next Counter
    
    Node.Expanded = True
End Sub

Public Function CheckTreeKey(strKey As String) As Boolean
    Dim nodes As Node
    For Each nodes In tvwPath.nodes
        If nodes.Key = strKey Then
            CheckTreeKey = True
            Exit Function
        End If
    Next
    CheckTreeKey = False
End Function
