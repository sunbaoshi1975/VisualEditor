VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frm_010 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "工作日设定"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frm_010.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1197"
   Begin VB.CommandButton cmdNodeTag 
      Height          =   333
      Left            =   6552
      Picture         =   "frm_010.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "1948"
      Top             =   7491
      Width           =   333
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "全部清除"
      Height          =   375
      Index           =   4
      Left            =   5400
      TabIndex        =   31
      Tag             =   "1210"
      Top             =   4140
      Width           =   1455
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "全部恢复"
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   30
      Tag             =   "1209"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "清除节假日"
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   29
      Tag             =   "1208"
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "恢复缺省值"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   28
      Tag             =   "1207"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   3630
      TabIndex        =   34
      Tag             =   "1144"
      Top             =   7500
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   2100
      TabIndex        =   33
      Tag             =   "1007"
      Top             =   7500
      Width           =   1065
   End
   Begin VB.Frame Frame4 
      Caption         =   "描述"
      Height          =   855
      Left            =   60
      TabIndex        =   27
      Tag             =   "1104"
      Top             =   6540
      Width           =   6825
      Begin VB.TextBox Txt_Description 
         Height          =   495
         Left            =   120
         MaxLength       =   32
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   240
         Width           =   6585
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "转移节点"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   60
      TabIndex        =   18
      Tag             =   "1203"
      Top             =   1860
      Width           =   6825
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   3
         Left            =   6390
         Picture         =   "frm_010.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "1145"
         Top             =   330
         Width           =   315
      End
      Begin VB.TextBox T_nd_trans 
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
         Index           =   2
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   24
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox T_nd_trans 
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
         Index           =   1
         Left            =   3300
         MaxLength       =   6
         TabIndex        =   22
         Top             =   300
         Width           =   765
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   2
         Left            =   4080
         Picture         =   "frm_010.frx":1906
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "1145"
         Top             =   330
         Width           =   315
      End
      Begin VB.TextBox T_nd_trans 
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
         Index           =   0
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   20
         Top             =   300
         Width           =   765
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   1
         Left            =   1830
         Picture         =   "frm_010.frx":1C90
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "1145"
         Top             =   330
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "其余时间"
         Height          =   195
         Index           =   3
         Left            =   4650
         TabIndex        =   35
         Tag             =   "1206"
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作日"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   26
         Tag             =   "1204"
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "节假日"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   2460
         TabIndex        =   19
         Tag             =   "1205"
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "日期范围"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   60
      TabIndex        =   13
      Tag             =   "1198"
      Top             =   900
      Width           =   6825
      Begin VB.ComboBox cb_Month 
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
         Index           =   1
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   300
         Width           =   795
      End
      Begin VB.ComboBox cb_Month 
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
         Index           =   0
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   795
      End
      Begin VB.ComboBox cb_Year 
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
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "个月"
         Height          =   195
         Index           =   5
         Left            =   4740
         TabIndex        =   17
         Tag             =   "1202"
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "月 开始 共"
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   16
         Tag             =   "1201"
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Tag             =   "1200"
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "从"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Tag             =   "1199"
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Tag             =   "1136"
      Top             =   60
      Width           =   6825
      Begin VB.TextBox T_nd_parent 
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
         Left            =   5580
         MaxLength       =   6
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox Cb_log 
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
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "0-不记录；1-16记录位"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox T_n_id 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox T_n_no 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   525
      End
      Begin VB.CommandButton cmdShowNodeList 
         Height          =   285
         Index           =   0
         Left            =   6390
         Picture         =   "frm_010.frx":201A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "1145"
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "父节点"
         Height          =   180
         Left            =   4980
         TabIndex        =   12
         Tag             =   "1150"
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "被访问日志"
         Height          =   195
         Left            =   2790
         TabIndex        =   11
         Tag             =   "1159"
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Lbl_n_no 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Tag             =   "1143"
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Lbl_n_id 
         AutoSize        =   -1  'True
         Caption         =   "节点ID"
         Height          =   195
         Left            =   1260
         TabIndex        =   5
         Tag             =   "1137"
         Top             =   300
         Width           =   525
      End
   End
   Begin FPSpreadADO.fpSpread sprCalen 
      Height          =   3555
      Left            =   120
      TabIndex        =   37
      Top             =   2880
      Width           =   5175
      _Version        =   458752
      _ExtentX        =   9128
      _ExtentY        =   6271
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AutoClipboard   =   0   'False
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   8
      ScrollBars      =   0
      SpreadDesigner  =   "frm_010.frx":23A4
   End
   Begin VB.Label lblNotes 
      Caption         =   "说明：红色日期表示节假日，双击切换节假日和工作日。"
      Height          =   1095
      Left            =   5400
      TabIndex        =   36
      Tag             =   "1211"
      Top             =   4740
      Width           =   1425
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuSwitch 
         Caption         =   "平日"
         Index           =   0
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "休息日"
         Index           =   1
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "节假日"
         Index           =   2
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "特别日1"
         Index           =   3
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "特别日2"
         Index           =   4
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "特别日3"
         Index           =   5
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "清除"
      End
   End
End
Attribute VB_Name = "frm_010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Node information
'//文件名：  Frm_010.frm
'//用途：    选择设定工作日
'//作者:     Tony Sun
'//创建日期：2004/04/22
'//修改日期：
'//文件描述：工作日设定
'//////////////////////////////////////////////////////////////////

Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean
Dim f_CurYear As Integer
Dim f_CurMonth As Byte
Dim f_Days As Integer
Dim f_DateType(256) As Byte

Private Sub Cb_log_Click()
    f_DataChanged = True
End Sub

Private Sub cb_Month_Click(Index As Integer)
    f_DataChanged = True
    
End Sub

Private Sub cb_Year_Click()
    f_DataChanged = True
    
End Sub

'Mike added this event @ 2008-1-30
Private Sub cmdNodeTag_Click()
    frmNodeTagEdit.iNodeID = CInt(T_n_id)
    frmNodeTagEdit.byNodeNo = CByte(T_n_no.Text)
    frmNodeTagEdit.Show vbModal
End Sub

Private Sub cmdShowNodeList_Click(Index As Integer)
    Select Case Index
    Case 0
        Set gSystem.crlCurItem = T_nd_parent
    Case Else
        Set gSystem.crlCurItem = T_nd_trans(Index - 1)
    End Select
    frmNodeList.Show vbModal

End Sub

Private Sub cmdSwitch_Click(Index As Integer)
    Dim lv_Col As Integer
    Dim lv_row As Integer
    Dim lv_lvp As Integer
    Dim lv_DateLoop As Date
    f_DataChanged = True
    Select Case Index
    Case 0                  '' 取消修改
        FillMonthDate
    Case 1                  '' 恢复缺省(Sunday & Saturday)
        For lv_row = 3 To 8
            sprCalen.Row = lv_row
            For lv_Col = 1 To 7
                sprCalen.Col = lv_Col
                If sprCalen.Text <> "" Then
                    If lv_Col Mod 6 = 1 Then
                        If sprCalen.ForeColor <> RGB(255, 0, 0) Then
                         sprCalen_DblClick lv_Col, lv_row
                         End If
                        'sprCalen.ForeColor = RGB(255, 0, 0)             ' Red
                    Else
                        'sprCalen.ForeColor = RGB(0, 0, 0)               ' Black
                        If sprCalen.ForeColor = RGB(255, 0, 0) Then
                        sprCalen_DblClick lv_Col, lv_row
                        End If
                    End If
                End If
            Next
        Next
    Case 2                  '' 清除节假日
        For lv_row = 3 To 8
            sprCalen.Row = lv_row
            For lv_Col = 1 To 7
                sprCalen.Col = lv_Col
                If sprCalen.Text <> "" Then
                    'sprCalen.ForeColor = RGB(0, 0, 0)               ' Black
                    If sprCalen.ForeColor = RGB(255, 0, 0) Then
                        sprCalen_DblClick lv_Col, lv_row
                        End If
                End If
            Next
        Next
    Case 3                  '' 全部恢复
        If Message("Q018") = vbYes Then
            f_CurYear = Val(cb_Year.List(cb_Year.ListIndex))
            f_CurMonth = Val(cb_Month(0).List(cb_Month(0).ListIndex))
            
            lv_DateLoop = CDate(Str(f_CurYear) & "-" & Str(f_CurMonth) & "-" & "01")
            For lv_lvp = 0 To 365
                If Weekday(lv_DateLoop + lv_lvp) Mod 6 = 1 Then
                    Call Set_Bit_Value(f_DateType(Int(lv_lvp / 8)), lv_lvp Mod 8, 1)
                Else
                    Call Set_Bit_Value(f_DateType(Int(lv_lvp / 8)), lv_lvp Mod 8, 0)
                End If
            Next
            
            f_Days = 0
            FillMonthDate
        End If
        
    Case 4                  '' 全部清除
        If Message("Q019") = vbYes Then
            f_CurYear = Val(cb_Year.List(cb_Year.ListIndex))
            f_CurMonth = Val(cb_Month(0).List(cb_Month(0).ListIndex))
            
            For lv_lvp = LBound(f_DateType) To UBound(f_DateType)
                f_DateType(lv_lvp) = 0
            Next
            
            f_Days = 0
            FillMonthDate
        End If
    End Select
    
End Sub

Private Sub Command1_Click()
'On Error Resume Next

Dim i As Integer

    If f_DataChanged Then
    
        Dim lv_nNewNode As Integer
        
        '' Sun added 2002-06-11
        gClipBoard.PushClipBoardStack
        
        ' 是否是主日历，不可见，系统识别
        Node10_Data1.maincalendar = 0
        
        ' 开始年份，YY
        Node10_Data1.startyear = CByte(cb_Year.ItemData(cb_Year.ListIndex))
        
        ' 开始月，1-12
        Node10_Data1.startmonth = CByte(cb_Month(0).ItemData(cb_Month(0).ListIndex))
        
        ' 共几个月（最大12个月）
        Node10_Data1.monthcount = CByte(cb_Month(1).ItemData(cb_Month(1).ListIndex))
        
        ' 保留
        Node10_Data1.reserved1 = 0
        
        ' 被访问日志
        If Cb_log.ListIndex < 0 Then
           Node10_Data1.log = 0
        Else
           Node10_Data1.log = CByte(Cb_log.ItemData(Cb_log.ListIndex))
        End If

        '保留
        Node10_Data1.reserved2(0) = 0

        '工作日设定
        For i = LBound(Node10_Data2.daytype) To UBound(Node10_Data2.daytype)
            Node10_Data2.daytype(i) = f_DateType(i)
        Next
        
        ' 父节点
        If Trim(T_nd_parent) = "" Then
            Node10_Data2.nd_parent = 0
        Else
            If (Val(T_nd_parent) > 32767 Or Val(T_nd_parent) < 256) And Val(T_nd_parent) <> 0 Then
                Message ("E035")
                T_nd_parent.SetFocus
                Exit Sub
            Else
                Node10_Data2.nd_parent = CInt(Trim(T_nd_parent.Text))
            End If
        End If
                
        ' 转节点ID
        For i = T_nd_trans.LBound To T_nd_trans.UBound
            If Trim(T_nd_trans(i)) = "" Then
                lv_nNewNode = 0
            Else
                If (Val(T_nd_trans(i)) > 32767 Or Val(T_nd_trans(i)) < 256) And Val(T_nd_trans(i)) <> 0 Then
                    Message ("E067")
                    T_nd_trans(i).SetFocus
                    Exit Sub
                Else
                    lv_nNewNode = CInt(Trim(T_nd_trans(i).Text))
                End If
            End If
            
            '' Sun added 2007-03-25
            If Node10_Data2.nd_daysec(i) <> lv_nNewNode Then
                Node10_Data2.nd_daysec(i) = lv_nNewNode
                Call gCallFlow.AutoChangeLineIndex(CByte(T_n_no), CInt(T_n_id), lv_nNewNode, i)
            End If
            
        Next
        
        ' 保留
        Node10_Data2.reserved1(0) = 0
        
        If Trim(Txt_Description.Text) = "" Or IsNull(Txt_Description) Then
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = LoadNationalResString(1147)
        Else
            gCallFlow.Node(gCallFlow.NodeSelectedID).Description = Trim(Txt_Description.Text)
        End If
        
        F_NodeData gCallFlow.NodeSelectedID, T_n_no
        gCallFlow.UpdateIvrRecord CInt(T_n_id), CByte(T_n_no.Text)
    
        f_DataChanged = False
   
    End If
    
    Unload Me
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next

SetMainFormItemsEnableWhenPropertyShow False

Dim i As Integer
Dim j As Integer

    '被访日志
    With Cb_log
        .AddItem LoadNationalResString(1178)
        .ItemData(.ListCount - 1) = 0
        For i = 1 To 16
            .AddItem Trim(Str(i)) & LoadNationalResString(1179)
            .ItemData(.ListCount - 1) = i
        Next
    End With
    
    T_n_id.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
    T_n_no.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
    
    ' 起止日期
    With cb_Year
        For j = 0 To 9
            '' Sun updated 2006-01-06
            ''' From
            '.AddItem Year(Now) + j
            '.ItemData(.ListCount - 1) = Year(Now) + j - 2000
            ''' To
            .AddItem Node10_Data1.startyear + 2000 + j
            .ItemData(.ListCount - 1) = Node10_Data1.startyear + j
        Next
    End With
    With cb_Month(0)
        For j = 1 To 12
            .AddItem Str(j)
            .ItemData(.ListCount - 1) = j
        Next
    End With
    With cb_Month(1)
        For j = 1 To 12
            .AddItem Str(j)
            .ItemData(.ListCount - 1) = j
        Next
    End With
    
    Cb_log.ListIndex = SearchItemDataIndex(Cb_log, CLng(Node10_Data1.log), 0)
    T_nd_parent.Text = Node10_Data2.nd_parent
    
    cb_Year.ListIndex = SearchListIndex(cb_Year, Str(Node10_Data1.startyear + 2000), 0)
    cb_Month(0).ListIndex = SearchListIndex(cb_Month(0), Str(Node10_Data1.startmonth), 0)
    cb_Month(1).ListIndex = SearchListIndex(cb_Month(1), Str(Node10_Data1.monthcount), 0)
    
    ' 日期类型
    For i = LBound(Node10_Data2.daytype) To UBound(Node10_Data2.daytype)
        f_DateType(i) = Node10_Data2.daytype(i)
    Next
    
    ' 转移节点
    For i = T_nd_trans.LBound To T_nd_trans.UBound
        T_nd_trans(i) = Node10_Data2.nd_daysec(i)
    Next
           
    Txt_Description.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
    
    ' 日历
    f_CurYear = Val(cb_Year.List(cb_Year.ListIndex))
    f_CurMonth = Val(cb_Month(0).List(cb_Month(0).ListIndex))
    f_Days = 0
    
    '' Fill month date
    FillMonthDate
    
    ' Data OK
    f_DataChanged = False
    LoadResStrings Me
End Sub

'' Fill month date
'
Private Sub FillMonthDate()
    Dim lv_Col As Integer
    Dim lv_row As Integer
    Dim lv_WeekDay As Integer
    Dim lv_DateLoop As Date
    Dim lv_DateCur As Date
    Dim lv_Index1 As Integer, lv_index2 As Byte
    Dim lv_MonthDays As Integer
    
    sprCalen.Row = 1
    sprCalen.Col = 3
    sprCalen.Text = Str(f_CurYear) & LoadNationalResString(1200)
    sprCalen.Col = 4
    sprCalen.Text = Str(f_CurMonth) & LoadNationalResString(1212)
    
    lv_DateLoop = CDate(f_CurYear & "-" & f_CurMonth & "-" & "01")
    lv_WeekDay = Weekday(lv_DateLoop)
    lv_MonthDays = 0
    
    For lv_row = 3 To 8
        sprCalen.Row = lv_row
        For lv_Col = 1 To 7
            lv_DateCur = lv_DateLoop + lv_Col - lv_WeekDay + (lv_row - 3) * 7
            sprCalen.Col = lv_Col
            If Month(lv_DateCur) <> f_CurMonth Then
                sprCalen.Text = ""
            Else
                lv_Index1 = Int((f_Days + lv_MonthDays) / 8)
                lv_index2 = (f_Days + lv_MonthDays) Mod 8
                If Get_Bit_Value(f_DateType(lv_Index1), lv_index2) Then
                    sprCalen.ForeColor = RGB(255, 0, 0)             ' Red
                Else
                    sprCalen.ForeColor = RGB(0, 0, 0)               ' Black
                End If
                sprCalen.Text = Day(lv_DateLoop + lv_Col - lv_WeekDay + (lv_row - 3) * 7)
                lv_MonthDays = lv_MonthDays + 1
            End If
        Next
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If f_DataChanged Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

'Private Sub sprCalen_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'If Row = 1 And Col = 1 Then MsgBox Col & ":" & Row
'End Sub
'
'Private Sub sprCalen_Click(ByVal Col As Long, ByVal Row As Long)
'If Row = 1 And Col = 1 Then MsgBox "sprCalen_Click " & Col & ":" & Row
'End Sub

Private Sub sprCalen_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim lv_endYear As Integer
    Dim lv_endMonth As Integer
    Dim lv_LongDate1 As Long
    Dim lv_LongDate2 As Long
    Dim lv_LongDate3 As Long

    If Row = 1 Then
        Select Case Col
        Case 1              '' First Month
            f_CurYear = Val(cb_Year.List(cb_Year.ListIndex))
            f_CurMonth = Val(cb_Month(0).List(cb_Month(0).ListIndex))
            f_Days = 0
            FillMonthDate

        Case 2              '' Previous Month
            If f_CurYear = Val(cb_Year.List(cb_Year.ListIndex)) And f_CurMonth = Val(cb_Month(0).List(cb_Month(0).ListIndex)) Then
                Message "M137"
                Exit Sub
            End If

            '' Pre-month
            f_CurMonth = f_CurMonth - 1
            If f_CurMonth <= 0 Then
                f_CurMonth = 12
                f_CurYear = f_CurYear - 1
            End If

            f_Days = f_Days - GetDaysInMonth(f_CurYear, f_CurMonth)
            FillMonthDate

        Case 5              '' This Month
            lv_endYear = Val(cb_Year.List(cb_Year.ListIndex))
            lv_endMonth = Val(cb_Month(0).List(cb_Month(1).ListIndex)) + Val(cb_Month(0).List(cb_Month(0).ListIndex)) - 1
            If lv_endMonth > 12 Then
                lv_endYear = lv_endYear + Int(lv_endMonth / 13)
                lv_endMonth = lv_endMonth Mod 13 + 1
            End If
            lv_LongDate1 = Val(cb_Year.List(cb_Year.ListIndex)) * 100 + Val(cb_Month(0).List(cb_Month(0).ListIndex))
            lv_LongDate2 = Year(Now) * 100 + Month(Now)
            lv_LongDate3 = CLng(lv_endYear) * 100 + lv_endMonth

            If lv_LongDate1 > lv_LongDate2 Or lv_LongDate2 > lv_LongDate3 Then
                Message "M138"
                Exit Sub
            End If

            '' Current month
            f_CurYear = Year(Now)
            f_CurMonth = Month(Now)
            f_Days = DateDiff("d", CDate(Str(f_CurYear) & "-" & Str(f_CurMonth) & "-01"), CDate(cb_Year.List(cb_Year.ListIndex) & "-" & cb_Month(0).List(cb_Month(0).ListIndex) & "-01"))
            If f_Days < 0 Then
                    f_Days = Abs(f_Days)
            End If
            FillMonthDate

        Case 6              '' Next Month
            lv_endYear = Val(cb_Year.List(cb_Year.ListIndex))
            lv_endMonth = Val(cb_Month(0).List(cb_Month(1).ListIndex)) + Val(cb_Month(0).List(cb_Month(0).ListIndex)) - 1
            If lv_endMonth > 12 Then
                lv_endYear = lv_endYear + Int(lv_endMonth / 13)
                lv_endMonth = lv_endMonth Mod 13 + 1
            End If
            lv_LongDate1 = CLng(f_CurYear) * 100 + f_CurMonth
            lv_LongDate2 = CLng(lv_endYear) * 100 + lv_endMonth
            If lv_LongDate1 >= lv_LongDate2 Then
                Message "M139"
                Exit Sub
            End If

            '' Next Month
            f_Days = f_Days + GetDaysInMonth(f_CurYear, f_CurMonth)
            f_CurMonth = f_CurMonth + 1
            If f_CurMonth > 12 Then
                f_CurMonth = 1
                f_CurYear = f_CurYear + 1
            End If

            FillMonthDate

        Case 7              '' Last Month
            lv_endYear = Val(cb_Year.List(cb_Year.ListIndex))
            lv_endMonth = Val(cb_Month(0).List(cb_Month(1).ListIndex)) + Val(cb_Month(0).List(cb_Month(0).ListIndex)) - 1
            If lv_endMonth > 12 Then
                lv_endYear = lv_endYear + Int(lv_endMonth / 13)
                lv_endMonth = lv_endMonth Mod 13 + 1
            End If

            '' Current month
            f_CurYear = lv_endYear
            f_CurMonth = lv_endMonth
            f_Days = DateDiff("d", CDate(cb_Year.List(cb_Year.ListIndex) & "-" & cb_Month(0).List(cb_Month(0).ListIndex) & "-01"), CDate(Str(f_CurYear) & "-" & Str(f_CurMonth) & "-01"))
            FillMonthDate
        End Select
    End If

End Sub

Private Sub sprCalen_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 3 Then Exit Sub
    
    Dim lv_Index1 As Integer, lv_index2 As Byte
    
    sprCalen.Row = Row
    sprCalen.Col = Col
    If sprCalen.Text <> "" Then
      
        lv_Index1 = Int((f_Days + Val(sprCalen.Text) - 1) / 8)
        lv_index2 = (f_Days + Val(sprCalen.Text) - 1) Mod 8
        
        If sprCalen.ForeColor = RGB(255, 0, 0) Then
            sprCalen.ForeColor = RGB(0, 0, 0)
            Call Set_Bit_Value(f_DateType(lv_Index1), lv_index2, 0)
        Else
            sprCalen.ForeColor = RGB(255, 0, 0)
            Call Set_Bit_Value(f_DateType(lv_Index1), lv_index2, 1)
        End If
        
        f_DataChanged = True
        
    End If
End Sub

Private Sub T_nd_parent_Change()
    f_DataChanged = True
End Sub

Private Sub T_nd_parent_GotFocus()
    T_nd_parent.SelStart = 0
    T_nd_parent.SelLength = Len(T_nd_parent)
End Sub

Private Sub T_nd_parent_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub T_nd_trans_Change(Index As Integer)
    f_DataChanged = True
End Sub

Private Sub T_nd_trans_GotFocus(Index As Integer)
    T_nd_trans(Index).SelStart = 0
    T_nd_trans(Index).SelLength = Len(T_nd_trans(Index))
End Sub

Private Sub T_nd_trans_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Txt_Description_Change()
    f_DataChanged = True
End Sub

