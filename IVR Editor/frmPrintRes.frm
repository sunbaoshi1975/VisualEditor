VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintRes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Դ��ӡ"
   ClientHeight    =   6795
   ClientLeft      =   2565
   ClientTop       =   3000
   ClientWidth     =   10185
   Icon            =   "frmPrintRes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdQuit 
      Caption         =   "ȡ��(&Q)"
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡԤ��(&P)"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   6240
      Width           =   1455
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   615
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
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   5595
   End
   Begin MSComctlLib.ListView lstResource 
      Height          =   5565
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9816
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
         Text            =   "��Դ���"
         Object.Width           =   1729
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "LID"
         Object.Tag             =   "1744"
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Type"
         Object.Tag             =   "1745"
         Text            =   "����"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Description"
         Object.Tag             =   "1450"
         Text            =   "��Դ����"
         Object.Width           =   10232
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Path"
         Object.Tag             =   "1451"
         Text            =   "��Դ·��"
         Object.Width           =   7057
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "CreateTime"
         Object.Tag             =   "1369"
         Text            =   "����ʱ��"
         Object.Width           =   3000
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "ModifyTime"
         Object.Tag             =   "1370"
         Text            =   "�޸�ʱ��"
         Object.Width           =   3000
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Notes"
         Object.Tag             =   "1746"
         Text            =   "��ע"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Դ��ĿID"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Tag             =   "1447"
      Top             =   210
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Դ��Ŀ����"
      Height          =   195
      Index           =   1
      Left            =   3150
      TabIndex        =   3
      Tag             =   "1448"
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frmPrintRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
On Error Resume Next
    
    Dim lv_strName As String
    
    '' Title
'    If gSystem.intCurStep >= 0 Then
'        Me.Caption = LoadNationalResString(1663) & gRID(gSystem.intCurStep).Caption
'    Else
'        Me.Caption = LoadNationalResString(1663)
'    End If
    
    '' Resource Project
    txtResID = Trim(Str(gCallFlow.ResourceID))
    txtResName = gCallFlow.ResourceName
    
    '' Add Items
    RefreshRSListView
    mdlcommon.ChangeMousePointer 0, True
    LoadResStrings Me
On Error GoTo 0
End Sub

'ˢ����Դ����б�
'
Public Sub RefreshRSListView()
On Error GoTo BackDoor

    Dim itmX As ListItem                    ' ListItem ����
    Dim intCount As Integer                 ' ����������
    Dim blnShowItem As Boolean
    Dim lngItemID As Long
    Dim lv_intItemType As Integer
    Dim lv_intLoop As Integer

    '' ɾ����������
    lstResource.ListItems.Clear
    
    '' �����������������
    With gCallFlow.RS_Resource
        
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
                
                    ''' ʹ�� Add ��������µ� ListItem ��Ϊ���������ö���
                    ''' ʹ�������������ԡ�
                    Set itmX = lstResource.ListItems.Add(, , Right("00000" & Trim(Str(lngItemID)), 5))
                    intCount = intCount + 1                 'Tag ���Լ�����������
                    'Michael Note : Run time Error, .Key is an invaid key
                    'itmX.Key = .AbsolutePosition
                    itmX.Tag = lv_intItemType
                    
                    '�� L_ID �ֶβ�Ϊ�գ������� subitem 1 Ϊ���ֶΡ�
                    If Not IsNull(.Fields("L_ID")) Then
                        itmX.SubItems(1) = Trim(.Fields("L_ID"))
                    End If
                    
                    '�� R_Type �ֶβ�Ϊ�գ������� subitem 2 Ϊ���ֶΡ�
                    itmX.SubItems(2) = gRID(lv_intItemType).Caption
                    
                    '�� Description �ֶβ�Ϊ�գ������� subitem 3 Ϊ���ֶΡ�
                    If Not IsNull(.Fields("R_Description")) Then
                        itmX.SubItems(3) = Trim(.Fields("R_Description"))
                    End If
                      
                    '�� PATH �ֶβ�Ϊ�գ������� subitem 4 Ϊ���ֶΡ�
                    If Not IsNull(.Fields("R_Path")) Then
                        itmX.SubItems(4) = Trim(.Fields("R_Path"))
                    End If
                    
                    '�� CreateTime �ֶβ�Ϊ�գ������� subitem 5 Ϊ���ֶΡ�
                    If Not IsNull(.Fields("CreateTime")) Then
                        itmX.SubItems(5) = Format(.Fields("CreateTime"), "yyyy-mm-dd hh:nn:ss")
                    End If
                    
                    '�� ModifyTime �ֶβ�Ϊ�գ������� subitem 6 Ϊ���ֶΡ�
                    If Not IsNull(.Fields("ModifyTime")) Then
                        itmX.SubItems(6) = Format(.Fields("ModifyTime"), "yyyy-mm-dd hh:nn:ss")
                    End If
                    
                    '�� R_Note �ֶβ�Ϊ�գ������� subitem 7 Ϊ���ֶΡ�
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
    End With
    lstResource.Refresh
    
BackDoor:
   On Error GoTo 0
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

'Print the list
'Modified @ Sep,6,07 ,��ֱ�Ӵ�ӡ��Ϊ��ӡԤ��
Private Sub cmdPrint_Click()
'    Dim strPrintHead As String
'    strPrintHead = "��Դ��ĿID : " & txtResID & "          ��Դ��Ŀ���� : " & txtResName
'    gPrintListView lstResource, strPrintHead, Printer
'******************************************************
' Show Print Previw Form

    frmPrintPreview.Show vbModal
End Sub

