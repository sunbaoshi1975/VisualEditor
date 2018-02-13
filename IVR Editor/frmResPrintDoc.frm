VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmResPrintDoc 
   Caption         =   "��Դ��ӡ"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmResPrintDoc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread spdResource 
      Height          =   5295
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   9855
      _Version        =   458752
      _ExtentX        =   17383
      _ExtentY        =   9340
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   1
      SpreadDesigner  =   "frmResPrintDoc.frx":030A
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
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   5595
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡԤ��(&P)"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "ȡ��(&Q)"
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Դ��Ŀ����"
      Height          =   195
      Index           =   1
      Left            =   3030
      TabIndex        =   5
      Tag             =   "1448"
      Top             =   210
      Width           =   1080
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
End
Attribute VB_Name = "frmResPrintDoc"
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
    FillResouse
    mdlcommon.ChangeMousePointer 0, True
    LoadResStrings Me
On Error GoTo 0
End Sub

'ˢ����Դ����б�
'
Public Sub FillResouse()
On Error GoTo BackDoor

    Dim itmX As ListItem                    ' ListItem ����
    Dim intCount As Integer                 ' ����������
    Dim blnShowItem As Boolean
    Dim lngItemID As Long
    Dim lv_intItemType As Integer
    Dim lv_intLoop As Integer
    
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
                    
                    spdResource.Row = spdResource.MaxRows
                    
                    '�� ItemID��Ϊ��, ������ "��Դ���"Ϊ���ֶ�
                    If Not IsNull(.Fields("P_ID")) Then
                        spdResource.Col = 1
                        spdResource.Text = Trim(.Fields("P_ID"))
                    End If
                    
                    '�� L_ID �ֶβ�Ϊ�գ������� "����" Ϊ���ֶΡ�
                    If Not IsNull(.Fields("L_ID")) Then
                        spdResource.Col = 2
                        spdResource.Text = Trim(.Fields("L_ID"))
                    End If
                    
                    '�� R_Type �ֶβ�Ϊ�գ������� "����" Ϊ���ֶΡ�
                    spdResource.Col = 3
                    spdResource.Text = gRID(lv_intItemType).Caption
                    
                    '�� Description �ֶβ�Ϊ�գ������� "��Դ����" Ϊ���ֶΡ�
                    If Not IsNull(.Fields("R_Description")) Then
                        spdResource.Col = 4
                        spdResource.Text = Trim(.Fields("R_Description"))
                    End If
                      
                    '�� PATH �ֶβ�Ϊ�գ������� "��Դ·��" Ϊ���ֶΡ�
                    If Not IsNull(.Fields("R_Path")) Then
                        spdResource.Col = 5
                        spdResource.Text = Trim(.Fields("R_Path"))
                    End If
                    
                    '�� CreateTime �ֶβ�Ϊ�գ������� "����ʱ��" Ϊ���ֶΡ�
                    If Not IsNull(.Fields("CreateTime")) Then
                        spdResource.Col = 6
                        spdResource.Text = Format(.Fields("CreateTime"), "yyyy-mm-dd hh:nn:ss")
                    End If
                    
                    '�� ModifyTime �ֶβ�Ϊ�գ������� "�޸�ʱ��" Ϊ���ֶΡ�
                    If Not IsNull(.Fields("ModifyTime")) Then
                        spdResource.Col = 7
                        spdResource.Text = Format(.Fields("ModifyTime"), "yyyy-mm-dd hh:nn:ss")
                    End If
                    
                    '�� R_Note �ֶβ�Ϊ�գ������� "��ע" Ϊ���ֶΡ�
                    If Not IsNull(.Fields("R_Note")) Then
                        spdResource.Col = 8
                        spdResource.Text = Trim(.Fields("R_Note"))
                    End If
                    
                    spdResource.MaxRows = spdResource.MaxRows + 1
                    
                End If
                
                '' Next Item
                .MoveNext
            Wend
        End If
    End With
    
    spdResource.Refresh
    
BackDoor:
   On Error GoTo 0
End Sub

Private Sub cmdQuit_Click()

    Unload Me
End Sub

'Modified @ Sep,6,07 ,��ֱ�Ӵ�ӡ��Ϊ��ӡԤ��
Private Sub cmdPrint_Click()
' Show Print Previw Form

    frmPrintPreview.Show vbModal
End Sub


