Attribute VB_Name = "Module1"
'*** SDI ���±�Ӧ�ó���ʾ��ȫ��ģ��               ***
'****************************************************
Option Explicit

' �洢�����Ӵ�����Ϣ���û���������
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type
   Public dd As Long
   Public GdeleteP_ID As Long
   Public GdeleteN_id As Long
   Public GN_id As Long
   Public GN_no As Long
   Public gstrData1 As String
   Public gstrData2 As String
Public M_Str_P_ModifyTime As String '״̬���������޸�ʱ��
Public M_Str_P_Version As String '״̬�������̰汾��
Public M_Str_P_User As String '״̬���������û�����

Public M_Cn As ADODB.Connection
Public FState As FormState              ' �û�����������
Public gFindString As String            ' ���������ı�
Public gFindCase As Integer             ' ���ִ�Сд��־
Public gFindDirection As Integer        ' ���������־
Public gCurPos As Integer               ' ���浱ǰ���λ��
Public gFirstTime As Integer            ' ��ʼλ��
Public Const ThisApp = "MDINote"        ' ע��� App ������
Public Const ThisKey = "Recent Files"   ' ע��� Key ������
Sub GetRecentFiles()
    ' ��������ʾ GetAllSettings �������÷������� Windows ע����з���ֵ�����顣
    ' ����������£�ע����������򿪵��ļ��б�ʹ�� SaveSetting ���������ʹ�õ��ļ�����
   ' ������� WriteRecentFiles ������ʹ��
    Dim i As Integer
    Dim varFiles As Variant ' �洢���ص�����ı���
    
    ' �� GetAllSettings ����ע����з������ʹ�õ��ļ���
    '  ThisApp �� ThisKey��ģ���ж���ĳ���
   
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
   ' For i = 0 To UBound(varFiles, 1)
   '     frmSDI.mnuRecentFile(0).Visible = True
   '     frmSDI.mnuRecentFile(i + 1).Caption = varFiles(i, 1)
    '    frmSDI.mnuRecentFile(i + 1).Visible = True
   ' Next i
End Sub
Sub ResizeNote()
    ' �ı������������ڲ�����
    If frmSDI.picToolbar.Visible Then
        frmSDI.tvwDB.Height = frmSDI.ScaleHeight - frmSDI.picToolbar.Height - 300
        frmSDI.tvwDB.Width = frmSDI.ScaleWidth - 7000
        frmSDI.tvwDB.Top = frmSDI.picToolbar.Height
        frmSDI.DataGrid1.Height = frmSDI.ScaleHeight - frmSDI.picToolbar.Height - 300
        frmSDI.DataGrid1.Left = frmSDI.ScaleWidth - 7000 + 150
        frmSDI.DataGrid1.Top = frmSDI.picToolbar.Height

    Else
        frmSDI.tvwDB.Height = frmSDI.ScaleHeight - 300
        frmSDI.tvwDB.Width = frmSDI.ScaleWidth - 7000
        frmSDI.tvwDB.Top = 0
        frmSDI.DataGrid1.Height = frmSDI.ScaleHeight - 300
        frmSDI.DataGrid1.Top = 0

    End If
End Sub

Sub WriteRecentFiles(OpenFileName)
    ' ������ʹ�� SaveSettings ��佫����򿪵��ļ���д��ϵͳע���
    ' SaveSettings ���Ҫ�������������������洢Ϊ�������ڱ�ģ���ڶ��塣
    ' GetRecentFiles ������ʹ�� GetAllSettings ������������������д洢���ļ�����
    Dim i As Integer
    Dim strFile As String
    Dim strKey As String

    ' ���ļ� RecentFile1 ���Ƹ� RecentFile2���ȵ�
    For i = 3 To 1 Step -1
        strKey = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, strKey)
        If strFile <> "" Then
            strKey = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, strKey, strFile
        End If
    Next i
  
    ' �����ڴ򿪵��ļ�д�����ʹ���ļ��б�ĵ�һ��
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

