Attribute VB_Name = "Mdlsend"
Option Explicit

'���ͼ�������Դ��Ϣ
'
Public Function tcpSendLoadMsg() As Boolean
    Dim lv_PKG As SCtiMsi_Package
'    Dim lv_logonPKG As SIvrlogon_Package

    '' ׼����Ϣ
    lv_PKG.Msgheader.PackageLen = Def_PKGLENGTH
    lv_PKG.Msgheader.PackageNo = 100
    lv_PKG.Msgheader.PackageType = PKGTYP_CONTROL
    lv_PKG.Msgheader.Sender = USER_PROGRAM
    lv_PKG.Msgheader.Receiver = USER_MSG
    lv_PKG.command = 1
    lv_PKG.intData = Val(Frm_load.Txt_pid.Text)
    lv_PKG.bytData = Val(Frm_load.Txt_Group.Text)
    '' ����
    tcpSendLoadMsg = tcpSendMsg(lv_PKG)
   
End Function
Public Function tcpLogonLoadMsg() As Boolean
    Dim lv_PKG As SCtiMsi_Package
'    Dim lv_logonPKG As SIvrlogon_Package
''send logon informatiom
    lv_PKG.Msgheader.PackageLen = Def_PKGLENGTH
    lv_PKG.Msgheader.PackageNo = 0
    lv_PKG.Msgheader.PackageType = PKGTYP_DATA
    lv_PKG.Msgheader.Sender = USER_PROGRAM
    lv_PKG.Msgheader.Receiver = USER_MSG
    lv_PKG.strdata = Frm_load.Winskget.LocalIP
    lv_PKG.intData = Frm_load.Winskget.LocalPort
    lv_PKG.bytData = CByte(16)
    lv_PKG.command = 1
    tcpLogonLoadMsg = tcpSendMsgLogon(lv_PKG)
End Function
'modify: Scott date:2001/08/23
Private Function tcpSendMsgLogon(F_PKG As SCtiMsi_Package) As Boolean
On Error GoTo BackDoor:

    Dim lv_Data(Def_PKGLENGTH) As Byte
    
''1
    CopyMemory lv_Data(0), F_PKG.Msgheader.Sender, Def_PKGLENGTH
    Frm_load.sckMain.SendData lv_Data
        
    
    '' ����ֵ
    tcpSendMsgLogon = True
        
BackDoor:
    If Err.Number > 0 Then
        Debug.Print Err.Number & ": " & Err.Source & " - " & Err.Description
    End If
  
    On Error GoTo 0
    
End Function

'������Ϣ
'
Private Function tcpSendMsg(F_PKG As SCtiMsi_Package) As Boolean
On Error GoTo BackDoor:

    Dim lv_Data(Def_PKGLENGTH) As Byte
    
    CopyMemory lv_Data(0), F_PKG.Msgheader.Sender, Def_PKGLENGTH
'    Frm_load.sckMain.Connect
'    If Frm_load.sckMain.State <> sckClosed Then
    Frm_load.sckMain.SendData lv_Data
'    End If
    
    '' ����ֵ
    tcpSendMsg = True
        
BackDoor:
    If Err.Number > 0 Then
        Debug.Print Err.Number & ": " & Err.Source & " - " & Err.Description
    End If
  
    On Error GoTo 0
    
End Function
