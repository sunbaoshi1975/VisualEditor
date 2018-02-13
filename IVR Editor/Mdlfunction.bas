Attribute VB_Name = "Mdlfunction"
Option Explicit

Public Function ByteArrayToString(f_Bytes() As Byte, ByVal f_Len As Integer) As String
On Error Resume Next

    Dim lv_Str As String
    Dim lv_lvp As Integer
    
    lv_Str = ""
    For lv_lvp = 0 To f_Len - 1
        lv_Str = lv_Str & Chr(f_Bytes(lv_lvp))
    Next
    
    ByteArrayToString = lv_Str
End Function

Public Sub StringToByteArray(f_str As String, f_Bytes() As Byte, ByVal f_Len As Integer)
On Error Resume Next

    Dim lv_lvp As Integer
    
    For lv_lvp = 0 To f_Len - 1
        If lv_lvp < Len(f_str) Then
            f_Bytes(lv_lvp) = Asc(Mid(f_str, lv_lvp + 1, 1))
        Else
            f_Bytes(lv_lvp) = 0
        End If
    Next
    
End Sub

'用途：    提取节点数据
'作者:     Scott
'创建日期：2001/04/23
'修改日期：2001/08/29
'描述：    Get Data into database Callcenter's tbivrprogram table
''提取接点数据 moidfy:Scott Data:2001/08/22
Public Function F_GetNodedata()
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

            gGetUserVar.IntP_id = gCallFlow.CallFlowID
            gGetUserVar.IntN_index = gCallFlow.NodeSelectedID
            gGetUserVar.IntN_id = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID
            gGetUserVar.IntN_no = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeNo
            gGetUserVar.IntN_page = gCallFlow.Node(gCallFlow.NodeSelectedID).InPage
            gGetUserVar.IntN_left = gCallFlow.Node(gCallFlow.NodeSelectedID).Left
            gGetUserVar.IntN_top = gCallFlow.Node(gCallFlow.NodeSelectedID).Top
            gGetUserVar.IntN_height = gCallFlow.Node(gCallFlow.NodeSelectedID).Height
            gGetUserVar.IntN_width = gCallFlow.Node(gCallFlow.NodeSelectedID).Width
            If Len(gCallFlow.Node(gCallFlow.NodeSelectedID).Data1) > 0 Or Not IsNull(gCallFlow.Node(gCallFlow.NodeSelectedID).Data1) Then
               gGetUserVar.StrN_data1 = gCallFlow.Node(gCallFlow.NodeSelectedID).Data1
            End If
            If Len(gCallFlow.Node(gCallFlow.NodeSelectedID).Data2) > 0 Or Not IsNull(gCallFlow.Node(gCallFlow.NodeSelectedID).Data2) Then
               gGetUserVar.StrN_data2 = gCallFlow.Node(gCallFlow.NodeSelectedID).Data2
            End If
            If Not IsNull(gCallFlow.Node(gCallFlow.NodeSelectedID).Description) Or Len(gCallFlow.Node(gCallFlow.NodeSelectedID).Description) > 0 Then
               gGetUserVar.StrN_description = gCallFlow.Node(gCallFlow.NodeSelectedID).Description
            End If
    
    'Debug.Print "Expla: "; gGetUserVar.IntN_id, gCallFlow.NodeSelectedID, gGetUserVar.StrN_data1, gGetUserVar.StrN_data2
    'Debug.Print "Expla: "; gCallFlow.NodeSelectedID, gCallFlow.Node(gCallFlow.NodeSelectedID).Data1, gCallFlow.Node(gCallFlow.NodeSelectedID).Data2
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Function

Public Function KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Function
    
    If KeyAscii = 13 Then SendKeys "{TAB}": Exit Function
    
    '' Sun added 2005-09-12
    If KeyAscii < 30 Or KeyAscii > 127 Then Exit Function
    
    If KeyAscii <> vbKeyTab Then
        If Not (KeyAscii <= 57 And KeyAscii >= 48) Then
            KeyAscii = 0
        End If
    End If
End Function

'用途：解析节点数据
'作者:     Scott
'创建日期：2001/04/23
'修改日期：2001/08/29
'描述：    解析N_data1和N_data2数据
Public Function F_ExplainNodeData(lv_nIndex As Integer)
On Error Resume Next

   Dim lv_Gdata1(DEF_NODE_DATA1_LEN) As Byte
   Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
   Dim lv_loop
   
   Call F_GetNodedata
 
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
   
   If gCallFlow.Node(lv_nIndex).Data2 <> "" Then
      For lv_loop = 0 To 63
          lv_Gdata2(lv_loop) = AscB(MidB(gCallFlow.Node(lv_nIndex).Data2, lv_loop + 1, 1))
      Next
   End If
   If gCallFlow.Node(lv_nIndex).Data1 <> "" Then
         For lv_loop = 0 To 12
             lv_Gdata1(lv_loop) = AscB(MidB(gCallFlow.Node(lv_nIndex).Data1, lv_loop + 1, 1))
         Next
   End If
   
   Select Case Val(gCallFlow.Node(lv_nIndex).NodeNo)
      Case 0
           CopyMemory Node0_Data2.key_repeat, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 1
           CopyMemory Node1_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
      Case 2
           CopyMemory Node2_Data2.uservar(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 6
           '' Sun added 2012-11-23
           CopyMemory Node6_Data1.Sleep, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node6_Data2.nd_goto, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 7
           CopyMemory Node7_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node7_Data2.vox_userid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 8
           CopyMemory Node8_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node8_Data2.vox_password, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 9
           CopyMemory Node9_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node9_Data2.workday, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 10
           CopyMemory Node10_Data1.maincalendar, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node10_Data2.daytype(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      
      ''--------------------------------
      '' Sun added 2004-12-30
      Case 16
           CopyMemory Node16_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node16_Data2.var_value(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      ''--------------------------------
      
      Case 17
           CopyMemory Node17_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node17_Data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 18
           CopyMemory Node18_Data1.seperator, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node18_Data2.typeflags(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 19
           CopyMemory Node19_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node19_Data2.delaytime, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 20
           CopyMemory Node20_Data1.reserved1, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node20_Data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 21
           CopyMemory Node21_Data1.reserved1, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node21_Data2.vox_pred, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 22
           CopyMemory Node22_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node22_Data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 23
           CopyMemory Node23_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node23_Data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
           
      ''--------------------------------
      '' Sun added 2004-12-30
      Case 28
           CopyMemory Node28_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node28_Data2.vox_string, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      ''--------------------------------
           
      Case 40
           CopyMemory Node40_Data1.rectime, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node40_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 41
           CopyMemory Node41_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node41_Data2.vox_play(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 50
           CopyMemory Node50_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node50_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 51
           CopyMemory Node51_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node51_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      
      ''--------------------------------
      '' Sun added 2006-12-31
      Case 55
           CopyMemory Node55_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node55_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      ''--------------------------------
      
      Case 60
           CopyMemory Node60_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node60_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 61
           CopyMemory Node61_Data1.maxwait, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node61_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 62
           CopyMemory Node62_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node62_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 63
           CopyMemory Node63_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node63_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 69
           CopyMemory Node69_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node69_Data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      
      '-------------------------------------------
      ' Sun added 2005-06-27
      Case 70
           CopyMemory Node70_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node70_Data2.routepointid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      
      Case 71
           CopyMemory Node71_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node71_Data2.agentid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      '-------------------------------------------
      
      Case 90
           CopyMemory Node90_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node90_Data2.predial(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      
      '-------------------------------------------
      ' Sun added 2005-05-26
      Case 91
           CopyMemory Node91_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node91_Data2.vox_talklen, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      '-------------------------------------------
      
      ''--------------------------------
      '' Sun added 2005-03-15
      Case 96
           CopyMemory Node96_Data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node96_Data2.command, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      ''--------------------------------
      
      Case 100
           CopyMemory Node100_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node100_Data2.dll_fid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 101
           CopyMemory Node101_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node101_Data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 102
           CopyMemory Node102_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node102_Data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
      Case 255
           CopyMemory Node255_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
           CopyMemory Node255_Data2.StartNode, lv_Gdata2(0), DEF_NODE_DATA2_LEN
      End Select
'      gGetUserVar.StrN_data1 = ""
'      gGetUserVar.StrN_data2 = ""
      
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Function

' 创建变量节点
'
Public Sub F_CreateVar(var As Integer)
On Error Resume Next

Dim lv_nCount As Integer

lv_nCount = Int((var - 1) / 4) + 1
If gCallFlow.VarNodeCount = lv_nCount Then
'' No change
   Exit Sub
ElseIf gCallFlow.VarNodeCount > lv_nCount Then
'' Delete Nodes
    While gCallFlow.VarNodeCount > lv_nCount
        gCallFlow.DeleteNode gCallFlow.SysNodeID
    Wend
Else
'' Add Nodes
    While gCallFlow.VarNodeCount < lv_nCount
        gCallFlow.AddVariableNode
    Wend
End If

End Sub

' 生成连线描述
'
Public Function F_GetLineCaption(ByVal f_nNodeNo As Integer, ByVal f_nLineIndex As Byte) As String

    Dim lv_strCaption As String
    
    lv_strCaption = ""
    
    Select Case f_nNodeNo
    Case 6
         
    Case 7
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1710)
        Else
            lv_strCaption = LoadNationalResString(1709)
        End If
        
    Case 8
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1710)
        Else
            lv_strCaption = LoadNationalResString(1709)
        End If
        
    Case 9
        lv_strCaption = Trim(Str(f_nLineIndex))
    
    Case 10
        If f_nLineIndex = 0 Then
            '' 工作日
            lv_strCaption = LoadNationalResString(1204)
        ElseIf f_nLineIndex = 1 Then
            '' 节假日
            lv_strCaption = LoadNationalResString(1205)
        Else
            '' 其他
            lv_strCaption = LoadNationalResString(1206)
        End If
    
    Case 16
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1708)
        Else
            lv_strCaption = LoadNationalResString(1707)
        End If
    
    Case 17
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1710)
        Else
            lv_strCaption = LoadNationalResString(1709)
        End If
    
    Case 18
         
    Case 19
    Case 20
    
    Case 21
    
    Case 22
        If f_nLineIndex >= 0 And f_nLineIndex <= 11 Then
            If f_nLineIndex < 10 Then
                lv_strCaption = Trim(Str(f_nLineIndex))
            ElseIf f_nLineIndex = 10 Then
                lv_strCaption = "*"
            Else
                lv_strCaption = "#"
            End If
        Else
            lv_strCaption = LoadNationalResString(1710)
        End If
    
    Case 23
    
    Case 28
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1710)
        Else
            lv_strCaption = LoadNationalResString(1709)
        End If
    
    Case 40
    Case 41
    
    Case 50
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1710)
        Else
            lv_strCaption = LoadNationalResString(1709)
        End If
    
    Case 51
         
    Case 55
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1710)
        Else
            lv_strCaption = LoadNationalResString(1709)
        End If
    
    Case 60
    Case 61
    Case 62
    Case 63
    Case 69
    
    Case 70
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1708)
        ElseIf f_nLineIndex = 1 Then
            lv_strCaption = LoadNationalResString(1707)
        Else
            lv_strCaption = LoadNationalResString(1710)
        End If
    
    Case 71
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1708)
        ElseIf f_nLineIndex = 1 Then
            lv_strCaption = LoadNationalResString(1707)
        Else
            lv_strCaption = LoadNationalResString(1710)
        End If
    
    Case 80
    
    Case 90
        If f_nLineIndex = 0 Then
            lv_strCaption = LoadNationalResString(1710)
        Else
            lv_strCaption = LoadNationalResString(1709)
        End If
    
    Case 91
    
    Case 96
        If f_nLineIndex = 0 Then
        Else
            lv_strCaption = LoadNationalResString(1711)
        End If
    
    Case 100
    Case 101
    Case 102
    Case Else
    End Select
    
    F_GetLineCaption = lv_strCaption

End Function

' 节点缺省值
'
Public Function F_NodeDefaultInfo(ByVal f_Index As Integer, ByVal f_NodeNo As Byte)
On Error Resume Next

    Dim lv_lvp As Integer
    
    Select Case f_NodeNo
    Case 0      '全程转移规则
        Node0_Data1.reserved1(0) = 0
        Node0_Data1.Languages = 1
        Node0_Data1.MajorVer = Def_CallFlow_MajorVersion
        Node0_Data1.MinorVer = Def_CallFlow_MinorVersion
        Node0_Data2.reserved1 = 0
        Node0_Data2.reserved2(0) = 0
        Node0_Data2.reserved3(0) = 0
        Node0_Data2.key_repeat = Default.key_repeat
        Node0_Data2.key_return = Default.key_return
        Node0_Data2.key_root = Default.key_root
        Node0_Data2.MainCOM = 0
        Node0_Data2.LogSwitchOff = 0
        Node0_Data2.ResourceProject = 0
        Node0_Data2.nd_SysSendData = 0
        Node0_Data2.nd_BeforeHookOn = 0
        Node0_Data2.nd_parent = Default.nd_parent
        Node0_Data2.nd_root = Default.nd_root
        
    Case 1      'Buffer定义日志
        Node1_Data1.reserved1(0) = 0
        Node1_Data1.uservars = Defaultuservar.uservars
        Node1_Data1.reserved2 = 0
            
    Case 2      'Buffer定义变量
        Node2_Data1.reserved1(0) = 0
        Node2_Data2.uservar(0) = 0
    
    Case 6      '无条件转移
        '' Sun added 2012-11-23
        Node6_Data1.Sleep = 0
        Node6_Data1.reserved1(0) = 0
        Node6_Data2.nd_goto = 0
        Node6_Data2.nd_parent = 0
        Node6_Data2.reserved1(0) = 0
        Node6_Data2.reserved2(0) = 0
    
    Case 7      '身份验证
        Node7_Data1.key_term = 35   ' #
        Node7_Data1.log = 3
        Node7_Data1.maxpassword = 9
        Node7_Data1.maxtrytime = 3
        Node7_Data1.maxuserid = 9
        Node7_Data1.reserved1(0) = 0
        Node7_Data1.timeout = 15
        Node7_Data1.var_password = 0
        Node7_Data1.var_result = 0
        Node7_Data1.var_trytime = 0
        Node7_Data1.var_userid = 0
        Node7_Data2.com_iid = 0
        Node7_Data2.nd_fail = 0
        Node7_Data2.nd_parent = 0
        Node7_Data2.reserved1(0) = 0
        Node7_Data2.reserved2(0) = 0
        Node7_Data2.reserved3(0) = 0
        Node7_Data2.vox_password = 0
        Node7_Data2.vox_tryagain = 0
        Node7_Data2.vox_userid = 0
    
    Case 8      '修改口令
        Node8_Data1.key_term = 35   ' #
        Node8_Data1.log = 3
        Node8_Data1.maxpassword = 9
        Node8_Data1.maxtrytime = 3
        Node8_Data1.reserved1 = 0
        Node8_Data1.reserved2 = 0
        Node8_Data1.reserved3(0) = 0
        Node8_Data1.timeout = 15
        Node8_Data1.var_password = 0
        Node8_Data1.var_result = 0
        Node8_Data1.var_trytime = 0
        Node8_Data2.com_iid = 0
        Node8_Data2.nd_fail = 0
        Node8_Data2.nd_parent = 0
        Node8_Data2.nd_succeed = 0
        Node8_Data2.reserved1(0) = 0
        Node8_Data2.reserved2(0) = 0
        Node8_Data2.reserved3(0) = 0
        Node8_Data2.vox_confirm = 0
        Node8_Data2.vox_password = 0
        Node8_Data2.vox_succeed = 0
        Node8_Data2.vox_tryagain = 0
    
    Case 9      '时间分支
        Node9_Data1.log = 0
        Node9_Data1.reserved1(0) = 0
        Node9_Data1.reserved2(0) = 0
        Node9_Data2.nd_parent = 0
        Node9_Data2.nd_sparetime = 0
        For lv_lvp = 0 To 5
            Node9_Data2.nd_timesec(lv_lvp) = 0
        Next
        Node9_Data2.reserved1(0) = 0
        For lv_lvp = 0 To 23
            Node9_Data2.timesec(lv_lvp) = 0
        Next
        Node9_Data2.workday = 62          '' Mon - Fri
        Node9_Data2.worktime = 0
        
    Case 10      '工作日设定
        Dim lv_DateLoop As Date
        
        Node10_Data1.maincalendar = 0
        Node10_Data1.startyear = Year(Now) Mod 100
        Node10_Data1.startmonth = Month(Now)
        Node10_Data1.monthcount = 1
        Node10_Data1.log = 0
        Node10_Data1.reserved1 = 0
        Node10_Data1.reserved2(0) = 0
        
        lv_DateLoop = CDate(Year(Now) & "-" & Month(Now) & "-" & "01")
        For lv_lvp = 0 To 365
            If Weekday(lv_DateLoop + lv_lvp) Mod 6 = 1 Then
                Call Set_Bit_Value(Node10_Data2.daytype(Int(lv_lvp / 8)), lv_lvp Mod 8, 1)
            Else
                Call Set_Bit_Value(Node10_Data2.daytype(Int(lv_lvp / 8)), lv_lvp Mod 8, 0)
            End If
        Next
        Node10_Data2.reserved1(0) = 0
        Node10_Data2.nd_parent = 0
        For lv_lvp = LBound(Node10_Data2.nd_daysec) To UBound(Node10_Data2.nd_daysec)
            Node10_Data2.nd_daysec(lv_lvp) = 0
        Next
        
    ''--------------------------------
    '' Sun added 2004-12-30
    Case 16     '条件分支
        Node16_Data1.log = 0
        Node16_Data1.logic = 0
        Node16_Data1.var_id = 0
        Node16_Data1.convert = 0
        Node16_Data1.param1 = 0
        Node16_Data1.param2 = 0
        Node16_Data1.reserved1(0) = 0
        Node16_Data1.reserved2(0) = 0
        Node16_Data2.var_value(0) = 0
        Node16_Data2.nd_parent = 0
        Node16_Data2.nd_succ = 0
        Node16_Data2.nd_fail = 0
        Node16_Data2.reserved3(0) = 0
    ''--------------------------------
        
    Case 17     '选择语言
        Node17_Data1.timeout = 15
        Node17_Data1.maxtrytime = 3
        Node17_Data1.var_lang = 0
        Node17_Data1.reserved1(0) = 0
        For lv_lvp = LBound(Node17_Data2.lang) To UBound(Node17_Data2.lang)
            Node17_Data2.lang(lv_lvp) = 0
        Next
        Node17_Data2.nd_fail = 0
        Node17_Data2.nd_parent = 0
        Node17_Data2.nd_succ = 0
        Node17_Data2.vox_play = 0
        Node17_Data2.reserved1(0) = 0
        Node17_Data2.reserved2(0) = 0
        
    Case 18     '发送数据
        Node18_Data1.seperator = Asc(";")
        Node18_Data1.reserved1(0) = 0
        Node18_Data2.typeflags(0) = 0
        Node18_Data2.prefix1(0) = 0
        Node18_Data2.prefix2(0) = 0
        Node18_Data2.valueid(0) = 0
        Node18_Data2.nd_parent = 0
        Node18_Data2.nd_child = 0
    
    Case 19     '无操作
        Node19_Data1.reserved1(0) = 0
        'Michael Added @ Jul,10,07
        Node19_Data1.leavequeue = 0
        'Add To Here
        Node19_Data2.delaytime = 0
        Node19_Data2.nd_parent = 0
        Node19_Data2.reserved1(0) = 0
        Node19_Data2.reserved2(0) = 0
    
    Case 255    '节点连线
        Node255_Data1.reserved1(0) = 0
        Node255_Data2.Color = 0
        Node255_Data2.EndNode = 0
        Node255_Data2.Index = 0
        Node255_Data2.reserved1(0) = 0
        Node255_Data2.StartNode = 0
        Node255_Data2.Style = 1
        Node255_Data2.Width = 1
    
    Case 20     '放音挂机
        Node20_Data1.log = 0
        Node20_Data1.playclear = 0
        Node20_Data1.reserved1 = 0
        Node20_Data1.reserved2(0) = 0
        Node20_Data1.reserved3(0) = 0
        Node20_Data2.nd_parent = 0
        Node20_Data2.reserved1(0) = 0
        Node20_Data2.reserved2(0) = 0
        Node20_Data2.vox_play = 0
        
    Case 21     '放音继续
        Node21_Data1.breakkey = 78      ' N - 不中断
        Node21_Data1.log = 0
        Node21_Data1.playclear = 0
        Node21_Data1.playtype = 0       ' 数字
        Node21_Data1.reserved1 = 0
        Node21_Data1.reserved2(0) = 0
        Node21_Data1.usevar = 0
        Node21_Data2.com_iid = 0
        Node21_Data2.nd_child = 0
        Node21_Data2.nd_parent = 0
        Node21_Data2.reserved1(0) = 0
        Node21_Data2.reserved2(0) = 0
        Node21_Data2.reserved3(0) = 0
        Node21_Data2.vox_pred = 0
        Node21_Data2.vox_succ = 0
            
    Case 22     '放音等待按键
        Node22_Data1.breakkey = 78      ' N
        Node22_Data1.getlength = 1
        Node22_Data1.log = 2
        Node22_Data1.maxinterval = 5
        Node22_Data1.playclear = 0
        Node22_Data1.reserved1(0) = 0
        Node22_Data1.timeout = 15
        Node22_Data1.var_key = 0
        Node22_Data1.maxtrytime = 3     '' Sun added 2007-03-20
        For lv_lvp = 0 To 11
            Node22_Data2.nd_key(lv_lvp) = 0
        Next
        Node22_Data2.nd_nodefail = 0
        Node22_Data2.nd_parent = 0
        Node22_Data2.reserved1(0) = 0
        Node22_Data2.reserved2(0) = 0
        Node22_Data2.vox_nodefail = 0
        Node22_Data2.vox_play = 0
                        
    Case 23     '放音转移
        Node23_Data1.breakkey = 65  'All Keys
        Node23_Data1.log = 0
        Node23_Data1.playclear = 1
        Node23_Data1.reserved1(0) = 0
        Node23_Data1.var_play = 0
        Node23_Data1.reserved2(0) = 0
        Node23_Data1.timeout = 15
        If gCallFlow.Node(gCallFlow.NodeSelectedID).NodeID = 256 Then
            Node23_Data2.nd_goto = 257
        Else
            Node23_Data2.nd_goto = 0
        End If
        Node23_Data2.nd_parent = 0
        Node23_Data2.reserved1(0) = 0
        Node23_Data2.reserved2(0) = 0
        Node23_Data2.vox_play = 0
        
    ''--------------------------------
    '' Sun added 2004-12-30
    Case 28     'TTS 放音
        Node28_Data1.breakkey = 65  'All Keys
        Node28_Data1.log = 0
        Node28_Data1.playclear = 1
        Node28_Data1.reserved1 = 0
        Node28_Data1.usevar = 0
        Node28_Data1.reserved2(0) = 0
        Node28_Data1.timeout = 15
        Node28_Data2.com_iid = 0
        Node28_Data2.nd_succ = 0
        Node28_Data2.nd_fail = 0
        Node28_Data2.nd_parent = 0
        Node28_Data2.reserved1(0) = 0
        Node28_Data2.reserved2(0) = 0
        Node28_Data2.reserved3(0) = 0
        Node28_Data2.vox_string = 0
        Node28_Data2.vox_alter = 0
        Node28_Data2.playtype = 0
    ''--------------------------------
        
    Case 40     '建立留言
        Node40_Data1.breakkey = 65      ' All keys
        Node40_Data1.log = 0
        Node40_Data1.playclear = 0
        Node40_Data1.rectime = 120
        Node40_Data1.maxsilencetime = 5
        Node40_Data1.var_agent = 0
        Node40_Data1.var_appfield(0) = 0
        Node40_Data1.var_appfield(1) = 0
        Node40_Data1.var_appfield(2) = 0
        Node40_Data1.var_filename = 0
        Node40_Data1.vmsclass = 0
        Node40_Data1.MinRecLength = 5
        Node40_Data2.nd_child = 0
        Node40_Data2.nd_parent = 0
        'Michael add @7-4-07 for explane the record time length
        Node40_Data2.rectime_ho = 0
        
        ' Sun added 2009-07-24
        Node40_Data1.toneoff = 0
        Node40_Data2.var_notifyintvl = 0
        Node40_Data2.var_rectime = 0
        
        ' Sun added 2012-04-18
        Node40_Data2.NotifyPL = 0
        
        Node40_Data2.reserved1(0) = 0
        Node40_Data2.reserved2(0) = 0
        Node40_Data2.vox_op = 0
        Node40_Data2.recfiletype = gSystem.intRecFileType     'Michael Note : should not init here
    
    Case 41     '察看留言
        Node41_Data1.breakkey = 65      ' All keys
        Node41_Data1.log = 0
        Node41_Data1.playclear = 0
        Node41_Data1.reserved1 = 0
        Node41_Data1.reserved2(0) = 0
        Node41_Data1.timeout = 15
        Node41_Data1.var_agent = 0
        Node41_Data1.vmstype = 0
        Node41_Data1.closewhencheck = 0
        Node41_Data1.vmsclass = 0
        Node41_Data2.key_op(0) = 49     ' "1"
        Node41_Data2.key_op(1) = 50     ' "2"
        Node41_Data2.key_op(2) = 51     ' "3"
        Node41_Data2.key_op(3) = 52     ' "4"
        Node41_Data2.key_op(4) = 53     ' "5"
        Node41_Data2.key_op(5) = 54     ' "6"
        Node41_Data2.key_op(6) = 57     ' "9"
        Node41_Data2.key_op(7) = 56     ' "8"
        Node41_Data2.nd_child = 0
        Node41_Data2.nd_parent = 0
        Node41_Data2.reserved1(0) = 0
        Node41_Data2.reserved2(0) = 0
        For lv_lvp = 0 To 3
            Node41_Data2.vox_play(lv_lvp) = 0
        Next
    
    Case 50     '简单传真
        Node50_Data1.timeout = 180
        Node50_Data1.filenametype = 0
        Node50_Data1.trytimes = 3
        Node50_Data1.record_cdr = 1
        Node50_Data1.log = 0
        Node50_Data1.var_faxfile = 0
        Node50_Data1.var_fromno = 0
        Node50_Data1.var_tono = 0
        Node50_Data1.var_result = 0
        Node50_Data1.var_appfield(0) = 0
        Node50_Data1.var_appfield(1) = 0
        Node50_Data1.var_appfield(2) = 0
        Node50_Data2.vox_op = 0
        Node50_Data2.fax_fileid = 0
        Node50_Data2.header_id = 0
        Node50_Data2.nd_parent = 0
        Node50_Data2.nd_succ = 0
        Node50_Data2.nd_fail = 0
        Node50_Data2.reserved1(0) = 0
        Node50_Data2.reserved2(0) = 0
        
    Case 51     'TTF传真
        Node51_Data1.log = 0
        Node51_Data1.reserved1(0) = 0
        Node51_Data1.reserved2(0) = 0
        Node51_Data1.timeout = 30
        Node51_Data2.com_iid = 0
        Node51_Data2.fax_format = 0
        Node51_Data2.fax_logo = 0
        Node51_Data2.nd_child = 0
        Node51_Data2.nd_parent = 0
        Node51_Data2.reserved1(0) = 0
        Node51_Data2.reserved2(0) = 0
        Node51_Data2.reserved3(0) = 0
        Node51_Data2.vox_op = 0
    
    ''--------------------------------
    '' Sun added 2006-12-31
    Case 55     '传真接收
        Node55_Data1.timeout = 180
        Node55_Data1.filenametype = 0
        Node55_Data1.var_faxfile = 0
        Node55_Data1.record_cdr = 1
        Node55_Data1.log = 0
        Node55_Data1.var_fromno = 0
        Node55_Data1.var_tono = 0
        Node55_Data1.var_extno = 0
        Node55_Data1.var_result = 0
        Node55_Data1.var_appfield(0) = 0
        Node55_Data1.var_appfield(1) = 0
        Node55_Data1.var_appfield(2) = 0
        Node55_Data2.vox_op = 0
        Node55_Data2.fax_fileid = 0
        Node55_Data2.nd_parent = 0
        Node55_Data2.nd_succ = 0
        Node55_Data2.nd_fail = 0
        Node55_Data2.reserved1(0) = 0
        Node55_Data2.reserved2(0) = 0
    ''--------------------------------
    
    Case 60     '转接座席
        Node60_Data1.agentid = 0
        Node60_Data1.getlength = 5
        Node60_Data1.log = 0
        Node60_Data1.looptimes = 3
        Node60_Data1.agentinfo = 0
        Node60_Data1.reserved1(0) = 0
        Node60_Data1.switchtype = 0
        Node60_Data1.timeout = 15
        Node60_Data1.var_key = 0
        Node60_Data2.nd_busy = 0
        Node60_Data2.nd_nobody = 0
        Node60_Data2.nd_ok = 0
        Node60_Data2.nd_parent = 0
        'Mike Added @ 2008-5-27
        Node60_Data2.length_agentinfo = 4
        Node60_Data2.reserved1(0) = 0
        Node60_Data2.reserved2(0) = 0
        Node60_Data2.vox_busy = 0
        Node60_Data2.vox_nobody = 0
        Node60_Data2.vox_ok = 0
        Node60_Data2.vox_op = 0
        Node60_Data2.vox_sw = 0
        Node60_Data2.vox_wt = 0
    
    Case 61     '转接座席组
        Node61_Data1.maxwait = 30
        Node61_Data1.toacd = 0
        Node61_Data1.log = 0
        Node61_Data1.agentinfo = 0
        Node61_Data1.looptimes = 3
        'Michael Modified @ Jul,10,07
        'Node61_Data1.reserved1 = 0
        'Michael Added @7-4-07
        Node61_Data1.usevar = 0
        Node61_Data1.switchtype = 0
        Node61_Data1.waitmethod = 0
        'Michael Added @ July,0,07
        Node61_Data1.readEWT = 0
        Node61_Data1.var_userid = 0
        Node61_Data1.var_loginid = 0
        'Michael Added @ Jul,10,07
        Node61_Data1.waitansto = 0
        '------  Add End -----
        'Michael Modify @ July,9,07
        'Node61_Data1.reserved2(0) = 0
        Node61_Data2.nd_busy = 0
        Node61_Data2.nd_nobody = 0
        Node61_Data2.nd_ok = 0
        Node61_Data2.nd_parent = 0
        'Michael Add @ 7-5-07
        Node61_Data2.nd_wait = 0
        'Add End
        Node61_Data2.acddn(0) = 0
        Node61_Data2.routepointid = 0
        'Mike Added @ 2008-5-27
        Node61_Data2.length_agentinfo = 4
        Node61_Data2.reserved1(0) = 0
        Node61_Data2.reserved2(0) = 0
        Node61_Data2.vox_busy = 0
        Node61_Data2.vox_nobody = 0
        Node61_Data2.vox_ok = 0
        Node61_Data2.vox_op = 0
        Node61_Data2.vox_sw = 0
        Node61_Data2.vox_wt = 0
        'Michael Added @ Jul,10,07
        Node61_Data2.waitansto_hi = 0
        
    
    Case 62     '发起会议
        Node62_Data1.timeout = 60
        Node62_Data1.log = 0
        Node62_Data1.looptimes = 3
        Node62_Data1.reserved1 = 0
        Node62_Data1.reserved2(0) = 0
        Node62_Data1.usevar = 0
        Node62_Data1.waitansto = 16
        Node62_Data1.var_waitansto = 0      '' Sun added 2014-01-29
                
        Node62_Data2.nd_failed = 0
        Node62_Data2.nd_noans = 0
        Node62_Data2.nd_ok = 0
        Node62_Data2.nd_parent = 0
        Node62_Data2.reserved1(0) = 0
        Node62_Data2.reserved2(0) = 0
        Node62_Data2.DialNo(0) = 0
        Node62_Data2.predial(0) = 0
        Node62_Data2.vox_ok = 0
        Node62_Data2.vox_op = 0
        Node62_Data2.vox_sw = 0
        Node62_Data2.vox_wt = 0
        Node62_Data2.vox_noconf = 0
        Node62_Data2.vox_noans = 0
        Node62_Data2.vox_ans = 0
        Node62_Data2.vox_syserror = 0
            
    Case 63     '增强转接座席组
        Node63_Data1.log = 0
        Node63_Data1.looptimes = 3
        Node63_Data1.reserved1(0) = 0
        Node63_Data1.reserved2(0) = 0
        Node63_Data1.usevar = 0
        Node63_Data2.nd_busy = 0
        Node63_Data2.nd_nobody = 0
        Node63_Data2.nd_ok = 0
        Node63_Data2.nd_parent = 0
        Node63_Data2.reserved1(0) = 0
        Node63_Data2.reserved2(0) = 0
        Node63_Data2.vox_busy = 0
        Node63_Data2.vox_nobody = 0
        Node63_Data2.vox_ok = 0
        Node63_Data2.vox_op = 0
        Node63_Data2.vox_sw = 0
        Node63_Data2.vox_wt = 0
            
    Case 69     '转虚拟分机
        Node69_Data1.log = 0
        Node69_Data1.reserved1(0) = 0
        Node69_Data1.reserved3(0) = 0
        Node69_Data1.usevar = 0
        Node69_Data1.switchtype = 0
        Node69_Data1.vagency = 0
        Node69_Data2.nd_child = 0
        Node69_Data2.maxtry = 3
        Node69_Data2.tryinterval = 1000
        Node69_Data2.nd_parent = 0
        Node69_Data2.reserved1(0) = 0
        Node69_Data2.reserved2(0) = 0
        Node69_Data2.vox_op = 0
            
    ''--------------------------------
    '' Sun added 2005-06-27
    Case 70     '查询路由点
        Node70_Data1.timeout = 5
        Node70_Data1.logic = 0
        Node70_Data1.log = 0
        Node70_Data1.usevar = 0
        Node70_Data1.var_result = 0
        Node70_Data1.paramindex = 0
        Node70_Data1.querytype = 0
        Node70_Data1.reserved1 = 0
        Node70_Data1.reserved2(0) = 0
        Node70_Data2.routepointid = 0
        Node70_Data2.com_iid = 0
        Node70_Data2.comparedvalue = 0
        Node70_Data2.nd_parent = 0
        Node70_Data2.nd_yes = 0
        Node70_Data2.nd_no = 0
        Node70_Data2.nd_fail = 0
        Node70_Data2.reserved1(0) = 0
        Node70_Data2.reserved2(0) = 0
        Node70_Data2.reserved3(0) = 0
    ''--------------------------------
    
    ''--------------------------------
    '' Sun added 2005-08-05
    Case 71     '查询座席状态
        Node71_Data1.timeout = 5
        Node71_Data1.usevar = 0
        Node71_Data1.log = 0
        Node71_Data1.dn_status = 0
        Node71_Data1.pos_status = 0
        Node71_Data1.dn_logic = 0
        Node71_Data1.pos_logic = 0
        Node71_Data1.querytype = 0
        Node71_Data1.conditions = 0
        Node71_Data1.reserved1(0) = 0
        Node71_Data2.agentid = 0
        Node71_Data2.nd_parent = 0
        Node71_Data2.nd_yes = 0
        Node71_Data2.nd_no = 0
        Node71_Data2.nd_fail = 0
        Node71_Data2.reserved1(0) = 0
        Node71_Data2.reserved2(0) = 0
    ''--------------------------------
    
    Case 90     '呼叫外线号码
        Node90_Data1.numbertype = 0
        Node90_Data1.dialtype = 0
        Node90_Data1.connecttype = 0
        Node90_Data1.timeout = 30
        Node90_Data1.reserved1(0) = 0
        Node90_Data1.trytimes = 3
        Node90_Data1.log = 0
        Node90_Data1.extdelay = 5
        Node90_Data1.usevar = 0
        Node90_Data1.resultvar = 0
        Node90_Data1.resultinform = 0
        Node90_Data1.explictoffhook = 0
        Node90_Data2.com_iid = 0
        Node90_Data2.nd_fail = 0
        Node90_Data2.nd_parent = 0
        Node90_Data2.nd_succ = 0
        For lv_lvp = 0 To 12
            Node90_Data2.predial(lv_lvp) = 0
        Next
        For lv_lvp = 0 To 31
            Node90_Data2.phoneno(lv_lvp) = 0
        Next
        Node90_Data2.reserved1(0) = 0
    
      '-------------------------------------------
      ' Sun added 2005-05-26
      Case 91   ' Calling Card
        Node91_Data1.timeout = 60
        Node91_Data1.talklentype = 0
        Node91_Data1.obgroup = 0
        Node91_Data1.remindminute = 0
        Node91_Data1.reserved1 = 0
        Node91_Data1.log = 0
        Node91_Data1.reserved2 = 0
        Node91_Data1.var_cardno = 0
        Node91_Data1.var_telno = 0
        Node91_Data1.var_connectlength = 0
        Node91_Data1.reserved3(0) = 0
        Node91_Data2.vox_talklen = 0
        Node91_Data2.vox_timeout = 0
        Node91_Data2.vox_noservice = 0
        Node91_Data2.reserved1(0) = 0
        Node91_Data2.com_talklength = 0
        Node91_Data2.com_billing = 0
        Node91_Data2.nd_parent = 0
        Node91_Data2.nd_child = 0
        Node91_Data2.reserved2(0) = 0
      '-------------------------------------------
    
    ''--------------------------------
    '' Sun added 2005-03-15
    Case 96     '异步通信
        Node96_Data1.timeout = 5
        Node96_Data1.seperator = Asc(";")
        Node96_Data1.extdata = 0
        Node96_Data1.extvar = 0
        Node96_Data1.reserved1 = 0
        Node96_Data1.log = 0
        
        '' Sun added 2012-11-23
        Node96_Data1.carryonasynplay = 0
                
        Node96_Data1.reserved2(0) = 0
        Node96_Data2.command = 0
        Node96_Data2.vox_wt = 0
        For lv_lvp = 0 To 9
            Node96_Data2.var_send(lv_lvp) = 0
        Next
        For lv_lvp = 0 To 9
            Node96_Data2.var_receive(lv_lvp) = 0
        Next
        Node96_Data2.fileprefix(0) = 0
        Node96_Data2.fileprefix(1) = 0
        Node96_Data2.reserved1(0) = 0
        Node96_Data2.nd_parent = 0
        Node96_Data2.nd_child = 0
        Node96_Data2.nd_timeout = 0
        Node96_Data2.reserved2(0) = 0
    ''--------------------------------
    
    Case 100    '用户DLL
        Node100_Data1.log = 0
        Node100_Data1.reserved1(0) = 0
        Node100_Data1.reserved2(0) = 0
        Node100_Data2.dll_fid = 0
        Node100_Data2.nd_child = 0
        Node100_Data2.nd_parent = 0
        Node100_Data2.reserved1(0) = 0
        Node100_Data2.reserved2(0) = 0
        
    Case 101    '用户COM
        Node101_Data1.log = 0
        Node101_Data1.reserved1(0) = 0
        Node101_Data1.reserved2(0) = 0
        Node101_Data2.com_iid = 0
        Node101_Data2.nd_child = 0
        Node101_Data2.nd_parent = 0
        Node101_Data2.reserved1(0) = 0
        Node101_Data2.reserved2(0) = 0
        Node101_Data2.reserved3(0) = 0
        
    Case 102    '记录变量
        Node102_Data1.log = 0
        Node102_Data1.var_chg = 0
        Node102_Data1.convert = 0
        Node102_Data1.param1 = 0
        Node102_Data1.param2 = 0
        Node102_Data1.reserved1(0) = 0
        Node102_Data1.reserved2(0) = 0
        Node102_Data2.com_iid = 0
        Node102_Data2.value(0) = 0
        Node102_Data2.nd_child = 0
        Node102_Data2.nd_parent = 0
        Node102_Data2.reserved1(0) = 0
        Node102_Data2.reserved2(0) = 0
        Node102_Data2.reserved3(0) = 0
        
    Case Else
        Exit Function
    End Select
    
    ' 节点数据整和
    F_NodeData f_Index, f_NodeNo
       
End Function

'用途：    整和节点数据
'作者:     Scott
'创建日期：2001/04/23
'修改日期：2001/08/29
'描述：    Auther:Scott Data:2001/08/29
Public Sub F_NodeData(ByVal f_Index As Integer, ByVal f_NodeNo As Byte)
On Error Resume Next
      
    Dim lv_Gdata1(DEF_NODE_DATA1_LEN) As Byte
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_loop
    
    Select Case Val(f_NodeNo)
    Case 0
        CopyMemory lv_Gdata1(0), Node0_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node0_Data2.key_repeat, DEF_NODE_DATA2_LEN
    Case 1
        CopyMemory lv_Gdata1(0), Node1_Data1.reserved1(0), DEF_NODE_DATA1_LEN
    Case 2
        CopyMemory lv_Gdata2(0), Node2_Data2.uservar(0), DEF_NODE_DATA2_LEN
    Case 6
        '' Sun added 2012-11-23
        CopyMemory lv_Gdata1(0), Node6_Data1.Sleep, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node6_Data2.nd_goto, DEF_NODE_DATA2_LEN
    Case 7
        CopyMemory lv_Gdata1(0), Node7_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node7_Data2.vox_userid, DEF_NODE_DATA2_LEN
    Case 8
        CopyMemory lv_Gdata1(0), Node8_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node8_Data2.vox_password, DEF_NODE_DATA2_LEN
    Case 9
        CopyMemory lv_Gdata1(0), Node9_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node9_Data2.workday, DEF_NODE_DATA2_LEN
    Case 10
        CopyMemory lv_Gdata1(0), Node10_Data1.maincalendar, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node10_Data2.daytype(0), DEF_NODE_DATA2_LEN
    
    ''--------------------------------
    '' Sun added 2004-12-30
    Case 16
        CopyMemory lv_Gdata1(0), Node16_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node16_Data2.var_value(0), DEF_NODE_DATA2_LEN
    ''--------------------------------
    
    Case 17
        CopyMemory lv_Gdata1(0), Node17_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node17_Data2.vox_play, DEF_NODE_DATA2_LEN
    Case 18
        CopyMemory lv_Gdata1(0), Node18_Data1.seperator, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node18_Data2.typeflags(0), DEF_NODE_DATA2_LEN
    Case 19
        CopyMemory lv_Gdata1(0), Node19_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node19_Data2.delaytime, DEF_NODE_DATA2_LEN
    Case 20
        CopyMemory lv_Gdata1(0), Node20_Data1.reserved1, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node20_Data2.vox_play, DEF_NODE_DATA2_LEN
    Case 21
        CopyMemory lv_Gdata2(0), Node21_Data2.vox_pred, DEF_NODE_DATA2_LEN
        CopyMemory lv_Gdata1(0), Node21_Data1.reserved1, DEF_NODE_DATA1_LEN
    Case 22
        CopyMemory lv_Gdata1(0), Node22_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node22_Data2.vox_play, DEF_NODE_DATA2_LEN
    Case 23
        CopyMemory lv_Gdata1(0), Node23_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node23_Data2.vox_play, DEF_NODE_DATA2_LEN
         
    ''--------------------------------
    '' Sun added 2004-12-30
    Case 28
        CopyMemory lv_Gdata1(0), Node28_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node28_Data2.vox_string, DEF_NODE_DATA2_LEN
    ''--------------------------------
         
    Case 40
        CopyMemory lv_Gdata1(0), Node40_Data1.rectime, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node40_Data2.vox_op, DEF_NODE_DATA2_LEN
    Case 41
        CopyMemory lv_Gdata1(0), Node41_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node41_Data2.vox_play(0), DEF_NODE_DATA2_LEN
    Case 50
        CopyMemory lv_Gdata1(0), Node50_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node50_Data2.vox_op, DEF_NODE_DATA2_LEN
    Case 51
        CopyMemory lv_Gdata1(0), Node51_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node51_Data2.vox_op, DEF_NODE_DATA2_LEN
    
    ''--------------------------------
    '' Sun added 2006-12-31
    Case 55
        CopyMemory lv_Gdata1(0), Node55_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node55_Data2.vox_op, DEF_NODE_DATA2_LEN
    ''--------------------------------
    
    Case 60
        CopyMemory lv_Gdata1(0), Node60_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node60_Data2.vox_op, DEF_NODE_DATA2_LEN
    Case 61
        CopyMemory lv_Gdata1(0), Node61_Data1.maxwait, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node61_Data2.vox_op, DEF_NODE_DATA2_LEN
    Case 62
        CopyMemory lv_Gdata1(0), Node62_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node62_Data2.vox_op, DEF_NODE_DATA2_LEN
    Case 63
        CopyMemory lv_Gdata1(0), Node63_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node63_Data2.vox_op, DEF_NODE_DATA2_LEN
    Case 69
        CopyMemory lv_Gdata1(0), Node69_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node69_Data2.vox_op, DEF_NODE_DATA2_LEN
    
    '-------------------------------------------
    ' Sun added 2005-06-27
    Case 70
        CopyMemory lv_Gdata1(0), Node70_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node70_Data2.routepointid, DEF_NODE_DATA2_LEN
    
    Case 71
        CopyMemory lv_Gdata1(0), Node71_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node71_Data2.agentid, DEF_NODE_DATA2_LEN
    '-------------------------------------------
    
    Case 90
        CopyMemory lv_Gdata1(0), Node90_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node90_Data2.predial(0), DEF_NODE_DATA2_LEN
    
    '-------------------------------------------
    ' Sun added 2005-05-26
    Case 91
        CopyMemory lv_Gdata1(0), Node91_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node91_Data2.vox_talklen, DEF_NODE_DATA2_LEN
    '-------------------------------------------
    
    ''--------------------------------
    '' Sun added 2005-03-15
    Case 96
        CopyMemory lv_Gdata1(0), Node96_Data1.timeout, DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node96_Data2.command, DEF_NODE_DATA2_LEN
    ''--------------------------------
    
    Case 100
        CopyMemory lv_Gdata1(0), Node100_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node100_Data2.dll_fid, DEF_NODE_DATA2_LEN
    Case 101
        CopyMemory lv_Gdata1(0), Node101_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node101_Data2.reserved1(0), DEF_NODE_DATA2_LEN
    Case 102
        CopyMemory lv_Gdata1(0), Node102_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node102_Data2.reserved1(0), DEF_NODE_DATA2_LEN
    Case 255
        CopyMemory lv_Gdata1(0), Node255_Data1.reserved1(0), DEF_NODE_DATA1_LEN
        CopyMemory lv_Gdata2(0), Node255_Data2.StartNode, DEF_NODE_DATA2_LEN
    End Select
      
    If f_NodeNo <> 1 Then
        gGetUserVar.StrN_data2 = ""
        For lv_loop = 0 To 63
            gGetUserVar.StrN_data2 = gGetUserVar.StrN_data2 & ChrB(lv_Gdata2(lv_loop))
        Next
        gCallFlow.Node(f_Index).Data2 = gGetUserVar.StrN_data2
    End If
      
    gGetUserVar.StrN_data1 = ""
    For lv_loop = 0 To 12
        gGetUserVar.StrN_data1 = gGetUserVar.StrN_data1 & ChrB(lv_Gdata1(lv_loop))
    Next
    gCallFlow.Node(f_Index).Data1 = gGetUserVar.StrN_data1
           
End Sub

Public Sub F_GetWorkPageScale(ByVal f_Page As Integer, f_Left As Integer, f_Top As Integer, f_Right As Integer, f_Bottom As Integer)
On Error Resume Next

    Dim lv_Index As Integer
    
    '' Init Var
    f_Left = 32767
    f_Top = 32767
    f_Right = 0
    f_Bottom = 0
    
    For lv_Index = 1 To gCallFlow.NewNodeID
        If (gCallFlow.Node(lv_Index).InPage = f_Page Or f_Page <= 0) And gCallFlow.Node(lv_Index).NodeID > 0 Then
            If gCallFlow.Node(lv_Index).Left < f_Left Then f_Left = gCallFlow.Node(lv_Index).Left
            If gCallFlow.Node(lv_Index).Top < f_Top Then f_Top = gCallFlow.Node(lv_Index).Top
            If gCallFlow.Node(lv_Index).Width + gCallFlow.Node(lv_Index).Left > f_Right Then f_Right = gCallFlow.Node(lv_Index).Width + gCallFlow.Node(lv_Index).Left
            If gCallFlow.Node(lv_Index).Height + gCallFlow.Node(lv_Index).Top > f_Bottom Then f_Bottom = gCallFlow.Node(lv_Index).Height + gCallFlow.Node(lv_Index).Top
        End If
    Next
    
End Sub

Public Function F_GetPrintPageScale(f_Width As Long, f_Height As Long) As Boolean
    Dim lv_PageHeight As Long
    Dim lv_PageWidth As Long

    F_GetPrintPageScale = True
    Select Case Printer.PaperSize
    Case vbPRPSA3           '' A3, 297 x 420 毫米
        lv_PageHeight = CLng(420 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
        lv_PageWidth = CLng(297 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
    Case vbPRPSA4           '' A4, 210 x 297 毫米
        lv_PageHeight = CLng(297 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
        lv_PageWidth = CLng(210 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
    Case vbPRPSA5           '' A5, 148 x 210 毫米
        lv_PageHeight = CLng(210 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
        lv_PageWidth = CLng(148 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
    Case vbPRPSB4           '' B4, 250 x 354 毫米
        lv_PageHeight = CLng(354 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
        lv_PageWidth = CLng(250 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
    Case vbPRPSB5           '' B5, 182 x 257 毫米
        lv_PageHeight = CLng(257 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
        lv_PageWidth = CLng(182 * CLng(Def_TWIPS_PER_CM) / 10 + 0.5)
    Case Else
        F_GetPrintPageScale = False
    End Select
    
    f_Width = lv_PageWidth
    f_Height = lv_PageHeight
    
    '' Sun added 2007-10-20
    gSystem.intPageWidth = lv_PageWidth
    gSystem.intPageHeight = lv_PageHeight

End Function

Public Sub GotoAnotherPage(ByVal f_NewPage As Integer)
On Error Resume Next

    If f_NewPage > gCallFlow.PageCount Then Exit Sub
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    gCallFlow.CurrentPage = f_NewPage
    
    '' Sun added 2002-03-30
    Call gCallFlow.SetAllNodeCaptionVisible(gblnShowNodeCaption)
    
    '' Sun added 2008-01-18
    Call gCallFlow.SetAllNodeTagVisible(gblnShowNodeTag)
    
    frmMain.StatusBar.Panels("Page").Text = LoadNationalResString(1554) & Trim(Str(gCallFlow.CurrentPage)) & LoadNationalResString(1132) & Trim(Str(gCallFlow.PageCount)) & LoadNationalResString(1133)

    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
            
End Sub

'Mike change the mnuEdit index(12,13) to (13,14) @ 08-1-29
Public Sub F_SetUnDoMenuState(ByVal f_State)

    Select Case f_State
    Case 0                  '' 初始化
        frmMain.mnuEdit(13).Enabled = False
        frmMain.mnuEdit(13).Caption = LoadNationalResString(1480)
        frmMain.mnuEdit(14).Enabled = False
        frmMain.mnuEdit(14).Caption = LoadNationalResString(1481)
        gClipBoard.ClipItem = 0
        gClipBoard.ReDoTimes = 0

    Case 1                  '' 新的粘贴板操作 或 ReDo到底
        frmMain.mnuEdit(13).Enabled = True
        frmMain.mnuEdit(13).Caption = LoadNationalResString(1482)
        frmMain.mnuEdit(14).Enabled = False
        frmMain.mnuEdit(14).Caption = LoadNationalResString(1481)
        gClipBoard.ReDoTimes = 0
        gClipBoard.ClipItem = gClipBoard.ClipItem + 1
        
    Case 2                  '' UnDo到底
        frmMain.mnuEdit(13).Enabled = False
        frmMain.mnuEdit(13).Caption = LoadNationalResString(1480)
        frmMain.mnuEdit(14).Enabled = True
        frmMain.mnuEdit(14).Caption = LoadNationalResString(1483)
        gClipBoard.ReDoTimes = gClipBoard.ReDoTimes + 1
        gClipBoard.ClipItem = 0
        
    Case 3                  '' UnDo前一次
        frmMain.mnuEdit(13).Enabled = True
        frmMain.mnuEdit(13).Caption = LoadNationalResString(1482)
        frmMain.mnuEdit(14).Enabled = True
        frmMain.mnuEdit(14).Caption = LoadNationalResString(1483)
        gClipBoard.ReDoTimes = gClipBoard.ReDoTimes + 1
        gClipBoard.ClipItem = gClipBoard.ClipItem - 1
    
    Case 4                  '' ReDo前一次
        frmMain.mnuEdit(13).Enabled = True
        frmMain.mnuEdit(13).Caption = LoadNationalResString(1482)
        frmMain.mnuEdit(14).Enabled = True
        frmMain.mnuEdit(14).Caption = LoadNationalResString(1483)
        gClipBoard.ReDoTimes = gClipBoard.ReDoTimes - 1
        gClipBoard.ClipItem = gClipBoard.ClipItem + 1
    
    End Select
    
    frmMain.tbToolBar.Buttons("撤销").Enabled = frmMain.mnuEdit(13).Enabled
    frmMain.tbToolBar.Buttons("重复").Enabled = frmMain.mnuEdit(14).Enabled
        
End Sub

'' Sun added 2002-09-10
''' Add Resource Description to ToolTip
Public Sub F_RefreshVoxBoxToolTip(ctlEditBox As TextBox)
On Error Resume Next

    Dim lv_FN As String
    Dim f_RID As Integer
    
    ctlEditBox.ToolTipText = ""
    f_RID = Val(ctlEditBox)
    If f_RID > 0 Then
        If gCallFlow.GetResourceDescriptionWithID(f_RID, lv_FN) Then
            ctlEditBox.ToolTipText = Trim(lv_FN)
        End If
    End If

End Sub

'' Sun added 2006-01-06
''' Set Root Node ID
Public Sub F_SetResourceID(ByVal f_nRID As Integer)

    F_ExplainNodeData 1
    If Node0_Data2.ResourceProject <> f_nRID Then
        Node0_Data2.ResourceProject = f_nRID
        F_NodeData 1, 0
        gCallFlow.UpdateAnotherIVRRecord 1
    End If
    
End Sub

'' Sun added 2002-12-04
''' Set Root Node ID
Public Sub F_SetRootNode(f_nNodeIndex As Integer)

    Dim lv_nNodeID As Integer
    
    lv_nNodeID = gCallFlow.Node(f_nNodeIndex).NodeID
    If lv_nNodeID < 256 Then
        Exit Sub
    End If
    
    F_ExplainNodeData 1
    If Node0_Data2.nd_root <> lv_nNodeID Then
        Node0_Data2.nd_root = lv_nNodeID
        F_NodeData 1, 0
        gCallFlow.UpdateAnotherIVRRecord 1
    End If
    
    F_SwitchRootNodeDisplay lv_nNodeID
    
End Sub

'' Sun added 2002-12-04
'' Set Root Node ID
Public Sub F_SwitchRootNodeDisplay(f_nRoot As Integer)
    Dim lv_Index As Integer
    
    If f_nRoot < 256 Then
        Exit Sub
    End If
    
    ' Close Old Root
    If gCallFlow.RootNodedID >= 256 Then
        lv_Index = gCallFlow.SearchNodeIndexWithID(gCallFlow.RootNodedID)
        If lv_Index > 0 Then
            gCallFlow.Node(lv_Index).IsRootNode = False
        End If
    End If
    
    ' Open New Root
    'f_nRoot
    lv_Index = gCallFlow.SearchNodeIndexWithID(f_nRoot)
    If lv_Index > 0 Then
        gCallFlow.Node(lv_Index).IsRootNode = True
    End If
    
    frmMain.StatusBar.Panels("Root").Text = LoadNationalResString(1072) & Trim(Str(f_nRoot))
    gCallFlow.RootNodedID = f_nRoot
    
End Sub

'' Sun added 2004-12-30
'' 根据 Language 返回字符串资源
Public Function LoadNationalResString(ByVal f_nResID As Long) As String
On Error Resume Next

Dim lv_nNationResID As Long
#If Language = 0 Then
    lv_nNationResID = f_nResID
#ElseIf Language = 1 Then
    lv_nNationResID = f_nResID + 10000
#End If

    If lv_nNationResID > 0 Then
        LoadNationalResString = LoadResString(lv_nNationResID)
    Else
        LoadNationalResString = ""
    End If
    
On Error GoTo 0
End Function

Public Sub LoadResStrings(frm As Form) ' 为form 载入字符
On Error Resume Next
    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer
    Dim nListCount As Integer
    '设置窗体的 caption 属性
    If frm.Tag <> "" Then
            frm.Caption = LoadNationalResString(CInt(frm.Tag))
    End If
Dim lv_minttemp As Integer
    '设置字体 '在这设置字体将改变所有控件的字体
'    Set fnt = frm.Font
'    fnt.Name = LoadNationalResString(20)
'    fnt.Size = CInt(LoadNationalResString(21))
    
    '设置控件的标题，对菜单项使用 caption 属性并对所有其他控件使用 Tag 属性
    For Each ctl In frm.Controls
'        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then 'Label
            If ctl.Tag <> "" Then
                    ctl.Caption = LoadNationalResString(CInt(ctl.Tag))
            End If
        ElseIf sCtlType = "Menu" Then '菜单
                If ctl.Caption <> "" Then
                        ctl.Caption = LoadNationalResString(CInt(ctl.Caption))
                End If
        ElseIf sCtlType = "TreeView" Then 'Tree
            For lv_minttemp = 1 To ctl.nodes.Count
                If ctl.nodes.Item(lv_minttemp).Text <> "" Then
                  ctl.nodes.Item(lv_minttemp).Text = LoadNationalResString(ctl.nodes.Item(lv_minttemp).Text)
                End If
            Next
        ElseIf sCtlType = "StatusBar" Then '状态栏
                   For lv_minttemp = 1 To ctl.Panels.Count
                  If ctl.Panels.Item(lv_minttemp).Text <> "" Then
                        ctl.Panels.Item(lv_minttemp).Text = LoadNationalResString(CInt(ctl.Panels.Item(lv_minttemp).Text))
                    End If
                    If ctl.Panels.Item(lv_minttemp).ToolTipText <> "" Then
                        ctl.Panels.Item(lv_minttemp).ToolTipText = LoadNationalResString(CInt(ctl.Panels.Item(lv_minttemp).ToolTipText))
                    End If
                   Next
        ElseIf sCtlType = "DataGrid" Then 'DataGrid
                    lv_minttemp = 0
                    For lv_minttemp = 0 To ctl.Columns.Count
                        If ctl.Columns.Item(lv_minttemp).Caption <> "" Then
                                ctl.Columns.Item(lv_minttemp).Caption = LoadNationalResString(CInt(ctl.Columns.Item(lv_minttemp).Caption))
                        End If
                    Next
        ElseIf sCtlType = "SSTab" Then 'SSTab
              lv_minttemp = 0
              For lv_minttemp = 0 To ctl.Tabs - 1
                  ctl.Tab = lv_minttemp
                  If ctl.Caption <> "" Then
                        ctl.Caption = LoadNationalResString(CInt(ctl.Caption))
                  End If
              Next
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                If obj.Tag <> "" Then
                
                    obj.Caption = LoadNationalResString(CInt(obj.Tag))
                End If
                If obj.ToolTipText <> "" Then
                    obj.ToolTipText = LoadNationalResString(CInt(obj.ToolTipText))
                End If
            Next
        ElseIf sCtlType = "Toolbar" Then
'Temp 2004-6-9
            For Each obj In ctl.Buttons
'                obj.Caption = LoadNationalResString(CInt(obj.Tag))
                If obj.ToolTipText <> "" Then
                        obj.ToolTipText = LoadNationalResString(CInt(obj.ToolTipText))
                End If
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                If obj.Tag <> "" Then
                    obj.Text = LoadNationalResString(CInt(obj.Tag))
                End If
            Next
        ElseIf sCtlType = "VerticalMenu" Then
             lv_minttemp = 0
             For lv_minttemp = 1 To ctl.MenusMax
                    ctl.MenuCur = lv_minttemp
                        If ctl.MenuCaption <> "" Then
                            ctl.MenuCaption = LoadNationalResString(ctl.MenuCaption)
                        End If
                            For nListCount = 1 To ctl.MenuItemsMax
                                    ctl.MenuItemCur = nListCount
                                        If ctl.MenuItemTag <> "" Then
                                            ctl.MenuItemCaption = LoadNationalResString(ctl.MenuItemTag)
                                        End If
                            Next
             Next
        ElseIf sCtlType = "fpSpread" Or sCtlType = "vaSpread" Then
                ctl.Row = 0
                lv_minttemp = 0
                For lv_minttemp = 1 To ctl.MaxCols
                        ctl.Col = lv_minttemp
                        If ctl.Text <> "" Then
                                ctl.Text = LoadNationalResString(ctl.Text)
                        End If
                Next
        Else
            nVal = 0
            If ctl.Tag <> "" Then
                nVal = Val(ctl.Tag)
                If nVal > 0 Then ctl.Caption = LoadNationalResString(nVal)
            End If
            nVal = 0
            If ctl.ToolTipText <> "" Then
                nVal = Val(ctl.ToolTipText)
                If nVal > 0 Then ctl.ToolTipText = LoadNationalResString(nVal)
            End If
        End If
    Next
End Sub

'Jeremy Add 2004-07-08
Public Sub DateStrToAsc(ByRef RsSet As ADODB.Recordset, strData1 As String, strData2 As String)
Dim nCyc As Integer
Dim sKey As String

Dim nMaxLen As Integer

'' Data Part 1 - 13 Bytes
nCyc = 1
sKey = "N_D1_"
If LenB(strData1) > DEF_NODE_DATA1_LEN Then
    nMaxLen = DEF_NODE_DATA1_LEN
Else
    nMaxLen = LenB(strData1)
End If
For nCyc = 1 To nMaxLen
    RsSet(sKey & CStr(nCyc)) = AscB(MidB(strData1, nCyc, 1))
Next

'' Sun added 2006-06-20
For nCyc = nMaxLen + 1 To DEF_NODE_DATA1_LEN
    RsSet(sKey & CStr(nCyc)) = 0
Next

'' Data Part 2 - 64 Bytes
nCyc = 1
sKey = "N_D2_"
If LenB(strData2) > DEF_NODE_DATA2_LEN Then
    nMaxLen = DEF_NODE_DATA2_LEN
Else
    nMaxLen = LenB(strData2)
End If
For nCyc = 1 To nMaxLen
    RsSet(sKey & CStr(nCyc)) = AscB(MidB(strData2, nCyc, 1))
Next

'' Sun added 2006-06-20
For nCyc = nMaxLen + 1 To DEF_NODE_DATA2_LEN
    RsSet(sKey & CStr(nCyc)) = 0
Next

End Sub

Public Function CharToStr(RsSet As ADODB.Recordset, nCount As Integer) As String
Dim lv_strData As String
Dim sKey As String
Dim nCyc As Integer
lv_strData = ""
        If nCount = DEF_NODE_DATA1_LEN Then
        sKey = "N_D1_"
                For nCyc = 1 To nCount
                        lv_strData = lv_strData & ChrB(RsSet.Fields(sKey & CStr(nCyc)))

                Next
        
        End If
        
        If nCount = DEF_NODE_DATA2_LEN Then
        sKey = "N_D2_"
                For nCyc = 1 To nCount
                    
                        lv_strData = lv_strData & ChrB(RsSet.Fields(sKey & CStr(nCyc)))
                    
                    
                Next
        End If
        CharToStr = lv_strData
        
End Function

' Sun added 2004-12-30
'
Public Sub RefreshVariablesList(objCtrl As Control)
On Error Resume Next

Dim lv_loop As Integer
Dim lv_strName As String
Dim lv_bytType As Byte
Dim lv_bytLen As Byte
Dim lv_nVarCount As Integer

Dim itmX As ListItem                    ' ListItem 变量

    With objCtrl
        .Clear
        
        ' 不使用变量
        .AddItem LoadNationalResString(1181)
        .ItemData(.ListCount - 1) = 0
        
        ' 变量列表
        lv_nVarCount = gCallFlow.GetUserVarCount
        For lv_loop = 1 To lv_nVarCount
            
            lv_strName = ""
            lv_bytType = 0
            lv_bytLen = 0
            Call gCallFlow.GetUserVarDefination(lv_loop, lv_strName, lv_bytType, lv_bytLen)
        
            objCtrl.AddItem lv_strName
            objCtrl.ItemData(objCtrl.ListCount - 1) = lv_loop
                        
        Next
    End With
    
On Error GoTo 0
End Sub

'end

'Public Functions
Public Function F_DealColor(nColor As OLE_COLOR) As OLE_COLOR
    If nColor >= 0 Then F_DealColor = nColor Else F_DealColor = GetSysColor(nColor + 2147483648#)
End Function

Public Sub F_CustomColor(nImportColor As OLE_COLOR, oObject As Object)
Dim lv_Color As CHOOSECOLOR

    With lv_Color
        .lStructSize = Len(lv_Color)
        .flags = CC_ANYCOLOR + CC_FULLOPEN + CC_RGBINIT
        .hInstance = App.hInstance
        .hwndOwner = oObject.hWnd
        .rgbResult = F_DealColor(nImportColor)
        .lpCustColors = 0
    End With
    If CHOOSECOLOR(lv_Color) Then
        oObject.BackColor = lv_Color.rgbResult
    End If
    
End Sub

Public Function F_PlayVoxFile(f_RID As Long, Optional ByVal f_LID As Integer = -1) As Boolean
On Error GoTo BackDoor
    
    Dim lv_FNO As Long
    Dim lv_FN As String
    Dim bVox As Boolean

    F_PlayVoxFile = False
    
    '' Get Resource Path
    If Not gCallFlow.GetResourceNameWithID(f_RID, f_LID, lv_FN) Then
        Exit Function
    End If
    
    bVox = False
    If InStr(1, LCase(lv_FN), ".vox") > 0 Then
        bVox = True
    End If
    Debug.Print lv_FN, bVox
    PlaySoundFile gSystem.strPath_SysVox & lv_FN, False, bVox
        
    F_PlayVoxFile = True
    
BackDoor:
    Debug.Print Err.Description
    On Error GoTo 0
End Function

Public Sub StopSound()
    sndPlaySound 0, 1
    frmMain.VOXPlayer.StopPlay
End Sub

Public Sub PlaySoundFile(strFile As String, bSwitch As Boolean, bVox As Boolean)
    If bVox = True Then
        frmMain.VOXPlayer.StopPlay
        frmMain.VOXPlayer.PlayVOXFile strFile
    Else
        sndPlaySound 0, 1
        sndPlaySound strFile, 1
    End If

End Sub

Public Sub F_PlayWavFile(f_FN As String, f_Switch As Boolean)
On Error Resume Next

    If f_Switch Then
        sndPlaySound f_FN, 1
    Else
        sndPlaySound 0, 1
    End If
     
On Error GoTo 0
End Sub

' 打开文件对话框
'
Public Function F_OpenFileDialog(lOwnerhWnd As Long, ByVal blnSaveDlg As Boolean, ByVal strTitile As String, ByVal strFilter As String, ByVal strExtName) As String
    Dim lv_FileName As String
    Dim m_ofn As OPENFILENAME
    
    F_OpenFileDialog = ""
    With m_ofn
        .lStructSize = Len(m_ofn)
        .hInstance = App.hInstance
        .flags = OFN_EXTENSIONDIFFERENT + OFN_FILEMUSTEXIST
        .hwndOwner = lOwnerhWnd
        .lpstrTitle = strTitile
        .lpstrFilter = strFilter
        .lpstrDefExt = strExtName
        .lpstrFile = Space(249) & "*." & strExtName
        .nMaxFile = 260
    
        If Not blnSaveDlg Then
            lv_FileName = IIf(GetOpenFileName(m_ofn), .lpstrFile, "")
        Else
            lv_FileName = IIf(GetSaveFileName(m_ofn), .lpstrFile, "")
        End If
            
        '' Cut off illegal characters
        F_OpenFileDialog = EraseInvalidCharacters(lv_FileName)
        
    End With

End Function

' 显示并确认资源文件
'
Public Function F_ConfirmResourceFileInfo(f_strFileName As String) As Byte
On Error GoTo BackDoor

    Dim FileNumber
    Dim lv_StrTemp As String
    Dim lv_strLine As String
    Dim lv_FileCaption As String
    Dim strArrCaptions() As String
    Dim lv_bytPID As Byte
    Dim lv_sSQL As String
    Dim lv_CN As ADODB.Connection        '' 连接
    Dim lv_RS As ADODB.Recordset
    
    ' 缺省返回值
    F_ConfirmResourceFileInfo = 0
    
    '*****************************************************
    '* Open File for Read
    '*****************************************************
    If Not CheckExistFile(f_strFileName) Then
        Call MsgBox(LoadNationalResString(1669) & " ‘" & f_strFileName & "' " & LoadNationalResString(1125), vbOK + vbExclamation, App.Title)
        Exit Function
    End If
    Debug.Print "Open File for Read: " & f_strFileName
    FileNumber = FreeFile   ' 取得未使用的文件号

    Open f_strFileName For Input Access Read Lock Write As #FileNumber
    If Err Then
        Err.Clear
        MsgBox Err.Description
        MsgBox LoadNationalResString(1110) & f_strFileName & LoadNationalResString(1111), vbCritical + vbOKOnly, App.Title
        Exit Function
    Else
        '*****************************************************
        '* place loading flow information codes here         *
        '*****************************************************
        ''' Read File
        Do While Not EOF(FileNumber)
            Input #FileNumber, lv_StrTemp
            If Err Then
                Err.Clear
                MsgBox LoadNationalResString(1126) & f_strFileName, vbCritical + vbOKOnly, App.Title
                GoTo BackDoor
            Else
                lv_strLine = Trim(lv_StrTemp)
                If lv_strLine <> "" Then
                    If Val(Left(lv_strLine, 2)) = -1 Then
                        strArrCaptions = Split(lv_strLine, vbTab)
                        If UBound(strArrCaptions) >= 9 Then
                            lv_FileCaption = LoadNationalResString(1703) & LoadNationalResString(1669) & " ‘" & f_strFileName & "' " & LoadNationalResString(1128) & " :" & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1114) & " :" & Trim(strArrCaptions(1)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1116) & " :" & Trim(strArrCaptions(3)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1117) & " :" & Trim(strArrCaptions(4)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1105) & " :" & Trim(strArrCaptions(6)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1118) & " :" & Trim(strArrCaptions(7)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1119) & " :" & Trim(strArrCaptions(4)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1106) & " :" & Trim(strArrCaptions(8)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1120) & " :" & Trim(strArrCaptions(9)) & vbCrLf
                            lv_FileCaption = lv_FileCaption & LoadNationalResString(1129)
                        
                            If MsgBox(lv_FileCaption, vbYesNo + vbQuestion, App.Title) = vbNo Then
                                GoTo BackDoor
                            Else
                                Exit Do
                            End If
                            
                        Else
                            Message "E068"
                            GoTo BackDoor
                            
                        End If
                    End If
                End If
            End If
        Loop
    End If

    '判断流程是否在数据库里存在
    lv_bytPID = CByte(Val(strArrCaptions(1)))

    Set lv_CN = New ADODB.Connection
    lv_CN.ConnectionString = gSystem.strConString
    lv_CN.CursorLocation = adUseClient
    lv_CN.Open

    '连接数据库
    lv_sSQL = "Select * From tbRefIVR Where P_Type='R' AND P_ID = " & Str(lv_bytPID)
    Set lv_RS = New ADODB.Recordset
    lv_RS.CursorLocation = adUseClient
    lv_RS.LockType = adLockReadOnly
    lv_RS.CursorType = adOpenForwardOnly
    lv_RS.Open lv_sSQL, lv_CN
    If lv_RS.RecordCount > 0 Then
        lv_RS.Close
        If Message("Q012") = vbNo Then
            GoTo BackDoor
        Else
            '' 覆盖
            lv_sSQL = "DELETE FROM tbRefIVR WHERE P_Type='R' AND P_ID= " & Str(lv_bytPID)
            lv_CN.Execute lv_sSQL, -1, adCmdText
        End If
    Else
        lv_RS.Close
    End If
    
    '' 资源描述
    lv_sSQL = "Select * From tbRefIVR Where P_Type='R' AND P_ID= " & Str(lv_bytPID)
    lv_RS.CursorLocation = adUseClient
    lv_RS.LockType = adLockPessimistic
    lv_RS.CursorType = adOpenStatic
    lv_RS.Open lv_sSQL, lv_CN
    
    lv_RS.AddNew
    lv_RS("P_ID") = lv_bytPID
    lv_RS("P_Type") = "R"
    lv_RS("P_Name") = Trim(strArrCaptions(3))
    lv_RS("P_Description") = Trim(strArrCaptions(4))
    lv_RS("P_version") = Trim(strArrCaptions(5))
    lv_RS("P_Auther") = Trim(strArrCaptions(6))
    lv_RS("P_User") = Trim(strArrCaptions(7))
    lv_RS("P_CreateTime") = Trim(strArrCaptions(8))
    lv_RS("P_ModifyTime") = Trim(strArrCaptions(9))
    lv_RS.Update
    lv_RS.Close
    
    '' 资源明细
    lv_sSQL = "DELETE FROM tbResource WHERE P_ID= " & Str(lv_bytPID)
    lv_CN.Execute lv_sSQL, -1, adCmdText

    lv_sSQL = "Select * From tbResource Where P_ID = " & Str(lv_bytPID)
    lv_RS.CursorLocation = adUseClient
    lv_RS.LockType = adLockPessimistic
    lv_RS.CursorType = adOpenStatic
    lv_RS.Open lv_sSQL, lv_CN

    Do While Not EOF(FileNumber)
        Input #FileNumber, lv_StrTemp
        If Err Then
            Err.Clear
            MsgBox LoadNationalResString(1126) & f_strFileName, vbCritical + vbOKOnly, App.Title
            lv_RS.Close
            GoTo BackDoor
        Else
            lv_strLine = Trim(lv_StrTemp)
            
            If lv_strLine <> "" Then

'Michael added @ 2008-9-9
#If Language = 1 Then
    lv_strLine = Replace(lv_strLine, "，", ",", , , vbTextCompare)
#End If

                If Val(Left(lv_strLine, 2)) = -2 Then
                    strArrCaptions = Split(lv_strLine, vbTab)
                    If UBound(strArrCaptions) >= 9 Then
                        lv_RS.AddNew
                        lv_RS("P_ID") = lv_bytPID
                        lv_RS("L_ID") = CByte(Val(strArrCaptions(2)))
                        lv_RS("R_ID") = Val(strArrCaptions(3))
                        lv_RS("R_Type") = Trim(strArrCaptions(4))
                        lv_RS("R_Description") = Trim(strArrCaptions(5))
                        lv_RS("R_Path") = Trim(strArrCaptions(6))
                        lv_RS("R_Note") = Trim(strArrCaptions(7))
                        lv_RS("CreateTime") = Trim(strArrCaptions(8))
                        lv_RS("ModifyTime") = Trim(strArrCaptions(9))
                        lv_RS.Update
                    End If
                End If
            End If
        End If
    Loop
    
    lv_RS.Close
    F_ConfirmResourceFileInfo = lv_bytPID

BackDoor:
    
    ' Close File
    Close #FileNumber

    If Err Then
        Debug.Print Err.Description
    End If
    
End Function

' 填充按键列表
' objCtrl 需要填充的List控件
' nType 填充内容分类
'   1 - 中断按键
'   2 - 全程转移按键
Public Sub F_FillPhoneKeyList(objCtrl As Control, ByVal nType As Integer)
    Dim ilp As Integer
    
    If objCtrl Is Nothing Then
        Exit Sub
    End If
    
    Select Case nType
    Case 1
        '输入终止符
        With objCtrl
            .Clear
            For ilp = 0 To 9
                .AddItem Trim(Str(ilp)) & LoadNationalResString(1172)
                .ItemData(.ListCount - 1) = 48 + ilp
            Next
            .AddItem LoadNationalResString(1173)
            .ItemData(.ListCount - 1) = 65
            .AddItem LoadNationalResString(1174)
            .ItemData(.ListCount - 1) = 78
            .AddItem LoadNationalResString(1175)
            .ItemData(.ListCount - 1) = 70
            .AddItem LoadNationalResString(1176)
            .ItemData(.ListCount - 1) = 42
            .AddItem LoadNationalResString(1177)
            .ItemData(.ListCount - 1) = 35
        End With
        
    Case 2
        '全程转移按键
        With objCtrl
            .Clear
            .AddItem "NULL"
            .ItemData(.ListCount - 1) = 0
            .AddItem "0"
            .ItemData(.ListCount - 1) = Asc("0")
            .AddItem "1"
            .ItemData(.ListCount - 1) = Asc("1")
            .AddItem "2"
            .ItemData(.ListCount - 1) = Asc("2")
            .AddItem "3"
            .ItemData(.ListCount - 1) = Asc("3")
            .AddItem "4"
            .ItemData(.ListCount - 1) = Asc("4")
            .AddItem "5"
            .ItemData(.ListCount - 1) = Asc("5")
            .AddItem "6"
            .ItemData(.ListCount - 1) = Asc("6")
            .AddItem "7"
            .ItemData(.ListCount - 1) = Asc("7")
            .AddItem "8"
            .ItemData(.ListCount - 1) = Asc("8")
            .AddItem "9"
            .ItemData(.ListCount - 1) = Asc("9")
            .AddItem "*"
            .ItemData(.ListCount - 1) = Asc("*")
            .AddItem "#"
            .ItemData(.ListCount - 1) = Asc("#")
        End With
    End Select

End Sub

 Function gPrintListView(ByRef pobjListView As ListView, pstrHeading As String, Prn As Object) As Boolean
     '--------------------------------------------------------------------------
     ' Name : gPrintListView
     ' Description : Print List View Control
     ' Parameters : Listview control, Printed page heading, Printer Object
     ' Returns : N/A
     ' Called From : Anywhere
     ' Date : 07/19/2007
     ' Notes :
     '--------------------------------------------------------------------------
     Dim objCol As ColumnHeader
     Dim objLI As ListItem
     Dim objILS As ImageList
     Dim objPic As Picture
    
     Dim dblXScale As Double
     Dim dblYScale As Double
     Dim sngFontSize As Single
     Dim lngX As Long
     Dim lngY As Long
     Dim lngX1 As Long
     Dim lngY1 As Long
     Dim lngX2 As Long
     Dim lngRows As Long
     Dim lngLeft As Long
     Dim lngPageNo As Long
     Dim lngEOP As Long
     Dim lngEnd As Long
     Dim lngWidth As Long
     Dim intCols As Integer
     Dim lngTop As Long
     Dim intOffset As Integer
     Dim px As Integer
     Dim py As Integer
     Dim intRowHeight As Integer
     Dim strText As String
     Dim strTextTrun As String
    
     '--------------------------------------------------------------------------
     'Establish print & screen metrics
     '--------------------------------------------------------------------------
     On Error GoTo Error_Handler
    
     Screen.MousePointer = vbHourglass
     For Each objCol In pobjListView.ColumnHeaders
        lngX = lngX + objCol.Width
     Next
    
     Set objILS = pobjListView.SmallIcons
    
     dblXScale = (Prn.Width * 0.9) / lngX
     dblYScale = Prn.Height / pobjListView.Height
     lngLeft = (Prn.Width - (Prn.Width * 0.95)) / 2
     sngFontSize = Prn.Font.Size
    
     If pstrHeading <> "" Then
        Prn.Font.Size = 16
        Prn.CurrentX = (Prn.Width / 2) - (Prn.TextWidth(pstrHeading) / 2)
        'Prn.Font.Underline = True
        Prn.Font.Bold = True
        Prn.Print pstrHeading
        Prn.Font.Underline = False
        Prn.Font.Size = sngFontSize
        lngTop = Prn.CurrentY + Prn.CurrentY
     End If
    
     intRowHeight = (Screen.TwipsPerPixelY * 17)
     lngEOP = Prn.Height - (intRowHeight * 3)
     lngX = lngLeft
     lngY = lngTop
     lngY1 = lngTop + (Screen.TwipsPerPixelY * 17)
     Prn.CurrentY = lngY
     Prn.Font.Bold = True
     Prn.DrawMode = vbCopyPen
     px = Screen.TwipsPerPixelX
     py = Screen.TwipsPerPixelY
    
     '--------------------------------------------------------------------------
     'Print column headers with slight 3D effect
     '--------------------------------------------------------------------------
     For Each objCol In pobjListView.ColumnHeaders
         lngX1 = lngX + (objCol.Width * dblXScale)
         Prn.Line (lngX, lngY)-(lngX1, lngY1), vbButtonShadow, BF
         Prn.Line (lngX, lngY)-(lngX1 - px, lngY1), RGB(245, 245, 245), BF
         Prn.Line (lngX + px, lngY + py)-(lngX1, lngY1), vbButtonShadow, BF
         Prn.Line (lngX + px, lngY + py)-(lngX1 - px, lngY1 - py), vbButtonFace, BF
         Prn.CurrentY = lngY + ((intRowHeight - Prn.TextHeight(objCol.Text)) / 2) + py
         Select Case objCol.Alignment
             Case ListColumnAlignmentConstants.lvwColumnCenter
                Prn.CurrentX = lngX + (((objCol.Width * dblXScale) - Prn.TextWidth(objCol.Text)) / 2)
             Case ListColumnAlignmentConstants.lvwColumnLeft
                Prn.CurrentX = lngX + (px * 5)
             Case ListColumnAlignmentConstants.lvwColumnRight
                Prn.CurrentX = lngX + ((objCol.Width * dblXScale) - Prn.TextWidth(objCol.Text)) - (px * 5)
         End Select
        
         Prn.Print objCol.Text
         lngX = lngX1
     Next
    
     lngEnd = lngX1 + px
     Prn.Font.Bold = False
    
     '--------------------------------------------------------------------------
     'Print list item data
     '--------------------------------------------------------------------------
     For Each objLI In pobjListView.ListItems
         If lngY1 > lngEOP - intRowHeight - intRowHeight Then
         '------------------------------------------------------------------
         'Print page number
         '------------------------------------------------------------------
            lngPageNo = lngPageNo + 1
            Prn.CurrentX = (Prn.Width / 2) - (Prn.TextWidth("第 " & lngPageNo & " 页") / 2)
            Prn.CurrentY = lngEOP - intRowHeight
            Prn.Print "第 " & lngPageNo & " 页" '"Page " & lngPageNo
            Prn.NewPage
            Prn.CurrentY = lngTop
            lngY = lngTop
         Else
            lngY = lngY + intRowHeight
         End If
        
         lngX = lngLeft
         lngY1 = lngY + intRowHeight
        
        For Each objCol In pobjListView.ColumnHeaders
         '------------------------------------------------------------------
         'Print the icon if on col 1
         '------------------------------------------------------------------
            If objCol.Index > 1 Then
               strText = objLI.SubItems(objCol.Index - 1)
               intOffset = 0
            Else
               strText = objLI.Text
               If IsEmpty(objLI.SmallIcon) Then
                   intOffset = 0
               Else
                Set objPic = objILS.Overlay(objLI.SmallIcon, objLI.SmallIcon)
                Prn.PaintPicture objPic, lngX + px, lngY + (py / 2), 16 * px, 16 * py, , , , , vbSrcCopy
                intOffset = px * 16
               End If
            End If
           
            '------------------------------------------------------------------
            'Make sure text fits
            '------------------------------------------------------------------
            lngWidth = (objCol.Width * dblXScale)
            lngX1 = lngX + lngWidth
            strTextTrun = strText
           
            Do Until Prn.TextWidth(strTextTrun) < lngWidth - (px * 5) - intOffset Or strText = ""
               strText = Left$(strText, Len(strText) - 1)
               strTextTrun = strText & "..."
            Loop
           
            Prn.Line (lngX, lngY)-(lngX1, lngY1), 1, B
           
            Prn.CurrentY = lngY + ((intRowHeight - Prn.TextHeight(strTextTrun)) / 2) + py
           
            Select Case objCol.Alignment
                Case ListColumnAlignmentConstants.lvwColumnCenter
                    Prn.CurrentX = lngX + intOffset + (((objCol.Width * dblXScale) - Prn.TextWidth(strTextTrun)) / 2)
                Case ListColumnAlignmentConstants.lvwColumnLeft
                    Prn.CurrentX = lngX + intOffset + (px * 5)
                Case ListColumnAlignmentConstants.lvwColumnRight
                    Prn.CurrentX = lngX + ((objCol.Width * dblXScale) - intOffset - Prn.TextWidth(strTextTrun)) - (px * 5)
            End Select
           
            '------------------------------------------------------------------
            'Print each colum
            '------------------------------------------------------------------
            Prn.Print strTextTrun
            lngX = lngX1
        Next
     Next
    
     '--------------------------------------------------------------------------
     'Print final page number
     '--------------------------------------------------------------------------
     lngPageNo = lngPageNo + 1
     Prn.CurrentX = (Prn.Width / 2) - (Prn.TextWidth("第 " & lngPageNo & " 页") / 2)
     Prn.CurrentY = lngEOP - intRowHeight
     Prn.Print "第 " & lngPageNo & " 页"
     Prn.EndDoc
     gPrintListView = True
     Screen.MousePointer = vbDefault
    
     Set objCol = Nothing
     Set objILS = Nothing
     Set objLI = Nothing
     Set objPic = Nothing
    
Exit Function
    
Error_Handler:
     Set objCol = Nothing
     Set objILS = Nothing
     Set objLI = Nothing
     Set objPic = Nothing
     Screen.MousePointer = vbDefault
     '--------------------------------------------------------------------------
     'Simple error message reporting
     '--------------------------------------------------------------------------
     MsgBox "系统打印出错:-" & vbCrLf & vbCrLf & _
     "错误号: " & Err.Number & vbCrLf & "错误内容:" & Err.Description, vbExclamation
End Function

'Michael Added This function @ Aug,6,07
Function StatusLogFolder(strSysPath As String) As Boolean


End Function
