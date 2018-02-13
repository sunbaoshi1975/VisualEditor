Attribute VB_Name = "Mdlnode"
'**********************************************************************************************
'Version      : 1.1
'Date         : 6-28-07
'Modify       : Node40 - rectime , change the data type from byte to integer(2Bytes)
'Author       : Michael
'---------------------------------------------------------------------------------------------
'Last Version : 1.2
'Date         : 7-4-07
'Modify       : Node61 - New add 4 members - usevar,switchtype,waitmethod,nd_wait
'Author       : Michael
'--------------------------------------------------------------------------------------------
'Last Version : 1.3
'Date         : July-9-07
'Modify       : Node61 - New add 3 members - var_userid,var_loginid, readEWT
'Author       : Michael
'--------------------------------------------------------------------------------------------
'Last Version : 1.4
'Date         : July-10-07
'Modify       : Node19 - Modify Data1 struct and add a new member leavequeue(Byte)
'               NOde61 - Data1 struct add new member waitansto(Byte)
'Author       : Michael
'--------------------------------------------------------------------------------------------
'Lase File Version : 1.6
'Last Priduct Version : 6.9.4
'Date        : May-27-2008
'Modify      : Node60 - Added a new field "agtinfoLen" to Data2
'Author      : Mike
'-----------------------------------------------------------------------------------
'Lase File Version : 1.7
'Last Priduct Version : 6.10.1
'Date        : 2012-04-18
'Modify      : Node62 - Change to "Setup Conference"
'Modify      : Node40 - Add property "NotifyPL"
'Author      : Tony
'-----------------------------------------------------------------------------------
'Lase File Version : 1.8
'Last Product Version : 6.10.2
'Date        : 2012-05-07
'Modify      : Node00 - Add property "nd_BeforeHookOn"
'Author      : Tony
'-----------------------------------------------------------------------------------

Option Explicit
Option Base 0
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' ============================================
' ��������
' 0 - Init
' 1 - Modify
' 2 - New
' 3 - Delete
'
' ============================================
Public Const DEF_OPERATION_INIT = 0
Public Const DEF_OPERATION_MODIFY = 1
Public Const DEF_OPERATION_NEW = 2
Public Const DEF_OPERATION_DELETE = 3

'' Sun added 2006-06-20
Public Const DEF_NODE_DATA1_LEN = 13
Public Const DEF_NODE_DATA2_LEN = 64
'----------------------------------------------------------------------

Public gNodeTypeNameList(255) As String

'------------------------------------------------------------------------
' �ڵ�000��ϵͳ�ڵ�
Public Type SData1_000
    reserved1(9) As Byte        ' ����
    Languages As Byte           ' ��������
    MajorVer As Byte            ' ���汾
    MinorVer As Byte            ' �ΰ汾
End Type
Public Node0_Data1 As SData1_000
' �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_000
    key_repeat As Byte          ' �ظ���ǰ�ڵ㰴��
    key_return As Byte          ' �ص���һ���ڵ㰴��
    key_root As Byte            ' �ص����˵�����
    
    'reserved1(44) As Byte      ' ���� sun 2002-12-03 Old
    reserved1 As Byte           ' ���� sun 2002-12-03
    ResourceProject As Integer  ' ��Դ��ĿID sun 2002-12-03
    MainCOM As Integer          ' ����COM�����Դ��� sun 2002-12-03
    LogSwitchOff As Byte        ' �ڵ���־ȫ�ֿ��ƿ���, sun added 2008-02-19
    reserved2(38) As Byte       ' ���� sun 2002-12-03
    
    nd_parent As Integer        ' ���ڵ�ID
    nd_root As Integer          ' ���˵�(��)ID
    nd_SysSendData As Integer   ' ϵͳȱʡ�������ݸ�ʽ����ڵ� sun 2004-12-30
    nd_BeforeHookOn As Integer  ' �һ�ǰת�ڵ㣬Sun added 2012-05-07
    reserved3(7) As Byte        ' ����
End Type
Public Node0_Data2 As SData2_000
Public Default As SData2_000

'------------------------------------------------------------------------
' �ڵ�001����������
' �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_001
     reserved1(10) As Byte      ' ����
     uservars As Byte           ' �û����������, 0-255
     reserved2 As Byte
End Type
Public Node1_Data1 As SData1_001
Public Defaultuservar As SData1_001

'------------------------------------------------------------------------
' �ڵ�002�������嵥
' �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_002
    reserved1(12) As Byte
End Type
Public Node2_Data1 As SData1_002
' �û���������ṹ
Public Type SData2_002
    uservar(63) As Byte
End Type
Public Node2_Data2 As SData2_002
Public Defaulttype As SData2_002

'------------------------------------------------------------------------
' �ڵ�006��������ת��
' �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_006
    Sleep As Integer            ' ��ʱ��sleep������λΪ����
    reserved1(10) As Byte       ' ����
End Type
Public Node6_Data1 As SData1_006

'�ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_006
    nd_goto As Integer          ' ��ת�ڵ�ID
    reserved1(46) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    reserved2(14) As Byte       ' ����
End Type
Public Node6_Data2 As SData2_006

'------------------------------------------------------------------------
' �ڵ�007�������֤
' �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_007
    timeout As Byte             ' �ڵ㳬ʱ(��)
    maxuserid As Byte           ' �û�������󳤶�
    maxpassword As Byte         ' �û�������󳤶�
    key_term As Byte            ' ������ֹ��, ����
    maxtrytime As Byte          ' ����Դ���
    log As Byte                 ' ��������־
    var_trytime As Byte         ' ��֤������¼(0 - maxtrytime)
    var_result As Byte          ' ��֤�����¼
    var_userid As Byte          ' �û������¼
    var_password As Byte        ' �û������¼
    reserved1(2) As Byte        ' ����
End Type
Public Node7_Data1 As SData1_007
' �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_007
    vox_userid As Integer       ' ��ʾ�û�������벥������
    vox_password As Integer     ' ��ʾ�û�������������
    vox_tryagain As Integer     ' ��ʾ�û��������벥������
    reserved1(9) As Byte        ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_succeed As Integer       ' �ɹ�ת�ڵ�ID
    nd_fail As Integer          ' ʧ��ת�ڵ�ID
    reserved3(9) As Byte        ' ����
End Type
Public Node7_Data2 As SData2_007

'------------------------------------------------------------------------
' �ڵ�008���޸Ŀ���
' �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_008
    timeout As Byte             ' �ڵ㳬ʱ(��)
    reserved1 As Byte           ' ����
    maxpassword As Byte         ' �û�������󳤶�
    key_term As Byte            ' ������ֹ��, ����
    maxtrytime As Byte          ' ����Դ���
    log  As Byte                ' ��������־
    var_trytime As Byte         ' ���Դ�����¼(0 - maxtrytime)
    var_result As Byte          ' ��֤�����¼
    reserved2 As Byte           ' ����
    var_password As Byte        ' �û������¼
    reserved3(2) As Byte        ' ����
End Type
Public Node8_Data1 As SData1_008
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_008
    vox_password As Integer     ' ��ʾ�û������¿��������
    vox_confirm As Integer      ' ��ʾ�û��ٴ�ȷ�ϲ�������
    vox_tryagain As Integer     ' ���β�һ���������벥������
    vox_succeed As Integer      ' ��ʾ�û��޸ĳɹ���������
    reserved1(7) As Byte        ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_succeed As Integer       ' �ɹ�ת�ڵ�ID
    nd_fail As Integer          ' ʧ��ת�ڵ�ID
    reserved3(9) As Byte        ' ����
End Type
Public Node8_Data2 As SData2_008

'------------------------------------------------------------------------
' �ڵ�009��ʱ���֧
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_009
    reserved1(4) As Byte        ' ����
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node9_Data1 As SData1_009
'/ ʱ��νṹ
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_009
    workday As Byte             ' �����հ���, Bit����, 0-6λ��Ч,0:��Ϣ��;1:������
    worktime As Byte            ' ������ʱ��ΰ���, Bit����, 0-5λ��Ч,0:��Ч;1:��Ч
    timesec(23) As Byte         ' ʱ���1-6
    reserved1(21) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_sparetime As Integer     ' ��Ϣ��ת�ڵ�ID
    nd_timesec(5) As Integer    ' ʱ���1-6ת�ڵ�ID
End Type
Public Node9_Data2 As SData2_009

'------------------------------------------------------------------------
' �ڵ�010���������趨
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_010
    maincalendar As Byte        ' �Ƿ��������������ɼ���ϵͳʶ��
    startyear As Byte           ' ��ʼ��ݣ�YY
    startmonth As Byte          ' ��ʼ�£�1-12
    monthcount As Byte          ' �������£����12���£�
    reserved1 As Byte           ' ����
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node10_Data1 As SData1_010
'/ ʱ��νṹ
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_010
    daytype(47) As Byte         ' λ����: 0 - �����գ�1 - �ڼ���
    nd_parent As Integer        ' ���ڵ�ID
    nd_daysec(2) As Integer     ' ת�ڵ�ID
    reserved1(7) As Byte        ' ����
End Type
Public Node10_Data2 As SData2_010

'------------------------------------------------------------------------
' Sun added 2004-12-30
' �ڵ�016��������֧
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_016
    reserved1(4) As Byte        ' ����
    log As Byte                 ' ��������־
    var_id As Byte              ' ����
    logic As Byte               ' �߼������
    convert As Byte             ' ת����ʽ
    param1 As Byte              ' ����1
    param2 As Byte              ' ����2
    reserved2(1) As Byte        ' ����
End Type
Public Node16_Data1 As SData1_016
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_016
    var_value(47) As Byte       ' ����ֵ
    nd_parent As Integer        ' ���ڵ�ID
    nd_succ As Integer          ' ��������ת�ƽڵ�ID
    nd_fail As Integer          ' ����������ת�ƽڵ�ID
    reserved3(9) As Byte        ' ����
End Type
Public Node16_Data2 As SData2_016

''/ �߼����������
Public Const DEF_NODE016_LOGIC_EQUE = 0                   ' =
Public Const DEF_NODE016_LOGIC_BIGGER = 1                 ' >
Public Const DEF_NODE016_LOGIC_SMALLER = 2                ' <
Public Const DEF_NODE016_LOGIC_EB = 3                     ' >=
Public Const DEF_NODE016_LOGIC_ES = 4                     ' <=
Public Const DEF_NODE016_LOGIC_NE = 5                     ' <>
Public Const DEF_NODE016_LOGIC_LIKE = 6                   ' Like

''/ ת����ʽ����
Public Const DEF_NODE016_CONVERT_NONE = 0                 ' Do Nothing
Public Const DEF_NODE016_CONVERT_LEFT = 1                 ' Left(n)
Public Const DEF_NODE016_CONVERT_RIGHT = 2                ' Right(n)
Public Const DEF_NODE016_CONVERT_MID = 3                  ' Mid(m, n)
Public Const DEF_NODE016_CONVERT_LEN = 4                  ' Len()

'------------------------------------------------------------------------
' Sun added 2003-04-24
' �ڵ�017����������ѡ��
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_017
    timeout As Byte             ' �ڵ㳬ʱ(��)
    maxtrytime As Byte          ' ����Դ���
    var_lang As Byte            ' ѡ������¼
    reserved1(9) As Byte        ' ����
End Type
Public Node17_Data1 As SData1_017
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_017
    vox_play As Integer         ' ��������
    lang(9) As Byte             ' ���Զ�Ӧ����
    reserved1(35)  As Byte      ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_succ As Integer          ' �ɹ��ڵ�ID
    nd_fail As Integer          ' ʧ�ܽڵ�ID
    reserved2(9) As Byte        ' ����
End Type
Public Node17_Data2 As SData2_017

'------------------------------------------------------------------------
' �ڵ�018����������
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_018
    seperator As Byte           ' �ָ���
    reserved1(11) As Byte       ' ����
End Type
Public Node18_Data1 As SData1_018
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_018
    typeflags(14) As Byte       ' ��ǣ�AppData/UserData
    prefix1(14) As Byte         ' ����ǰ׺1
    prefix2(14) As Byte         ' ����ǰ׺2
    valueid(14) As Byte         ' ����ID
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
End Type
Public Node18_Data2 As SData2_018

'------------------------------------------------------------------------
' �ڵ�019���޲���
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_019
    'Michael Modified @Jul,10,07
    'reserved1(12) As Byte       ' ����
    reserved1(11) As Byte       ' ����
    'Michael Added @ the same day
    leavequeue    As Byte
End Type
Public Node19_Data1 As SData1_019
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_019
    delaytime   As Integer      ' ��ʱʱ��(ms)
    reserved1(45) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    reserved2(13) As Byte       ' ����
End Type
Public Node19_Data2 As SData2_019

'------------------------------------------------------------------------
' �ڵ�020�������һ�
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_020
    reserved1 As Byte           ' ����
    playclear As Byte           ' ������ձ�־: 0 -- �����;1 -- ���
    reserved2(2) As Byte        ' ����
    log As Byte                 ' ��������־
    reserved3(6) As Byte        ' ����
End Type
Public Node20_Data1 As SData1_020
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_020
    vox_play As Integer         ' ��������
    reserved1(45) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    reserved2(13) As Byte       ' ����
End Type
Public Node20_Data2 As SData2_020

'------------------------------------------------------------------------
' �ڵ�021����������
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_021
    reserved1 As Byte           ' ����
    playclear As Byte           ' ������ձ�־: 0 -- �����;1 -- ���
    playtype As Byte            ' ��������(�����):0����1���2����4��ĸ8����16����
    usevar As Byte              ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    breakkey As Byte            ' �жϰ���
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node21_Data1 As SData1_021
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_021
    vox_pred As Integer         ' ǰ����������
    vox_succ As Integer         ' ������������
    reserved1(11) As Byte       ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved3(11) As Byte       ' ����
End Type
Public Node21_Data2 As SData2_021

'---------------------------------------------------------------------------------
' �ڵ�022�������ȴ�����
'' �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_022
    timeout As Byte             ' �ڵ㳬ʱ(��)
    playclear As Byte           ' ������ձ�־: 0 -- �����;1 -- ���
    getlength As Byte           ' �û����볤��
    maxinterval As Byte         ' ���������(��)
    breakkey As Byte            ' �жϰ���
    log As Byte                 ' ��������־
    var_key As Byte             ' �û�������¼, 1-255����ID
    maxtrytime As Byte          ' ����Դ���
    reserved1(4) As Byte        ' ����
End Type
Public Node22_Data1 As SData1_022
' �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_022
    vox_play As Integer         ' ��������
    vox_nodefail As Integer     ' �ڵ�ʧ�ܲ�������[V2]
    reserved1(27) As Byte       ' ����(27)
    nd_parent As Integer        ' ���ڵ�ID
    nd_key(11) As Integer       ' �ӽڵ�ID: 0 - 11
    nd_nodefail As Integer      ' �ӽڵ�ID: ʧ��ת��ڵ�[V2]
    reserved2(3) As Byte        ' ����(3)
End Type
Public Node22_Data2 As SData2_022

'------------------------------------------------------------------------
' �ڵ�023������ת��
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_023
    timeout As Byte             ' �ڵ㳬ʱ(��)
    playclear As Byte           ' ������ձ�־: 0 -- �����;1 -- ���
    breakkey As Byte            ' �жϰ���
    reserved1(1) As Byte        ' ����
    log As Byte                 ' ��������־
    var_play As Byte            ' ����, 1-255����ID
    reserved2(5) As Byte        ' ����
End Type
Public Node23_Data1 As SData1_023
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_023
    vox_play As Integer         ' ����������var_play = 0ʱ��Ч
    reserved1(45) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_goto As Integer          ' ת�ƽڵ�ID
    reserved2(11) As Byte       ' ����
End Type
Public Node23_Data2 As SData2_023

'------------------------------------------------------------------------
' Sun added 2004-12-30
' �ڵ�028��TTS����
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_028
    timeout As Byte             ' �ڵ㳬ʱ(��)
    playclear As Byte           ' ������ձ�־: 0 -- �����;1 -- ���
    usevar As Byte              ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    breakkey As Byte            ' �жϰ���
    reserved1 As Byte           ' ����
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node28_Data1 As SData1_028
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_028
    vox_string As Integer       ' �����ַ�����֧��"%s"�����Ӵ�
    vox_alter As Integer        ' TTS������ʱ�������
    playtype As Byte            ' �������ͣ�
                                '' 0 - �ַ���
                                '' 1 - �ı��ļ���vox_string��ʽ������Ϊ�ļ�·����
                                '' 2 - ExtData
                                '' 3 - UserData
    reserved1(10) As Byte       ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_succ As Integer          ' TTS���ųɹ�ת�ƽڵ�ID
    nd_fail As Integer          ' TTS����ʧ��ת�ƽڵ�ID
    reserved3(9) As Byte        ' ����
End Type
Public Node28_Data2 As SData2_028

'------------------------------------------------------------------------
' �ڵ�040����������
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_040
    rectime As Byte             ' ¼��ʱ��(��)
    playclear As Byte           ' ������ձ�־: 0 -- �����;1 -- ���
    breakkey As Byte            ' �жϰ���
    var_agent As Byte           ' Agent[V2]
    var_filename As Byte        ' ¼���ļ�����¼����
    log As Byte                 ' ��������־
    var_appfield(2) As Byte     ' Ӧ��������
    maxsilencetime As Byte      ' �����ʱ��(��)
    vmsclass As Byte            ' ���Է���[2007-02-28]
    MinRecLength As Byte        ' ���¼��ʱ��(��)[2007-03-20]
    toneoff As Byte             ' ������¼����ʼ��ʾ[2009-07-24]��0 - ���ţ�1 - ������
End Type
Public Node40_Data1 As SData1_040
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_040
    vox_op As Integer           ' ¼��ǰ��ʾ����
    'new add
    recfiletype As Byte         ' ¼���ļ�����[2007-06-28]: 0 - vox; 1 - wav
    rectime_ho As Byte          ' ¼��ʱ����8λ
    var_notifyintvl As Byte     ' ��ʾ���(��)[2009-07-24]��0 - ����ʾ
    var_rectime As Byte         ' ¼��ʵ��ʱ����¼����(��)[2009-07-24]
    NotifyPL As Byte            ' �Ƿ�֪ͨPL¼��ϵͳ[2012-04-18]
    reserved1(40) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved2(11) As Byte       ' ����
End Type
Public Node40_Data2 As SData2_040

''/ ���Է��ඨ��
Public Const DEF_NODE040_VMSCLASS_UNKNOWN = 0               ' δ֪
Public Const DEF_NODE040_VMSCLASS_ALL = 1                   ' ����
Public Const DEF_NODE040_VMSCLASS_GROUP = 2                 ' ��
Public Const DEF_NODE040_VMSCLASS_EXT = 3                   ' �ֻ�
Public Const DEF_NODE040_VMSCLASS_USER = 4                  ' �û�

'------------------------------------------------------------------------
' �ڵ�041���쿴����
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_041
    timeout As Byte             ' �ڵ㳬ʱ(��)
    playclear As Byte           ' ������ձ�־: 0 -- �����;1 -- ���
    breakkey As Byte            ' �жϰ���
    var_agent As Byte           ' Agent[V2]
    reserved1 As Byte           ' ����
    log As Byte                 ' ��������־
    vmstype As Byte             ' �쿴��������
    closewhencheck As Byte      ' �쿴���Զ���Ϊ�ر�״̬
    vmsclass As Byte            ' ���Է���[2007-02-28]
    reserved2(3) As Byte        ' ����
 End Type
 Public Node41_Data1 As SData1_041
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_041
    vox_play(3) As Integer      ' ����״̬����ǰ��/����״̬�������/�����ʾ����/������ʾ����
    key_op(7) As Byte           ' ����һ�����԰���/ǰһ�����԰���/����һ�����԰���/�����һ�����԰���/������ǰ���԰���/ɾ����ǰ���԰���/�˳����ڵ㰴��/ת�����Ͱ���
    reserved1(31) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved2(11) As Byte       ' ����
End Type
Public Node41_Data2 As SData2_041

''/ �쿴�������Ͷ���
Public Const DEF_NODE041_VMSTYPE_NEW = 0                  ' �����ԣ���ɾ�����浵
Public Const DEF_NODE041_VMSTYPE_CLOSED = 1               ' �浵���ԣ���ɾ��
Public Const DEF_NODE041_VMSTYPE_DELETED = 2              ' ɾ�����ԣ���ȡ��ɾ����Ϊ�浵

'------------------------------------------------------------------------
' Sun added 2006-12-31, V6.5.11
'' ���촫���ļ���
Public Const DEF_NODE050_FAXN_TYPE_AUTO = 0               ' �Զ�����
Public Const DEF_NODE050_FAXN_TYPE_RESID = 1              ' ��ԴID
Public Const DEF_NODE050_FAXN_TYPE_VAR2RESID = 2          ' ������Ӧ��ԴID
Public Const DEF_NODE050_FAXN_TYPE_VAR2NAME = 3           ' ������Ӧ�ļ���
Public Const DEF_NODE050_FAXN_TYPE_FORMAT = 4             ' �����滻��Դ�е�ͨ���

' Sun updated 2006-12-31, V6.5.11
' �ڵ�050���򵥴���
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_050
    timeout As Integer          ' �ڵ㳬ʱ(��)
    filenametype As Byte        ' �����ļ�������:
                                '' 1 -- ��ԴID
                                '' 2 -- ������Ӧ��ԴID
                                '' 3 -- ������Ӧ�ļ���
                                '' 4 -- �����滻��Դ�е�ͨ���
                                '' һ�ζ�����Ͷ���ļ��÷ֺŷָ�����5��
    trytimes As Byte            ' ���Դ���
    record_cdr As Byte          ' �Ƿ��¼���淢���굥: 0 --����¼��1 -- ��¼
    log As Byte                 ' ��������־
    var_faxfile  As Byte        ' �����ļ�ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_fromno As Byte          ' ��������ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_tono As Byte            ' ���պ���ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_result As Byte          ' �����¼����ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_appfield(2) As Byte     ' Ӧ��������
End Type
Public Node50_Data1 As SData1_050
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_050
    vox_op As Integer           ' ������ʾ����
    fax_fileid As Integer       ' �����ļ�(��ԴID)
    header_id As Integer        ' �������(��ԴID)
    reserved1(41) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_succ As Integer          ' �ɹ�ת�ƽڵ�ID
    nd_fail As Integer          ' ʧ��ת�ƽڵ�ID
    reserved2(9) As Byte        ' ����
End Type
Public Node50_Data2 As SData2_050

'------------------------------------------------------------------------
' �ڵ�051��TTF����
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_051
    timeout As Byte             ' �ڵ㳬ʱ(��)
    reserved1(3) As Byte        ' ����
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node51_Data1 As SData1_051
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_051
    vox_op As Integer           ' ������ʾ����
    fax_logo As Integer         ' LOGO�ļ�(��ԴID)
    fax_format As Integer       ' ����ʽ�ļ�(��ԴID)
    header_id As Integer        ' �������(��ԴID)
    from_id As Integer          ' ��������(��ԴID)
    reserved1(5) As Byte        ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved3(11) As Byte       ' ����
End Type
Public Node51_Data2 As SData2_051


'------------------------------------------------------------------------
' Sun added 2006-12-31, V6.5.11
' �ڵ�055���������
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_055
    timeout As Integer          ' �ڵ㳬ʱ(��)
    filenametype As Byte        ' �����ļ�������:
                                '' 0 -- �Զ����ɣ�$Record\Group%n\%Y%m%d%H%M%S_c<CH>.tif��
                                '' 1 -- ��ԴID
                                '' 2 -- ������Ӧ��ԴID
                                '' 3 -- ������Ӧ�ļ���
                                '' 4 -- �����滻��Դ�е�ͨ���
    var_faxfile As Byte         ' �����ļ�ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    record_cdr As Byte          ' �Ƿ��¼���淢���굥: 0 --����¼��1 -- ��¼
    log As Byte                 ' ��������־
    var_fromno As Byte          ' ��������ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_tono As Byte            ' ���պ���ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_extno As Byte           ' �ֻ������¼����ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_result As Byte          ' �����¼����ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_appfield(2) As Byte     ' Ӧ��������
End Type
Public Node55_Data1 As SData1_055
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_055
    vox_op As Integer           ' ������ʾ����
    fax_fileid As Integer       ' �����ļ�(��ԴID)
    reserved1(43) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_succ As Integer          ' �ɹ�ת�ƽڵ�ID
    nd_fail As Integer          ' ʧ��ת�ƽڵ�ID
    reserved2(9) As Byte        ' ����
End Type
Public Node55_Data2 As SData2_055

'------------------------------------------------------------------------
' �ڵ�060��ת����ϯ
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_060
    timeout As Byte             ' �ڵ㳬ʱ(��)
    switchtype As Byte          ' ת�ӷ�ʽ(0���Զ����ӣ�1��ָ����ϯ��2���û�����)
    agentid As Byte             ' ָ����ϯID, ��ʽ1ʱ / �жϰ���, ��ʽ2ʱ
    getlength As Byte           ' �û����볤��, ת�ӷ�ʽΪ2ʱ��Ч, ת�ӷ�ʽΪ1ʱΪָ����ϯID�ĸ�λ
    looptimes As Byte           ' �ȴ�ѭ�����Ŵ���
    log As Byte                 ' ��������־
    var_key As Byte             ' �û�������¼, ת�ӷ�ʽΪ2ʱ��Ч��1-255����ID
    agentinfo As Byte           ' �Ƿ�������ϯ��Ϣ
    reserved1(4) As Byte        ' ����
End Type
Public Node60_Data1 As SData1_060
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_060
    vox_op As Integer           ' ������ʾ����
    vox_sw As Integer           ' ת����ʾ����
    vox_wt As Integer           ' �ȴ�ѭ����������
    vox_nobody As Integer       ' û���ϰ���ʾ����
    vox_busy As Integer         ' ��ϯæ��ʾ����
    vox_noanswer As Integer     ' ��ϯ��Ӧ����ʾ����
    vox_ok   As Integer         ' ��ϯת�ӳɹ�
    'Mike Added @ 2008-5-27     # moved a Byte storage space from reserved1(33)
    length_agentinfo As Byte    ' ��ϯ��Ϣ���磺���ţ�λ����Ĭ��4λ��������ǰ�油0��0����ԭʼ����
    reserved1(32) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_nobody As Integer        ' û���ϰ�ת�ڵ�ID
    nd_busy As Integer          ' ��ϯæת�ڵ�ID
    nd_noanswer As Integer      ' ��ϯ��Ӧ��ת�ڵ�ID
    nd_ok As Integer            ' ת�ӳɹ�ת�ڵ�ID
    reserved2(5) As Byte        ' ����
End Type
Public Node60_Data2 As SData2_060

'------------------------------------------------------------------------
' �ڵ�061��ת����ϯ��
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_061
    maxwait      As Integer   ' �ȴ���ʱ(��)��0 ��ʾ���޵ȴ�
    toacd        As Byte      ' ת��ACD����RoutePoint��0 �� RoutePoint��1 �� ACD
    usevar       As Byte      ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    looptimes    As Byte      ' �ȴ�ѭ�����Ŵ�����0 ��ʾ����ѭ��
    log          As Byte      ' ��������־
    agentinfo    As Byte      ' �Ƿ�������ϯ��Ϣ
    readEWT      As Byte      ' �Ƿ񲥱�Ԥ��ȴ�ʱ��, 0 - ������;1 - ����(��);2 - ����(ʱ����)
    switchtype   As Byte      ' ת�ӷ�ʽ: 0 - ���ţ�1 - CTI
    waitmethod   As Byte      ' �ȴ���ʽ: 0 - ����������1 - �첽������
    var_userid   As Byte      ' ��¼��ϯID��¼����ID  : 0 --����¼;1 -- 255 ����ID
    var_loginid  As Byte      ' ��¼��ϯ���ż�¼����ID: 0 --����¼;1 -- 255 ����ID
    waitansto    As Byte      ' �ȴ���ϯӦ��ʱ(��)��0 ��ʾ���޵ȴ�
End Type
Public Node61_Data1 As SData1_061
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_061
    vox_op        As Integer     ' ������ʾ����
    vox_sw        As Integer     ' ת����ʾ����
    vox_wt        As Integer     ' �ȴ�ѭ����������
    vox_nobody    As Integer     ' û���ϰ���ʾ����
    vox_busy      As Integer     ' ��ϯ��ȫæ��ʾ����
    vox_noanswer  As Integer     ' ��ϯ��Ӧ����ʾ����
    vox_ok        As Integer     ' ��ϯ��ת�ӳɹ�
    routepointid  As Integer     ' ת�ӵ�·�ɵ���
    acddn(19)     As Byte        ' ת�ӵ�ACD����
    length_agentinfo As Byte     ' ��ϯ��Ϣ���磺���ţ�λ����Ĭ��4λ��������ǰ�油0��0����ԭʼ����
    reserved1(9) As Byte
    waitansto_hi  As Byte        '����ȴ���ϯӦ��ʱ�ĸ�8λ
    nd_parent     As Integer     ' ���ڵ�ID
    nd_nobody     As Integer     ' û���ϰ�ת�ڵ�ID
    nd_busy       As Integer     ' ��ϯ��ȫæת�ڵ�ID
    nd_noanswer   As Integer     ' ��ϯ��Ӧ��ת�ڵ�ID
    nd_ok         As Integer     ' ת�ӳɹ�ת�ڵ�ID
    nd_wait       As Integer     ' �ȴ�������ڽڵ��
    reserved2(3)  As Integer     ' ����
End Type
Public Node61_Data2 As SData2_061

'------------------------------------------------------------------------
' �ڵ�062��������飬2012-04-17
'' �ڵ�062����ǿת����ϯ
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
'Public Type SData1_062
'    reserved1(2) As Byte        ' ����
'    usevar As Byte              ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
'    looptimes As Byte           ' �ȴ�ѭ�����Ŵ���
'    log As Byte                 ' ��������־
'    reserved2(6) As Byte        ' ����
'End Type
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_062
    timeout As Byte             ' �ڵ㳬ʱ(��)
    reserved1 As Byte           ' ����
    usevar As Byte              ' ʹ�ñ���ID: 0 -- ��ʹ�ã�1 -- 255 ����ID
    looptimes As Byte           ' �ȴ�ѭ�����Ŵ�����0 ��ʾʹ��ϵͳĬ�ϴ���
    switchtype As Byte          ' ת�ӷ�ʽ: 0 - ���ţ���TAPI������Ч����1 - CTI
    log          As Byte        ' ��������־
    waitansto    As Byte        ' �ȴ�Ӧ��ʱ(��)��0 ��ʾ���޵ȴ�
    var_waitansto As Byte       ' �ȴ���ϯӦ��ʱ����ID: 0 --��ʹ�ã�1 -- 255 ����ID, Sun added 2014-01-29
    reserved2(4) As Byte        ' ����
End Type
Public Node62_Data1 As SData1_062
''/ �ڵ����ݽṹ-Data2(30-6F/64b)
'Public Type SData2_062
'    vox_op As Integer           ' ������ʾ����
'    vox_sw As Integer           ' ת����ʾ����
'    vox_wt As Integer           ' �ȴ�ѭ����������
'    vox_nobody As Integer       ' û���ϰ���ʾ����
'    vox_busy As Integer         ' ��ϯȫæ��ʾ����
'    vox_ok As Integer           ' ��ϯת�ӳɹ�
'    reserved1(3) As Byte        ' ����
'    com_iid As Integer          ' COM�ӿ�ID
'    reserved2(29) As Byte       ' ����
'    nd_parent As Integer        ' ���ڵ�ID
'    nd_nobody As Integer        ' û���ϰ�ת�ڵ�ID
'    nd_busy As Integer          ' ��ϯȫæת�ڵ�ID
'    nd_ok As Integer            ' ת�ӳɹ�ת�ڵ�ID
'    reserved3(7) As Byte        ' ����
'End Type
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_062
    vox_op As Integer           ' ������ʾ����
    vox_sw As Integer           ' ������������ʾ����
    vox_wt As Integer           ' �ȴ�ѭ����������
    vox_noconf As Integer       ' ����ʧ����ʾ����
    vox_noans As Integer        ' ��Ӧ����ʾ����
    vox_ans As Integer          ' �ɹ�Ӧ����ʾ����
    vox_ok As Integer           ' ��ɻ�����ʾ����
    vox_syserror As Integer     ' ϵͳ������ʾ����
    DialNo(23) As Byte          ' ������룬Ҳ������:������
    predial(5) As Byte          ' ����ǰ׺
    reserved1(1) As Byte        ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_noans As Integer         ' ��Ӧ��ת�ڵ�ID
    nd_failed As Integer        ' ����ʧ��ת�ڵ�ID
    nd_ok As Integer            ' ����ɹ�ת�ڵ�ID
    reserved2(7) As Byte        ' ����
End Type
Public Node62_Data2 As SData2_062

'------------------------------------------------------------------------
' �ڵ�063����ǿת����ϯ��
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_063
    reserved1(2) As Byte        ' ����
    usevar As Byte              ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    looptimes As Byte           ' �ȴ�ѭ�����Ŵ���
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node63_Data1 As SData1_063
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_063
    vox_op As Integer           ' ������ʾ����
    vox_sw As Integer           ' ת����ʾ����
    vox_wt  As Integer          ' �ȴ�ѭ����������
    vox_nobody As Integer       ' û���ϰ���ʾ����
    vox_busy  As Integer        ' ��ϯȫæ��ʾ����
    vox_ok   As Integer         ' ��ϯת�ӳɹ�
    reserved1(35) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_nobody  As Integer       ' û���ϰ�ת�ڵ�ID
    nd_busy As Integer          ' ��ϯȫæת�ڵ�ID
    nd_ok  As Integer           ' ת�ӳɹ�ת�ڵ�ID
    reserved2(7) As Byte        ' ����
End Type
Public Node63_Data2 As SData2_063

'------------------------------------------------------------------------
' �ڵ�069��ת����ֻ�
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_069
    reserved1(2) As Byte        ' ����
    usevar As Byte              ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    switchtype As Byte          ' ת�ӷ�ʽ: 0 - ���ţ�1 - CTI
    log As Byte                 ' ��������־
    vagency As Long             ' ����ֻ�����
    reserved3(2) As Byte        ' ����
End Type
Public Node69_Data1 As SData1_069
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_069
    vox_op As Integer           ' ������ʾ����
    maxtry As Integer           ' ����Դ�����2012-04-18
    tryinterval As Integer      ' ���Լ����2012-04-18
    reserved1(41) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child  As Integer        ' �ӽڵ�ID
    reserved2(11) As Byte       ' ����
End Type
Public Node69_Data2 As SData2_069

'------------------------------------------------------------------------
' �ڵ�070����ѯ·�ɵ�
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_070
    timeout As Byte             ' �ڵ㳬ʱ(��)
    reserved1 As Byte           ' ����
    usevar As Byte              ' ʹ�ñ���ID: 0 -- ��ʹ�ã�1 -- 255 ����ID
    paramindex As Byte          ' �������
    logic As Byte               ' �߼������
    log As Byte                 ' ��������־
    querytype As Byte           ' ��ѯ���ͣ�0 - ·�ɵ㣻1 - ���У�2 - �飻3 - С��
    var_result As Byte          ' ��ѯ�����¼����: 0 -- ��ʹ�ã�1 -- 255 ����ID��[2007-04-12]
    reserved2(4) As Byte        ' ����
End Type
Public Node70_Data1 As SData1_070
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_070
    routepointid As Integer     ' ·�ɵ���
    comparedvalue As Integer    ' �ο�ֵ
    reserved1(11) As Byte       ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_yes As Integer           ' ��������ת�ڵ�ID
    nd_no As Integer            ' ����������ת�ڵ�ID
    nd_fail As Integer          ' ��ѯʧ�ܽڵ�ID
    reserved3(7) As Byte        ' ����
End Type
Public Node70_Data2 As SData2_070

'------------------------------------------------------------------------
' �ڵ�071����ѯ��ϯ״̬
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_071
    timeout As Byte             ' �ڵ㳬ʱ(��)
    usevar As Byte              ' ʹ�ñ���ID: 0 -- ��ʹ�ã�1 -- 255 ����ID
    dn_status As Byte           ' DN״̬���
    pos_status As Byte          ' POS״̬���
    dn_logic As Byte            ' DN�߼������
    pos_logic As Byte           ' POS�߼������
    log As Byte                 ' ��������־
    querytype As Byte           ' ��ѯ���ͣ�0 - Agent; 1 - User (����ʹ�ñ���ȷ���û���); 2 - ����
    conditions As Byte          ' ��������������
    reserved1(3) As Byte        ' ����
End Type
Public Node71_Data1 As SData1_071
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_071
    agentid As Long             ' ��ϯ��Ż򹤺�
    reserved1(43) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_yes As Integer           ' ��������ת�ڵ�ID
    nd_no As Integer            ' ����������ת�ڵ�ID
    nd_fail As Integer          ' ��ѯʧ�ܽڵ�ID
    reserved2(7) As Byte        ' ����
End Type
Public Node71_Data2 As SData2_071

''/ ��������
Public Const DEF_NODE071_CONDITION_NONE = 0
Public Const DEF_NODE071_CONDITION_FIRST = 1
Public Const DEF_NODE071_CONDITION_SECOND = 2
Public Const DEF_NODE071_CONDITION_BOTH = 3
Public Const DEF_NODE071_CONDITION_EITHER = 4

'------------------------------------------------------------------------
' �ڵ�080������ACD�Ŷ�
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_080
    reserved1(2) As Byte        ' ����
    usevar As Byte              ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    reserved2 As Byte           ' ����
    log As Byte                 ' ��������־
    maxwaittime As Integer      ' ��Ŷӵȴ�ʱ��(��)
    reserved3(4) As Byte        ' ����
End Type
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_080
    vox_wait As Integer         ' �Ŷӵȴ�����
    reserved1(13) As Byte       ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved3(11) As Byte       ' ����
End Type

'------------------------------------------------------------------------
' �ڵ�090���������ߺ���
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_090
    timeout As Byte             ' �Ⲧ��ʱ����
    numbertype As Byte          ' ��������: 0 -- �̶��趨��1 -- ������2 -- �Ⲧ��Ϣ��3 -- COM
    dialtype As Byte            ' ��������: 0 -- ���������ţ�1 -- CTI����
    connecttype As Byte         ' ���Ž�ͨ�б�: 0 -- �������б�1 -- CTI��Ϣ�б�
    trytimes As Byte            ' ���Դ���
    log As Byte                 ' ��������־
    extdelay As Byte            ' ���ֻ�ǰ��ʱ����
    usevar As Byte              ' ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    resultvar As Byte           ' �����¼����ID: 0 --��ʹ�ã�1 -- 255 ����ID
    resultinform As Byte        ' ����Ƿ�֪ͨ�ⲿϵͳ
    explictoffhook As Byte      ' �Ƿ�ǿ��ժ����sun added 2012-01-17
    reserved1(1) As Byte        ' ����
End Type
Public Node90_Data1 As SData1_090
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_090
    predial(13) As Byte         ' ����ǰ׺
    phoneno(31) As Byte         ' ���к��룬��������Ϊ0ʱ��Ч
    com_iid As Integer          ' COM�ӿ�ID
    nd_parent As Integer        ' ���ڵ�ID
    nd_succ As Integer          ' �ɹ�ת�ڵ�ID
    nd_fail As Integer          ' ʧ��ת�ڵ�ID
    setANI As Integer           ' ��ʽ�������к��루��ԴID��
    reserved1(7) As Byte        ' ����
End Type
Public Node90_Data2 As SData2_090

'------------------------------------------------------------------------
' �ڵ�091��Calling Card
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_091
    timeout As Byte             ' �ڵ㳬ʱ����
    talklentype As Byte         ' ��ͨ��ʱ��������ʽ: 0 -- ����Ҫ������1 -- ʱ���룻2 -- ���룻3 -- �֣��������룩��4 -- ��
    obgroup As Byte             ' �Ⲧ���
    remindminute As Byte        ' ͨ��ʱ��ʣ����ٷ�������
    reserved1 As Byte           ' ����
    log As Byte                 ' ��������־
    reserved2 As Byte           ' ����
    var_cardno As Byte          ' ����ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_telno As Byte           ' �Ⲧ����ʹ�ñ���ID: 0 --��ʹ�ã�1 -- 255 ����ID
    var_connectlength As Byte   ' ͨ��ʱ����¼����ID: 0 --����¼��1 -- 255 ����ID
    reserved3(2) As Byte        ' ����
End Type
Public Node91_Data1 As SData1_091
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_091
    vox_talklen As Integer      ' ������ͨ��ʱ��ǰ��������ϵͳ�Զ�����talklentype������ʱ�䵥λʹ��ϵͳ������Դ��
    vox_timeout As Integer      ' ͨ����ʱ��������
    vox_noservice As Integer    ' �������Ч��ʾ����
    reserved1(37) As Byte       ' ����
    com_talklength As Integer   ' Ԥ�ȼ����ͨ��ʱ��COM�ӿ�ID������ʱ���룩��������ʽ����0ʱ��Ч
    com_billing As Integer      ' �Ʒ�COM�ӿ�ID�����ر���ͨ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved2(11) As Byte       ' ����
End Type
Public Node91_Data2 As SData2_091

'------------------------------------------------------------------------
' �ڵ�096���첽ͨ�Žڵ�
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_096
    timeout As Byte             ' �ȴ���ʱ����
    seperator As Byte           ' �ָ���
    extdata As Byte             ' ��չ���ݴ���ʽ: 0 -- ������1 -- �����ļ���2 -- ��¼��������3 -- ��¼��AppData��4 -- ��¼��UserData��5 -- TTS��
    extvar As Byte              ' extdata��¼����
    reserved1 As Byte           ' ����
    log As Byte                 ' ��������־
    carryonasynplay As Byte     ' �Ƿ��յ�ͨ����Ӧ�Ƿ�����첽������Ĭ��Ϊ0-������
    reserved2(5) As Byte        ' ����
End Type
Public Node96_Data1 As SData1_096
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_096
    command As Integer          ' �������
    vox_wt  As Integer          ' �ȴ�ѭ����������
    var_send(9) As Byte         ' ���ͱ���ID
    var_receive(9) As Byte      ' ���ձ���ID
    fileprefix(1) As Byte       ' �ļ���ǰ׺
    reserved1(21) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    nd_timeout As Integer       ' ��ʱ�ڵ�ID
    reserved2(9) As Byte        ' ����
End Type
Public Node96_Data2 As SData2_096

'------------------------------------------------------------------------
' �ڵ�100���û�DLL
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_100
    reserved1(4) As Byte        ' ����
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node100_Data1 As SData1_100
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_100
    dll_fid As Integer          ' DLL�ļ�ID
    reserved1(45) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved2(11) As Byte       ' ����
End Type
Public Node100_Data2 As SData2_100

'------------------------------------------------------------------------
' �ڵ�101���û�COM
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_101
    reserved1(4) As Byte        ' ����
    log As Byte                 ' ��������־
    reserved2(6) As Byte        ' ����
End Type
Public Node101_Data1 As SData1_101
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_101
    reserved1(15) As Byte       ' ����
    com_iid As Integer          ' COM�ӿ�ID
    reserved2(29) As Byte       ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved3(11) As Byte       ' ����
End Type
Public Node101_Data2 As SData2_101

'------------------------------------------------------------------------
' �ڵ�102����¼����
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_102
    reserved1(4) As Byte        ' ����
    log As Byte                 ' ��������־
    var_chg As Byte             ' ��¼����ID: 0 --����¼��1 -- 255 ����ID
    convert As Byte             ' ת����ʽ
    param1 As Byte              ' ����1
    param2 As Byte              ' ����2
    reserved2(2) As Byte        ' ����
End Type
Public Node102_Data1 As SData1_102
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_102
    reserved1(15) As Byte       ' ����
    com_iid As Integer          ' COM�ӿ�ID
    value(23) As Byte           ' ����ֵ
    reserved2(5) As Byte        ' ����
    nd_parent As Integer        ' ���ڵ�ID
    nd_child As Integer         ' �ӽڵ�ID
    reserved3(11) As Byte       ' ����
End Type
Public Node102_Data2 As SData2_102

'------------------------------------------------------------------------
' Sun added 2001-09-28
' �ڵ�255 - �ڵ�����
'/ �ڵ����ݽṹ-Data1(03-0F/13b)
Public Type SData1_255
    reserved1(12) As Byte       ' ����
End Type
Public Node255_Data1 As SData1_255
'/ �ڵ����ݽṹ-Data2(30-6F/64b)
Public Type SData2_255
    StartNode As Integer        ' ��ʼ�ڵ�ID
    EndNode As Integer          ' ��ֹ�ڵ�ID
    Index As Byte               ' 0��Ψһ���ߣ�����������������
    Style As Byte               ' ��������
    Width As Byte               ' ���ߴ�ϸ
    Color As Long               ' ������ɫ
    reserved1(52) As Integer    ' ����
End Type
Public Node255_Data2 As SData2_255

'------------------------------------------------------------------------
' Sun added 2002-05-23
' Ver 2 Mend 1
'' Exchange File Header Definition
'' 512 Bytes
Public Type SCFFileHeader
    P_ID_Name       As String * 10  ' �ָ���                  (10)
    P_ID            As Byte         ' ��ĿID                    (1)
    P_Type_Name     As String * 10  ' �ָ���                  (10)
    P_Type          As Byte         ' ��Ŀ����                  (1)
    P_Name_Name     As String * 10  ' �ָ���                  (10)
    P_Name          As String * 50  ' ��Ŀ����                  (50)
    P_Description_Name As String * 10  ' �ָ���               (10)
    P_Description   As String * 200 ' ��Ŀ��������              (200)
    P_Version_Name  As String * 10  ' �ָ���                  (10)
    P_Version       As String * 10  ' ��Ŀ�汾��                (10)
    P_Auther_Name   As String * 10  ' �ָ���                  (10)
    P_Auther        As String * 50  ' ��Ŀ����                  (50)
    P_User_Name     As String * 10  ' �ָ���                  (10)
    P_User          As String * 50  ' ��Ŀ�û�                  (50)
    P_CreateTime_Name As String * 10  ' �ָ���                (10)
    P_CreateTime    As Date         ' ��Ŀ����ʱ��              (8)
    P_ModifyTime_Name As String * 10  ' �ָ���                (10)
    P_ModifyTime    As Date         ' ��Ŀ����޸�ʱ��          (8)
    P_RCN_Name      As String * 10  ' �ָ���                  (10)
    P_RCN           As Long         ' ��¼��                    (4)
    P_Reserved_Name As String * 10  ' �ָ���                  (10)
    Reserved        As String * 20  ' ��䱣��                  (20)
End Type
'' Exchange File Record Definition
'' 128 Bytes
Public Type SCFFileRecord
    N_ID            As Integer      ' �ڵ�ID                    (2)
    N_NO            As Byte         ' �ڵ���                  (1)
    N_Description   As String * 32  ' �ڵ���������              (32)
    N_Page          As Byte         ' �ڵ�����ҳ��              (1)
    N_Left          As Integer      ' �ڵ�������                (2)
    N_Top           As Integer      ' �ڵ�������                (2)
    N_Height        As Integer      ' �ڵ�߶�                  (2)
    N_Width         As Integer      ' �ڵ���                  (2)
    N_Data1(12)     As Byte         ' �ڵ����ݶ�1               (13)
    N_Data2(63)     As Byte         ' �ڵ����ݶ�2               (64)
    '' Sun modified 2008-01-18
    ''' From
    'Reserved        As String * 7   ' ����                      (7)
    ''' To
    N_Tag           As String * 7   ' Tag                      (7)
End Type

Public Sub LoadNodeTypeNameList()

    gNodeTypeNameList(0) = LoadNationalResString(1490) '
    gNodeTypeNameList(1) = LoadNationalResString(1491) ' "1-�û�Buffer������־"
    gNodeTypeNameList(2) = LoadNationalResString(1492) ' "2-�û�Buffer�������"
    gNodeTypeNameList(6) = LoadNationalResString(1493) ' "6-������ת��"
    gNodeTypeNameList(7) = LoadNationalResString(1494) ' "7-�����֤"
    gNodeTypeNameList(8) = LoadNationalResString(1495) ' "8-�޸Ŀ���"
    gNodeTypeNameList(9) = LoadNationalResString(1496) ' "9-ʱ���֧"
    gNodeTypeNameList(10) = LoadNationalResString(1497) ' "10-�������趨"
    gNodeTypeNameList(16) = LoadNationalResString(1529) ' "16-������֧"
    gNodeTypeNameList(17) = LoadNationalResString(1498) ' "17-ѡ���������"
    gNodeTypeNameList(18) = LoadNationalResString(1499) ' "18-��������"
    gNodeTypeNameList(19) = LoadNationalResString(1500) ' "19-�޲���"
    gNodeTypeNameList(20) = LoadNationalResString(1501) ' "20-�����һ�"
    gNodeTypeNameList(21) = LoadNationalResString(1502) ' "21-��������"
    gNodeTypeNameList(22) = LoadNationalResString(1503) ' "22-�����ȴ�����"
    gNodeTypeNameList(23) = LoadNationalResString(1504) ' "23-����ת��"
    gNodeTypeNameList(28) = LoadNationalResString(1525) ' "28-TTS ����"
    gNodeTypeNameList(40) = LoadNationalResString(1505) ' "40-��������"
    gNodeTypeNameList(41) = LoadNationalResString(1506) ' "41-�쿴����"
    gNodeTypeNameList(50) = LoadNationalResString(1507) ' "50-�䵥����"
    gNodeTypeNameList(51) = LoadNationalResString(1508) ' "51-TTF����"
    gNodeTypeNameList(60) = LoadNationalResString(1509) ' "60-ת����ϯ"
    gNodeTypeNameList(61) = LoadNationalResString(1510) ' "61-ת����ϯ��"
    gNodeTypeNameList(62) = LoadNationalResString(1511) ' "62-��ǿת����ϯ"
    gNodeTypeNameList(63) = LoadNationalResString(1512) ' "63-��ǿת����ϯ��"
    gNodeTypeNameList(69) = LoadNationalResString(1513) ' "69-ת����ֻ�"
    gNodeTypeNameList(70) = LoadNationalResString(1600) ' "70-��ѯ·�ɵ�"
    gNodeTypeNameList(71) = LoadNationalResString(1633) ' "71-��ѯ��ϯ״̬"
    gNodeTypeNameList(80) = LoadNationalResString(1514) ' "80-����ACD"
    gNodeTypeNameList(90) = LoadNationalResString(1515) ' "90-�������ߺ���"
    gNodeTypeNameList(91) = LoadNationalResString(1583) ' "91-Calling Card"
    gNodeTypeNameList(96) = LoadNationalResString(1569) ' "96-�첽ͨ��"
    gNodeTypeNameList(100) = LoadNationalResString(1516) ' "100-�û�DLL"
    gNodeTypeNameList(101) = LoadNationalResString(1517) ' "101-�û�COM"
    gNodeTypeNameList(102) = LoadNationalResString(1518) ' "102-��¼����"
    gNodeTypeNameList(255) = LoadNationalResString(1519) ' "255-�ڵ�����"
        
End Sub

Public Function GetNodePictureFileName(ByVal f_NodeNo As Byte) As String

    GetNodePictureFileName = App.path & "\Bitmaps\" & Right("000" & Trim(Str(f_NodeNo)), 3) & ".ico"
    
End Function
