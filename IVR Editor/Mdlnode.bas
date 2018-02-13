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
' 操作定义
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
' 节点000：系统节点
Public Type SData1_000
    reserved1(9) As Byte        ' 保留
    Languages As Byte           ' 语言数量
    MajorVer As Byte            ' 主版本
    MinorVer As Byte            ' 次版本
End Type
Public Node0_Data1 As SData1_000
' 节点数据结构-Data2(30-6F/64b)
Public Type SData2_000
    key_repeat As Byte          ' 重复当前节点按键
    key_return As Byte          ' 回到上一级节点按键
    key_root As Byte            ' 回到主菜单按键
    
    'reserved1(44) As Byte      ' 保留 sun 2002-12-03 Old
    reserved1 As Byte           ' 保留 sun 2002-12-03
    ResourceProject As Integer  ' 资源项目ID sun 2002-12-03
    MainCOM As Integer          ' 流程COM组件资源编号 sun 2002-12-03
    LogSwitchOff As Byte        ' 节点日志全局控制开关, sun added 2008-02-19
    reserved2(38) As Byte       ' 保留 sun 2002-12-03
    
    nd_parent As Integer        ' 父节点ID
    nd_root As Integer          ' 主菜单(根)ID
    nd_SysSendData As Integer   ' 系统缺省发送数据格式定义节点 sun 2004-12-30
    nd_BeforeHookOn As Integer  ' 挂机前转节点，Sun added 2012-05-07
    reserved3(7) As Byte        ' 保留
End Type
Public Node0_Data2 As SData2_000
Public Default As SData2_000

'------------------------------------------------------------------------
' 节点001：变量总数
' 节点数据结构-Data1(03-0F/13b)
Public Type SData1_001
     reserved1(10) As Byte      ' 保留
     uservars As Byte           ' 用户定义变量数, 0-255
     reserved2 As Byte
End Type
Public Node1_Data1 As SData1_001
Public Defaultuservar As SData1_001

'------------------------------------------------------------------------
' 节点002：变量清单
' 节点数据结构-Data1(03-0F/13b)
Public Type SData1_002
    reserved1(12) As Byte
End Type
Public Node2_Data1 As SData1_002
' 用户定义变量结构
Public Type SData2_002
    uservar(63) As Byte
End Type
Public Node2_Data2 As SData2_002
Public Defaulttype As SData2_002

'------------------------------------------------------------------------
' 节点006：无条件转移
' 节点数据结构-Data1(03-0F/13b)
Public Type SData1_006
    Sleep As Integer            ' 延时（sleep），单位为毫秒
    reserved1(10) As Byte       ' 保留
End Type
Public Node6_Data1 As SData1_006

'节点数据结构-Data2(30-6F/64b)
Public Type SData2_006
    nd_goto As Integer          ' 跳转节点ID
    reserved1(46) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    reserved2(14) As Byte       ' 保留
End Type
Public Node6_Data2 As SData2_006

'------------------------------------------------------------------------
' 节点007：身份验证
' 节点数据结构-Data1(03-0F/13b)
Public Type SData1_007
    timeout As Byte             ' 节点超时(秒)
    maxuserid As Byte           ' 用户号码最大长度
    maxpassword As Byte         ' 用户口令最大长度
    key_term As Byte            ' 输入终止符, 按键
    maxtrytime As Byte          ' 最大尝试次数
    log As Byte                 ' 被访问日志
    var_trytime As Byte         ' 验证次数记录(0 - maxtrytime)
    var_result As Byte          ' 验证结果记录
    var_userid As Byte          ' 用户号码记录
    var_password As Byte        ' 用户口令记录
    reserved1(2) As Byte        ' 保留
End Type
Public Node7_Data1 As SData1_007
' 节点数据结构-Data2(30-6F/64b)
Public Type SData2_007
    vox_userid As Integer       ' 提示用户输入号码播放语音
    vox_password As Integer     ' 提示用户输入口令播放语音
    vox_tryagain As Integer     ' 提示用户重新输入播放语音
    reserved1(9) As Byte        ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_succeed As Integer       ' 成功转节点ID
    nd_fail As Integer          ' 失败转节点ID
    reserved3(9) As Byte        ' 保留
End Type
Public Node7_Data2 As SData2_007

'------------------------------------------------------------------------
' 节点008：修改口令
' 节点数据结构-Data1(03-0F/13b)
Public Type SData1_008
    timeout As Byte             ' 节点超时(秒)
    reserved1 As Byte           ' 保留
    maxpassword As Byte         ' 用户口令最大长度
    key_term As Byte            ' 输入终止符, 按键
    maxtrytime As Byte          ' 最大尝试次数
    log  As Byte                ' 被访问日志
    var_trytime As Byte         ' 尝试次数记录(0 - maxtrytime)
    var_result As Byte          ' 验证结果记录
    reserved2 As Byte           ' 保留
    var_password As Byte        ' 用户口令记录
    reserved3(2) As Byte        ' 保留
End Type
Public Node8_Data1 As SData1_008
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_008
    vox_password As Integer     ' 提示用户输入新口令播放语音
    vox_confirm As Integer      ' 提示用户再次确认播放语音
    vox_tryagain As Integer     ' 两次不一致重新输入播放语音
    vox_succeed As Integer      ' 提示用户修改成功播放语音
    reserved1(7) As Byte        ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_succeed As Integer       ' 成功转节点ID
    nd_fail As Integer          ' 失败转节点ID
    reserved3(9) As Byte        ' 保留
End Type
Public Node8_Data2 As SData2_008

'------------------------------------------------------------------------
' 节点009：时间分支
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_009
    reserved1(4) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node9_Data1 As SData1_009
'/ 时间段结构
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_009
    workday As Byte             ' 工作日安排, Bit操作, 0-6位有效,0:休息日;1:工作日
    worktime As Byte            ' 工作日时间段安排, Bit操作, 0-5位有效,0:无效;1:有效
    timesec(23) As Byte         ' 时间段1-6
    reserved1(21) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_sparetime As Integer     ' 休息日转节点ID
    nd_timesec(5) As Integer    ' 时间段1-6转节点ID
End Type
Public Node9_Data2 As SData2_009

'------------------------------------------------------------------------
' 节点010：工作日设定
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_010
    maincalendar As Byte        ' 是否是主日历，不可见，系统识别
    startyear As Byte           ' 开始年份，YY
    startmonth As Byte          ' 开始月，1-12
    monthcount As Byte          ' 共几个月（最大12个月）
    reserved1 As Byte           ' 保留
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node10_Data1 As SData1_010
'/ 时间段结构
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_010
    daytype(47) As Byte         ' 位操作: 0 - 工作日；1 - 节假日
    nd_parent As Integer        ' 父节点ID
    nd_daysec(2) As Integer     ' 转节点ID
    reserved1(7) As Byte        ' 保留
End Type
Public Node10_Data2 As SData2_010

'------------------------------------------------------------------------
' Sun added 2004-12-30
' 节点016：条件分支
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_016
    reserved1(4) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    var_id As Byte              ' 变量
    logic As Byte               ' 逻辑运算符
    convert As Byte             ' 转换公式
    param1 As Byte              ' 参数1
    param2 As Byte              ' 参数2
    reserved2(1) As Byte        ' 保留
End Type
Public Node16_Data1 As SData1_016
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_016
    var_value(47) As Byte       ' 变量值
    nd_parent As Integer        ' 父节点ID
    nd_succ As Integer          ' 条件满足转移节点ID
    nd_fail As Integer          ' 条件不满足转移节点ID
    reserved3(9) As Byte        ' 保留
End Type
Public Node16_Data2 As SData2_016

''/ 逻辑运算符定义
Public Const DEF_NODE016_LOGIC_EQUE = 0                   ' =
Public Const DEF_NODE016_LOGIC_BIGGER = 1                 ' >
Public Const DEF_NODE016_LOGIC_SMALLER = 2                ' <
Public Const DEF_NODE016_LOGIC_EB = 3                     ' >=
Public Const DEF_NODE016_LOGIC_ES = 4                     ' <=
Public Const DEF_NODE016_LOGIC_NE = 5                     ' <>
Public Const DEF_NODE016_LOGIC_LIKE = 6                   ' Like

''/ 转换公式定义
Public Const DEF_NODE016_CONVERT_NONE = 0                 ' Do Nothing
Public Const DEF_NODE016_CONVERT_LEFT = 1                 ' Left(n)
Public Const DEF_NODE016_CONVERT_RIGHT = 2                ' Right(n)
Public Const DEF_NODE016_CONVERT_MID = 3                  ' Mid(m, n)
Public Const DEF_NODE016_CONVERT_LEN = 4                  ' Len()

'------------------------------------------------------------------------
' Sun added 2003-04-24
' 节点017：服务语言选择
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_017
    timeout As Byte             ' 节点超时(秒)
    maxtrytime As Byte          ' 最大尝试次数
    var_lang As Byte            ' 选择结果记录
    reserved1(9) As Byte        ' 保留
End Type
Public Node17_Data1 As SData1_017
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_017
    vox_play As Integer         ' 播放语音
    lang(9) As Byte             ' 语言对应按键
    reserved1(35)  As Byte      ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_succ As Integer          ' 成功节点ID
    nd_fail As Integer          ' 失败节点ID
    reserved2(9) As Byte        ' 保留
End Type
Public Node17_Data2 As SData2_017

'------------------------------------------------------------------------
' 节点018：发送数据
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_018
    seperator As Byte           ' 分隔符
    reserved1(11) As Byte       ' 保留
End Type
Public Node18_Data1 As SData1_018
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_018
    typeflags(14) As Byte       ' 标记：AppData/UserData
    prefix1(14) As Byte         ' 变量前缀1
    prefix2(14) As Byte         ' 变量前缀2
    valueid(14) As Byte         ' 变量ID
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
End Type
Public Node18_Data2 As SData2_018

'------------------------------------------------------------------------
' 节点019：无操作
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_019
    'Michael Modified @Jul,10,07
    'reserved1(12) As Byte       ' 保留
    reserved1(11) As Byte       ' 保留
    'Michael Added @ the same day
    leavequeue    As Byte
End Type
Public Node19_Data1 As SData1_019
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_019
    delaytime   As Integer      ' 延时时间(ms)
    reserved1(45) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    reserved2(13) As Byte       ' 保留
End Type
Public Node19_Data2 As SData2_019

'------------------------------------------------------------------------
' 节点020：放音挂机
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_020
    reserved1 As Byte           ' 保留
    playclear As Byte           ' 放音清空标志: 0 -- 不清空;1 -- 清空
    reserved2(2) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    reserved3(6) As Byte        ' 保留
End Type
Public Node20_Data1 As SData1_020
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_020
    vox_play As Integer         ' 播放语音
    reserved1(45) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    reserved2(13) As Byte       ' 保留
End Type
Public Node20_Data2 As SData2_020

'------------------------------------------------------------------------
' 节点021：放音继续
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_021
    reserved1 As Byte           ' 保留
    playclear As Byte           ' 放音清空标志: 0 -- 不清空;1 -- 清空
    playtype As Byte            ' 播放类型(可组合):0数字1金额2号码4字母8汉字16日期
    usevar As Byte              ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
    breakkey As Byte            ' 中断按键
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node21_Data1 As SData1_021
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_021
    vox_pred As Integer         ' 前导播放语音
    vox_succ As Integer         ' 后续播放语音
    reserved1(11) As Byte       ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved3(11) As Byte       ' 保留
End Type
Public Node21_Data2 As SData2_021

'---------------------------------------------------------------------------------
' 节点022：放音等待按键
'' 节点数据结构-Data1(03-0F/13b)
Public Type SData1_022
    timeout As Byte             ' 节点超时(秒)
    playclear As Byte           ' 放音清空标志: 0 -- 不清空;1 -- 清空
    getlength As Byte           ' 用户输入长度
    maxinterval As Byte         ' 按键最大间隔(秒)
    breakkey As Byte            ' 中断按键
    log As Byte                 ' 被访问日志
    var_key As Byte             ' 用户按键记录, 1-255变量ID
    maxtrytime As Byte          ' 最大尝试次数
    reserved1(4) As Byte        ' 保留
End Type
Public Node22_Data1 As SData1_022
' 节点数据结构-Data2(30-6F/64b)
Public Type SData2_022
    vox_play As Integer         ' 播放语音
    vox_nodefail As Integer     ' 节点失败播放语音[V2]
    reserved1(27) As Byte       ' 保留(27)
    nd_parent As Integer        ' 父节点ID
    nd_key(11) As Integer       ' 子节点ID: 0 - 11
    nd_nodefail As Integer      ' 子节点ID: 失败转入节点[V2]
    reserved2(3) As Byte        ' 保留(3)
End Type
Public Node22_Data2 As SData2_022

'------------------------------------------------------------------------
' 节点023：放音转移
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_023
    timeout As Byte             ' 节点超时(秒)
    playclear As Byte           ' 放音清空标志: 0 -- 不清空;1 -- 清空
    breakkey As Byte            ' 中断按键
    reserved1(1) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    var_play As Byte            ' 放音, 1-255变量ID
    reserved2(5) As Byte        ' 保留
End Type
Public Node23_Data1 As SData1_023
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_023
    vox_play As Integer         ' 播放语音，var_play = 0时有效
    reserved1(45) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_goto As Integer          ' 转移节点ID
    reserved2(11) As Byte       ' 保留
End Type
Public Node23_Data2 As SData2_023

'------------------------------------------------------------------------
' Sun added 2004-12-30
' 节点028：TTS放音
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_028
    timeout As Byte             ' 节点超时(秒)
    playclear As Byte           ' 放音清空标志: 0 -- 不清空;1 -- 清空
    usevar As Byte              ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
    breakkey As Byte            ' 中断按键
    reserved1 As Byte           ' 保留
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node28_Data1 As SData1_028
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_028
    vox_string As Integer       ' 播放字符串，支持"%s"插入子串
    vox_alter As Integer        ' TTS不可用时替代语音
    playtype As Byte            ' 播放类型：
                                '' 0 - 字符串
                                '' 1 - 文本文件（vox_string格式化后作为文件路径）
                                '' 2 - ExtData
                                '' 3 - UserData
    reserved1(10) As Byte       ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_succ As Integer          ' TTS播放成功转移节点ID
    nd_fail As Integer          ' TTS播放失败转移节点ID
    reserved3(9) As Byte        ' 保留
End Type
Public Node28_Data2 As SData2_028

'------------------------------------------------------------------------
' 节点040：建立留言
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_040
    rectime As Byte             ' 录音时长(秒)
    playclear As Byte           ' 放音清空标志: 0 -- 不清空;1 -- 清空
    breakkey As Byte            ' 中断按键
    var_agent As Byte           ' Agent[V2]
    var_filename As Byte        ' 录音文件名记录变量
    log As Byte                 ' 被访问日志
    var_appfield(2) As Byte     ' 应用数据项
    maxsilencetime As Byte      ' 最大静音时长(秒)
    vmsclass As Byte            ' 留言分类[2007-02-28]
    MinRecLength As Byte        ' 最短录音时长(秒)[2007-03-20]
    toneoff As Byte             ' 不播放录音开始提示[2009-07-24]，0 - 播放；1 - 不播放
End Type
Public Node40_Data1 As SData1_040
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_040
    vox_op As Integer           ' 录音前提示语音
    'new add
    recfiletype As Byte         ' 录音文件类型[2007-06-28]: 0 - vox; 1 - wav
    rectime_ho As Byte          ' 录音时长高8位
    var_notifyintvl As Byte     ' 提示间隔(秒)[2009-07-24]，0 - 不提示
    var_rectime As Byte         ' 录音实际时长记录变量(秒)[2009-07-24]
    NotifyPL As Byte            ' 是否通知PL录音系统[2012-04-18]
    reserved1(40) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved2(11) As Byte       ' 保留
End Type
Public Node40_Data2 As SData2_040

''/ 留言分类定义
Public Const DEF_NODE040_VMSCLASS_UNKNOWN = 0               ' 未知
Public Const DEF_NODE040_VMSCLASS_ALL = 1                   ' 公共
Public Const DEF_NODE040_VMSCLASS_GROUP = 2                 ' 组
Public Const DEF_NODE040_VMSCLASS_EXT = 3                   ' 分机
Public Const DEF_NODE040_VMSCLASS_USER = 4                  ' 用户

'------------------------------------------------------------------------
' 节点041：察看留言
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_041
    timeout As Byte             ' 节点超时(秒)
    playclear As Byte           ' 放音清空标志: 0 -- 不清空;1 -- 清空
    breakkey As Byte            ' 中断按键
    var_agent As Byte           ' Agent[V2]
    reserved1 As Byte           ' 保留
    log As Byte                 ' 被访问日志
    vmstype As Byte             ' 察看留言类型
    closewhencheck As Byte      ' 察看后自动变为关闭状态
    vmsclass As Byte            ' 留言分类[2007-02-28]
    reserved2(3) As Byte        ' 保留
 End Type
 Public Node41_Data1 As SData1_041
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_041
    vox_play(3) As Integer      ' 留言状态报告前导/留言状态报告后续/浏览提示语音/操作提示语音
    key_op(7) As Byte           ' 听第一条留言按键/前一条留言按键/听下一条留言按键/听最后一条留言按键/重听当前留言按键/删除当前留言按键/退出本节点按键/转换类型按键
    reserved1(31) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved2(11) As Byte       ' 保留
End Type
Public Node41_Data2 As SData2_041

''/ 察看留言类型定义
Public Const DEF_NODE041_VMSTYPE_NEW = 0                  ' 新留言，可删除、存档
Public Const DEF_NODE041_VMSTYPE_CLOSED = 1               ' 存档留言，可删除
Public Const DEF_NODE041_VMSTYPE_DELETED = 2              ' 删除留言，可取消删除变为存档

'------------------------------------------------------------------------
' Sun added 2006-12-31, V6.5.11
'' 构造传真文件名
Public Const DEF_NODE050_FAXN_TYPE_AUTO = 0               ' 自动生成
Public Const DEF_NODE050_FAXN_TYPE_RESID = 1              ' 资源ID
Public Const DEF_NODE050_FAXN_TYPE_VAR2RESID = 2          ' 变量对应资源ID
Public Const DEF_NODE050_FAXN_TYPE_VAR2NAME = 3           ' 变量对应文件名
Public Const DEF_NODE050_FAXN_TYPE_FORMAT = 4             ' 变量替换资源中的通配符

' Sun updated 2006-12-31, V6.5.11
' 节点050：简单传真
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_050
    timeout As Integer          ' 节点超时(秒)
    filenametype As Byte        ' 传真文件名类型:
                                '' 1 -- 资源ID
                                '' 2 -- 变量对应资源ID
                                '' 3 -- 变量对应文件名
                                '' 4 -- 变量替换资源中的通配符
                                '' 一次多个发送多个文件用分号分割，但最多5个
    trytimes As Byte            ' 尝试次数
    record_cdr As Byte          ' 是否记录传真发送详单: 0 --不记录；1 -- 记录
    log As Byte                 ' 被访问日志
    var_faxfile  As Byte        ' 传真文件使用变量ID: 0 --不使用；1 -- 255 变量ID
    var_fromno As Byte          ' 发出号码使用变量ID: 0 --不使用；1 -- 255 变量ID
    var_tono As Byte            ' 接收号码使用变量ID: 0 --不使用；1 -- 255 变量ID
    var_result As Byte          ' 结果记录变量ID: 0 --不使用；1 -- 255 变量ID
    var_appfield(2) As Byte     ' 应用数据项
End Type
Public Node50_Data1 As SData1_050
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_050
    vox_op As Integer           ' 操作提示语音
    fax_fileid As Integer       ' 传真文件(资源ID)
    header_id As Integer        ' 传真标题(资源ID)
    reserved1(41) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_succ As Integer          ' 成功转移节点ID
    nd_fail As Integer          ' 失败转移节点ID
    reserved2(9) As Byte        ' 保留
End Type
Public Node50_Data2 As SData2_050

'------------------------------------------------------------------------
' 节点051：TTF传真
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_051
    timeout As Byte             ' 节点超时(秒)
    reserved1(3) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node51_Data1 As SData1_051
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_051
    vox_op As Integer           ' 操作提示语音
    fax_logo As Integer         ' LOGO文件(资源ID)
    fax_format As Integer       ' 表格格式文件(资源ID)
    header_id As Integer        ' 传真标题(资源ID)
    from_id As Integer          ' 发出号码(资源ID)
    reserved1(5) As Byte        ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved3(11) As Byte       ' 保留
End Type
Public Node51_Data2 As SData2_051


'------------------------------------------------------------------------
' Sun added 2006-12-31, V6.5.11
' 节点055：传真接收
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_055
    timeout As Integer          ' 节点超时(秒)
    filenametype As Byte        ' 传真文件名类型:
                                '' 0 -- 自动生成（$Record\Group%n\%Y%m%d%H%M%S_c<CH>.tif）
                                '' 1 -- 资源ID
                                '' 2 -- 变量对应资源ID
                                '' 3 -- 变量对应文件名
                                '' 4 -- 变量替换资源中的通配符
    var_faxfile As Byte         ' 传真文件使用变量ID: 0 --不使用；1 -- 255 变量ID
    record_cdr As Byte          ' 是否记录传真发送详单: 0 --不记录；1 -- 记录
    log As Byte                 ' 被访问日志
    var_fromno As Byte          ' 发出号码使用变量ID: 0 --不使用；1 -- 255 变量ID
    var_tono As Byte            ' 接收号码使用变量ID: 0 --不使用；1 -- 255 变量ID
    var_extno As Byte           ' 分机号码记录变量ID: 0 --不使用；1 -- 255 变量ID
    var_result As Byte          ' 结果记录变量ID: 0 --不使用；1 -- 255 变量ID
    var_appfield(2) As Byte     ' 应用数据项
End Type
Public Node55_Data1 As SData1_055
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_055
    vox_op As Integer           ' 操作提示语音
    fax_fileid As Integer       ' 传真文件(资源ID)
    reserved1(43) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_succ As Integer          ' 成功转移节点ID
    nd_fail As Integer          ' 失败转移节点ID
    reserved2(9) As Byte        ' 保留
End Type
Public Node55_Data2 As SData2_055

'------------------------------------------------------------------------
' 节点060：转接座席
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_060
    timeout As Byte             ' 节点超时(秒)
    switchtype As Byte          ' 转接方式(0：自动传接；1：指定座席；2：用户输入)
    agentid As Byte             ' 指定座席ID, 方式1时 / 中断按键, 方式2时
    getlength As Byte           ' 用户输入长度, 转接方式为2时有效, 转接方式为1时为指定座席ID的高位
    looptimes As Byte           ' 等待循环播放次数
    log As Byte                 ' 被访问日志
    var_key As Byte             ' 用户按键记录, 转接方式为2时有效，1-255变量ID
    agentinfo As Byte           ' 是否宣读座席信息
    reserved1(4) As Byte        ' 保留
End Type
Public Node60_Data1 As SData1_060
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_060
    vox_op As Integer           ' 操作提示语音
    vox_sw As Integer           ' 转接提示语音
    vox_wt As Integer           ' 等待循环播放语音
    vox_nobody As Integer       ' 没有上班提示语音
    vox_busy As Integer         ' 座席忙提示语音
    vox_noanswer As Integer     ' 座席无应答提示语音
    vox_ok   As Integer         ' 座席转接成功
    'Mike Added @ 2008-5-27     # moved a Byte storage space from reserved1(33)
    length_agentinfo As Byte    ' 座席信息（如：工号）位数，默认4位，不足在前面补0，0采用原始长度
    reserved1(32) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_nobody As Integer        ' 没有上班转节点ID
    nd_busy As Integer          ' 座席忙转节点ID
    nd_noanswer As Integer      ' 座席无应答转节点ID
    nd_ok As Integer            ' 转接成功转节点ID
    reserved2(5) As Byte        ' 保留
End Type
Public Node60_Data2 As SData2_060

'------------------------------------------------------------------------
' 节点061：转接座席组
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_061
    maxwait      As Integer   ' 等待超时(秒)，0 表示无限等待
    toacd        As Byte      ' 转接ACD或是RoutePoint，0 － RoutePoint；1 － ACD
    usevar       As Byte      ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
    looptimes    As Byte      ' 等待循环播放次数，0 表示无限循环
    log          As Byte      ' 被访问日志
    agentinfo    As Byte      ' 是否宣读座席信息
    readEWT      As Byte      ' 是否播报预测等待时间, 0 - 不宣读;1 - 宣读(秒);2 - 宣读(时分秒)
    switchtype   As Byte      ' 转接方式: 0 - 拨号；1 - CTI
    waitmethod   As Byte      ' 等待方式: 0 - 播放语音；1 - 异步走流程
    var_userid   As Byte      ' 登录座席ID记录变量ID  : 0 --不记录;1 -- 255 变量ID
    var_loginid  As Byte      ' 登录座席工号记录变量ID: 0 --不记录;1 -- 255 变量ID
    waitansto    As Byte      ' 等待坐席应答超时(秒)，0 表示无限等待
End Type
Public Node61_Data1 As SData1_061
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_061
    vox_op        As Integer     ' 操作提示语音
    vox_sw        As Integer     ' 转接提示语音
    vox_wt        As Integer     ' 等待循环播放语音
    vox_nobody    As Integer     ' 没有上班提示语音
    vox_busy      As Integer     ' 座席组全忙提示语音
    vox_noanswer  As Integer     ' 座席无应答提示语音
    vox_ok        As Integer     ' 座席组转接成功
    routepointid  As Integer     ' 转接的路由点编号
    acddn(19)     As Byte        ' 转接的ACD号码
    length_agentinfo As Byte     ' 座席信息（如：工号）位数，默认4位，不足在前面补0，0采用原始长度
    reserved1(9) As Byte
    waitansto_hi  As Byte        '保存等待座席应答超时的高8位
    nd_parent     As Integer     ' 父节点ID
    nd_nobody     As Integer     ' 没有上班转节点ID
    nd_busy       As Integer     ' 座席组全忙转节点ID
    nd_noanswer   As Integer     ' 座席无应答转节点ID
    nd_ok         As Integer     ' 转接成功转节点ID
    nd_wait       As Integer     ' 等待流程入口节点号
    reserved2(3)  As Integer     ' 保留
End Type
Public Node61_Data2 As SData2_061

'------------------------------------------------------------------------
' 节点062：发起会议，2012-04-17
'' 节点062：增强转接座席
'/ 节点数据结构-Data1(03-0F/13b)
'Public Type SData1_062
'    reserved1(2) As Byte        ' 保留
'    usevar As Byte              ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
'    looptimes As Byte           ' 等待循环播放次数
'    log As Byte                 ' 被访问日志
'    reserved2(6) As Byte        ' 保留
'End Type
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_062
    timeout As Byte             ' 节点超时(秒)
    reserved1 As Byte           ' 保留
    usevar As Byte              ' 使用变量ID: 0 -- 不使用；1 -- 255 变量ID
    looptimes As Byte           ' 等待循环播放次数，0 表示使用系统默认次数
    switchtype As Byte          ' 转接方式: 0 - 拨号（仅TAPI环境有效）；1 - CTI
    log          As Byte        ' 被访问日志
    waitansto    As Byte        ' 等待应答超时(秒)，0 表示无限等待
    var_waitansto As Byte       ' 等待坐席应答超时变量ID: 0 --不使用；1 -- 255 变量ID, Sun added 2014-01-29
    reserved2(4) As Byte        ' 保留
End Type
Public Node62_Data1 As SData1_062
''/ 节点数据结构-Data2(30-6F/64b)
'Public Type SData2_062
'    vox_op As Integer           ' 操作提示语音
'    vox_sw As Integer           ' 转接提示语音
'    vox_wt As Integer           ' 等待循环播放语音
'    vox_nobody As Integer       ' 没有上班提示语音
'    vox_busy As Integer         ' 座席全忙提示语音
'    vox_ok As Integer           ' 座席转接成功
'    reserved1(3) As Byte        ' 保留
'    com_iid As Integer          ' COM接口ID
'    reserved2(29) As Byte       ' 保留
'    nd_parent As Integer        ' 父节点ID
'    nd_nobody As Integer        ' 没有上班转节点ID
'    nd_busy As Integer          ' 座席全忙转节点ID
'    nd_ok As Integer            ' 转接成功转节点ID
'    reserved3(7) As Byte        ' 保留
'End Type
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_062
    vox_op As Integer           ' 操作提示语音
    vox_sw As Integer           ' 会议磋商完成提示语音
    vox_wt As Integer           ' 等待循环播放语音
    vox_noconf As Integer       ' 磋商失败提示语音
    vox_noans As Integer        ' 无应答提示语音
    vox_ans As Integer          ' 成功应答提示语音
    vox_ok As Integer           ' 完成会议提示语音
    vox_syserror As Integer     ' 系统错误提示语音
    DialNo(23) As Byte          ' 会议号码，也可以是:变量名
    predial(5) As Byte          ' 拨号前缀
    reserved1(1) As Byte        ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_noans As Integer         ' 无应答转节点ID
    nd_failed As Integer        ' 会议失败转节点ID
    nd_ok As Integer            ' 会议成功转节点ID
    reserved2(7) As Byte        ' 保留
End Type
Public Node62_Data2 As SData2_062

'------------------------------------------------------------------------
' 节点063：增强转接座席组
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_063
    reserved1(2) As Byte        ' 保留
    usevar As Byte              ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
    looptimes As Byte           ' 等待循环播放次数
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node63_Data1 As SData1_063
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_063
    vox_op As Integer           ' 操作提示语音
    vox_sw As Integer           ' 转接提示语音
    vox_wt  As Integer          ' 等待循环播放语音
    vox_nobody As Integer       ' 没有上班提示语音
    vox_busy  As Integer        ' 座席全忙提示语音
    vox_ok   As Integer         ' 座席转接成功
    reserved1(35) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_nobody  As Integer       ' 没有上班转节点ID
    nd_busy As Integer          ' 座席全忙转节点ID
    nd_ok  As Integer           ' 转接成功转节点ID
    reserved2(7) As Byte        ' 保留
End Type
Public Node63_Data2 As SData2_063

'------------------------------------------------------------------------
' 节点069：转虚拟分机
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_069
    reserved1(2) As Byte        ' 保留
    usevar As Byte              ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
    switchtype As Byte          ' 转接方式: 0 - 拨号；1 - CTI
    log As Byte                 ' 被访问日志
    vagency As Long             ' 虚拟分机号码
    reserved3(2) As Byte        ' 保留
End Type
Public Node69_Data1 As SData1_069
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_069
    vox_op As Integer           ' 操作提示语音
    maxtry As Integer           ' 最大尝试次数，2012-04-18
    tryinterval As Integer      ' 尝试间隔，2012-04-18
    reserved1(41) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child  As Integer        ' 子节点ID
    reserved2(11) As Byte       ' 保留
End Type
Public Node69_Data2 As SData2_069

'------------------------------------------------------------------------
' 节点070：查询路由点
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_070
    timeout As Byte             ' 节点超时(秒)
    reserved1 As Byte           ' 保留
    usevar As Byte              ' 使用变量ID: 0 -- 不使用；1 -- 255 变量ID
    paramindex As Byte          ' 参数编号
    logic As Byte               ' 逻辑运算符
    log As Byte                 ' 被访问日志
    querytype As Byte           ' 查询类型，0 - 路由点；1 - 队列；2 - 组；3 - 小组
    var_result As Byte          ' 查询结果记录变量: 0 -- 不使用；1 -- 255 变量ID，[2007-04-12]
    reserved2(4) As Byte        ' 保留
End Type
Public Node70_Data1 As SData1_070
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_070
    routepointid As Integer     ' 路由点编号
    comparedvalue As Integer    ' 参考值
    reserved1(11) As Byte       ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_yes As Integer           ' 满足条件转节点ID
    nd_no As Integer            ' 不满足条件转节点ID
    nd_fail As Integer          ' 查询失败节点ID
    reserved3(7) As Byte        ' 保留
End Type
Public Node70_Data2 As SData2_070

'------------------------------------------------------------------------
' 节点071：查询座席状态
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_071
    timeout As Byte             ' 节点超时(秒)
    usevar As Byte              ' 使用变量ID: 0 -- 不使用；1 -- 255 变量ID
    dn_status As Byte           ' DN状态编号
    pos_status As Byte          ' POS状态编号
    dn_logic As Byte            ' DN逻辑运算符
    pos_logic As Byte           ' POS逻辑运算符
    log As Byte                 ' 被访问日志
    querytype As Byte           ' 查询类型，0 - Agent; 1 - User (必须使用变量确定用户名); 2 - 工号
    conditions As Byte          ' 条件，条件定义
    reserved1(3) As Byte        ' 保留
End Type
Public Node71_Data1 As SData1_071
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_071
    agentid As Long             ' 座席编号或工号
    reserved1(43) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_yes As Integer           ' 满足条件转节点ID
    nd_no As Integer            ' 不满足条件转节点ID
    nd_fail As Integer          ' 查询失败节点ID
    reserved2(7) As Byte        ' 保留
End Type
Public Node71_Data2 As SData2_071

''/ 条件定义
Public Const DEF_NODE071_CONDITION_NONE = 0
Public Const DEF_NODE071_CONDITION_FIRST = 1
Public Const DEF_NODE071_CONDITION_SECOND = 2
Public Const DEF_NODE071_CONDITION_BOTH = 3
Public Const DEF_NODE071_CONDITION_EITHER = 4

'------------------------------------------------------------------------
' 节点080：进入ACD排队
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_080
    reserved1(2) As Byte        ' 保留
    usevar As Byte              ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
    reserved2 As Byte           ' 保留
    log As Byte                 ' 被访问日志
    maxwaittime As Integer      ' 最长排队等待时间(秒)
    reserved3(4) As Byte        ' 保留
End Type
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_080
    vox_wait As Integer         ' 排队等待语音
    reserved1(13) As Byte       ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved3(11) As Byte       ' 保留
End Type

'------------------------------------------------------------------------
' 节点090：呼叫外线号码
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_090
    timeout As Byte             ' 外拨超时，秒
    numbertype As Byte          ' 号码类型: 0 -- 固定设定；1 -- 变量；2 -- 外拨消息；3 -- COM
    dialtype As Byte            ' 拨号类型: 0 -- 语音卡拨号；1 -- CTI拨号
    connecttype As Byte         ' 拨号接通判别: 0 -- 语音卡判别；1 -- CTI消息判别
    trytimes As Byte            ' 尝试次数
    log As Byte                 ' 被访问日志
    extdelay As Byte            ' 拨分机前延时，秒
    usevar As Byte              ' 使用变量ID: 0 --不使用；1 -- 255 变量ID
    resultvar As Byte           ' 结果记录变量ID: 0 --不使用；1 -- 255 变量ID
    resultinform As Byte        ' 结果是否通知外部系统
    explictoffhook As Byte      ' 是否强制摘机，sun added 2012-01-17
    reserved1(1) As Byte        ' 保留
End Type
Public Node90_Data1 As SData1_090
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_090
    predial(13) As Byte         ' 拨号前缀
    phoneno(31) As Byte         ' 被叫号码，号码类型为0时有效
    com_iid As Integer          ' COM接口ID
    nd_parent As Integer        ' 父节点ID
    nd_succ As Integer          ' 成功转节点ID
    nd_fail As Integer          ' 失败转节点ID
    setANI As Integer           ' 显式设置主叫号码（资源ID）
    reserved1(7) As Byte        ' 保留
End Type
Public Node90_Data2 As SData2_090

'------------------------------------------------------------------------
' 节点091：Calling Card
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_091
    timeout As Byte             ' 节点超时，秒
    talklentype As Byte         ' 可通话时长播报方式: 0 -- 不需要播报；1 -- 时分秒；2 -- 分秒；3 -- 分（四舍五入）；4 -- 秒
    obgroup As Byte             ' 外拨组号
    remindminute As Byte        ' 通话时间剩余多少分钟提醒
    reserved1 As Byte           ' 保留
    log As Byte                 ' 被访问日志
    reserved2 As Byte           ' 保留
    var_cardno As Byte          ' 卡号使用变量ID: 0 --不使用；1 -- 255 变量ID
    var_telno As Byte           ' 外拨号码使用变量ID: 0 --不使用；1 -- 255 变量ID
    var_connectlength As Byte   ' 通话时长记录变量ID: 0 --不记录；1 -- 255 变量ID
    reserved3(2) As Byte        ' 保留
End Type
Public Node91_Data1 As SData1_091
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_091
    vox_talklen As Integer      ' 播报可通话时长前导语音（系统自动根据talklentype播报，时间单位使用系统语音资源）
    vox_timeout As Integer      ' 通话超时提醒语音
    vox_noservice As Integer    ' 余额不足或卡无效提示语音
    reserved1(37) As Byte       ' 保留
    com_talklength As Integer   ' 预先计算可通话时长COM接口ID（返回时分秒），播报方式大于0时有效
    com_billing As Integer      ' 计费COM接口ID（返回本次通话金额）
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved2(11) As Byte       ' 保留
End Type
Public Node91_Data2 As SData2_091

'------------------------------------------------------------------------
' 节点096：异步通信节点
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_096
    timeout As Byte             ' 等待超时，秒
    seperator As Byte           ' 分隔符
    extdata As Byte             ' 扩展数据处理方式: 0 -- 不处理；1 -- 保存文件；2 -- 记录到变量；3 -- 记录到AppData；4 -- 记录到UserData；5 -- TTS；
    extvar As Byte              ' extdata记录变量
    reserved1 As Byte           ' 保留
    log As Byte                 ' 被访问日志
    carryonasynplay As Byte     ' 是否收到通信响应是否继续异步放音：默认为0-不继续
    reserved2(5) As Byte        ' 保留
End Type
Public Node96_Data1 As SData1_096
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_096
    command As Integer          ' 命令代码
    vox_wt  As Integer          ' 等待循环播放语音
    var_send(9) As Byte         ' 发送变量ID
    var_receive(9) As Byte      ' 接收变量ID
    fileprefix(1) As Byte       ' 文件名前缀
    reserved1(21) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    nd_timeout As Integer       ' 超时节点ID
    reserved2(9) As Byte        ' 保留
End Type
Public Node96_Data2 As SData2_096

'------------------------------------------------------------------------
' 节点100：用户DLL
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_100
    reserved1(4) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node100_Data1 As SData1_100
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_100
    dll_fid As Integer          ' DLL文件ID
    reserved1(45) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved2(11) As Byte       ' 保留
End Type
Public Node100_Data2 As SData2_100

'------------------------------------------------------------------------
' 节点101：用户COM
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_101
    reserved1(4) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    reserved2(6) As Byte        ' 保留
End Type
Public Node101_Data1 As SData1_101
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_101
    reserved1(15) As Byte       ' 保留
    com_iid As Integer          ' COM接口ID
    reserved2(29) As Byte       ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved3(11) As Byte       ' 保留
End Type
Public Node101_Data2 As SData2_101

'------------------------------------------------------------------------
' 节点102：记录变量
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_102
    reserved1(4) As Byte        ' 保留
    log As Byte                 ' 被访问日志
    var_chg As Byte             ' 记录变量ID: 0 --不记录；1 -- 255 变量ID
    convert As Byte             ' 转换公式
    param1 As Byte              ' 参数1
    param2 As Byte              ' 参数2
    reserved2(2) As Byte        ' 保留
End Type
Public Node102_Data1 As SData1_102
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_102
    reserved1(15) As Byte       ' 保留
    com_iid As Integer          ' COM接口ID
    value(23) As Byte           ' 变量值
    reserved2(5) As Byte        ' 保留
    nd_parent As Integer        ' 父节点ID
    nd_child As Integer         ' 子节点ID
    reserved3(11) As Byte       ' 保留
End Type
Public Node102_Data2 As SData2_102

'------------------------------------------------------------------------
' Sun added 2001-09-28
' 节点255 - 节点连线
'/ 节点数据结构-Data1(03-0F/13b)
Public Type SData1_255
    reserved1(12) As Byte       ' 保留
End Type
Public Node255_Data1 As SData1_255
'/ 节点数据结构-Data2(30-6F/64b)
Public Type SData2_255
    StartNode As Integer        ' 起始节点ID
    EndNode As Integer          ' 终止节点ID
    Index As Byte               ' 0：唯一连线；其他：连线索引号
    Style As Byte               ' 连线线形
    Width As Byte               ' 连线粗细
    Color As Long               ' 连线颜色
    reserved1(52) As Integer    ' 保留
End Type
Public Node255_Data2 As SData2_255

'------------------------------------------------------------------------
' Sun added 2002-05-23
' Ver 2 Mend 1
'' Exchange File Header Definition
'' 512 Bytes
Public Type SCFFileHeader
    P_ID_Name       As String * 10  ' 分割标记                  (10)
    P_ID            As Byte         ' 项目ID                    (1)
    P_Type_Name     As String * 10  ' 分割标记                  (10)
    P_Type          As Byte         ' 项目类型                  (1)
    P_Name_Name     As String * 10  ' 分割标记                  (10)
    P_Name          As String * 50  ' 项目名称                  (50)
    P_Description_Name As String * 10  ' 分割标记               (10)
    P_Description   As String * 200 ' 项目描述文字              (200)
    P_Version_Name  As String * 10  ' 分割标记                  (10)
    P_Version       As String * 10  ' 项目版本号                (10)
    P_Auther_Name   As String * 10  ' 分割标记                  (10)
    P_Auther        As String * 50  ' 项目作者                  (50)
    P_User_Name     As String * 10  ' 分割标记                  (10)
    P_User          As String * 50  ' 项目用户                  (50)
    P_CreateTime_Name As String * 10  ' 分割标记                (10)
    P_CreateTime    As Date         ' 项目创建时间              (8)
    P_ModifyTime_Name As String * 10  ' 分割标记                (10)
    P_ModifyTime    As Date         ' 项目最后修改时间          (8)
    P_RCN_Name      As String * 10  ' 分割标记                  (10)
    P_RCN           As Long         ' 记录数                    (4)
    P_Reserved_Name As String * 10  ' 分割标记                  (10)
    Reserved        As String * 20  ' 填充保留                  (20)
End Type
'' Exchange File Record Definition
'' 128 Bytes
Public Type SCFFileRecord
    N_ID            As Integer      ' 节点ID                    (2)
    N_NO            As Byte         ' 节点编号                  (1)
    N_Description   As String * 32  ' 节点描述文字              (32)
    N_Page          As Byte         ' 节点所在页码              (1)
    N_Left          As Integer      ' 节点左坐标                (2)
    N_Top           As Integer      ' 节点上坐标                (2)
    N_Height        As Integer      ' 节点高度                  (2)
    N_Width         As Integer      ' 节点宽度                  (2)
    N_Data1(12)     As Byte         ' 节点数据段1               (13)
    N_Data2(63)     As Byte         ' 节点数据段2               (64)
    '' Sun modified 2008-01-18
    ''' From
    'Reserved        As String * 7   ' 保留                      (7)
    ''' To
    N_Tag           As String * 7   ' Tag                      (7)
End Type

Public Sub LoadNodeTypeNameList()

    gNodeTypeNameList(0) = LoadNationalResString(1490) '
    gNodeTypeNameList(1) = LoadNationalResString(1491) ' "1-用户Buffer定义日志"
    gNodeTypeNameList(2) = LoadNationalResString(1492) ' "2-用户Buffer定义变量"
    gNodeTypeNameList(6) = LoadNationalResString(1493) ' "6-无条件转移"
    gNodeTypeNameList(7) = LoadNationalResString(1494) ' "7-身份验证"
    gNodeTypeNameList(8) = LoadNationalResString(1495) ' "8-修改口令"
    gNodeTypeNameList(9) = LoadNationalResString(1496) ' "9-时间分支"
    gNodeTypeNameList(10) = LoadNationalResString(1497) ' "10-工作日设定"
    gNodeTypeNameList(16) = LoadNationalResString(1529) ' "16-条件分支"
    gNodeTypeNameList(17) = LoadNationalResString(1498) ' "17-选择服务语言"
    gNodeTypeNameList(18) = LoadNationalResString(1499) ' "18-发送数据"
    gNodeTypeNameList(19) = LoadNationalResString(1500) ' "19-无操作"
    gNodeTypeNameList(20) = LoadNationalResString(1501) ' "20-放音挂机"
    gNodeTypeNameList(21) = LoadNationalResString(1502) ' "21-放音继续"
    gNodeTypeNameList(22) = LoadNationalResString(1503) ' "22-放音等待按键"
    gNodeTypeNameList(23) = LoadNationalResString(1504) ' "23-放音转移"
    gNodeTypeNameList(28) = LoadNationalResString(1525) ' "28-TTS 放音"
    gNodeTypeNameList(40) = LoadNationalResString(1505) ' "40-建立留言"
    gNodeTypeNameList(41) = LoadNationalResString(1506) ' "41-察看留言"
    gNodeTypeNameList(50) = LoadNationalResString(1507) ' "50-间单传真"
    gNodeTypeNameList(51) = LoadNationalResString(1508) ' "51-TTF传真"
    gNodeTypeNameList(60) = LoadNationalResString(1509) ' "60-转接坐席"
    gNodeTypeNameList(61) = LoadNationalResString(1510) ' "61-转接坐席组"
    gNodeTypeNameList(62) = LoadNationalResString(1511) ' "62-增强转接坐席"
    gNodeTypeNameList(63) = LoadNationalResString(1512) ' "63-增强转接座席组"
    gNodeTypeNameList(69) = LoadNationalResString(1513) ' "69-转虚拟分机"
    gNodeTypeNameList(70) = LoadNationalResString(1600) ' "70-查询路由点"
    gNodeTypeNameList(71) = LoadNationalResString(1633) ' "71-查询座席状态"
    gNodeTypeNameList(80) = LoadNationalResString(1514) ' "80-进入ACD"
    gNodeTypeNameList(90) = LoadNationalResString(1515) ' "90-呼叫外线号码"
    gNodeTypeNameList(91) = LoadNationalResString(1583) ' "91-Calling Card"
    gNodeTypeNameList(96) = LoadNationalResString(1569) ' "96-异步通信"
    gNodeTypeNameList(100) = LoadNationalResString(1516) ' "100-用户DLL"
    gNodeTypeNameList(101) = LoadNationalResString(1517) ' "101-用户COM"
    gNodeTypeNameList(102) = LoadNationalResString(1518) ' "102-记录变量"
    gNodeTypeNameList(255) = LoadNationalResString(1519) ' "255-节点连线"
        
End Sub

Public Function GetNodePictureFileName(ByVal f_NodeNo As Byte) As String

    GetNodePictureFileName = App.path & "\Bitmaps\" & Right("000" & Trim(Str(f_NodeNo)), 3) & ".ico"
    
End Function
