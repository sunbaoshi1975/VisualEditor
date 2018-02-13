Attribute VB_Name = "mdlMessage"

'**************************************************************
'Error message defination
 
#If Language = 0 Then

 Const E001 = "本机网络设置有误，请检查网卡和软件设置！"
 Const E002 = "无法连接服务器，请检查网络连接！"
 Const E003 = "串口无应答，请检查设备连接情况！"
 Const E004 = "接收缓冲区溢出！"
 Const E005 = "传输缓冲区已满！"
 Const E006 = "Parity 错误！"
 Const E007 = "端口参数设置有误，请从新设置！"
 Const E008 = "连接超时，请稍后重试！"
 Const E009 = "系统文件不完整，请查找或重新安装！"
 Const E010 = "口令错误！请检查CapsLock键，并重新输入！"
 Const E011 = "非法操作者！登录取消！"
 Const E012 = "新口令与确认口令不一致，请重新输入。"
 Const E013 = "口令必须为3-8个字符，请重新输入。"
 Const E014 = "表已满无法新增，请先删除无用的记录。"
 Const E015 = "请先安装打印机！"
 Const E016 = "数据库无法打开，请检查数据源设置！"
 Const E017 = "数据库无法关闭，强制终止！"
 Const E018 = "数据库出错并且无法恢复，强制终止！"
 Const E019 = "安全验证数据库出错，请与系统管理员联系！"
 Const E020 = "没有此操作者，请重试！"
 Const E021 = "缓冲区数据丢失，无法去回复原先数据！"
 Const E022 = "数据无法写入服务器！请检查网络连接和软件设置！"
 Const E023 = "无法完成打印作业，请检查打印机设置。"
 Const E024 = "无法连接服务器，请检查网络连接！"
 Const E025 = "请输入内容小于32767。"
 Const E026 = "删除流程失败！"
 Const E027 = "编号范围不正确，请重新输入。"
 Const E028 = "无法从服务器读取数据！请检查网络连接和软件设置！"
 Const E029 = "找不到指定编号的资源！"
 Const E030 = "不能在此添加新的资源，请选择适当的插入位置！"
 Const E031 = "数据类型不正确，请重新输入。"
 Const E032 = "节点编号不能<256，请重新输入。"
 Const E033 = "节点无法编辑。"
 Const E034 = "根节点：请输入内容大于255，小于32767。"
 Const E035 = "父节点：请输入内容大于255，小于32767。"
 Const E036 = "请输入用户定义变量数小于256。"
 Const E037 = "请输入“提示用户输入号码”语音ID小于32767。"
 Const E038 = "请输入“提示用户输入口令”语音ID小于32767。"
 Const E039 = "请输入“提示用户重新输入”语音ID小于32767。"
 Const E040 = "请输入“COM接口ID”大于0，小于32767。"
 Const E041 = "请输入“成功转节点ID”大于255，小于32767。"
 Const E042 = "请输入“失败转节点ID”大于255，小于32767。"
 Const E043 = "请输入“用户号码记录”大于0，小于256。"
 Const E044 = "请输入“用户口令记录”大于0，小于256。"
 Const E045 = "请输入“验证次数日志”大于0，小于256。"
 Const E046 = "请输入“验证结果日志”大于0，小于256。"
 Const E047 = "请输入“最大尝试次数”大于0，小于256。"
 Const E048 = "请输入“被访问日志”大于0，小于16。"
 Const E049 = "请输入“用户口令最大长度”大于0，小于256。"
 Const E050 = "请输入“用户号码最大长度”大于0，小于256。"
 Const E051 = "请输入“节点超时(S)”大于0，小于256。"
 Const E052 = "请输入“修改结果日志”大于0，小于256。"
 Const E053 = "请输入“提示用户输入新口令”语音ID小于32767。"
 Const E054 = "请输入“提示用户再次确认”语音ID小于32767。"
 Const E055 = "请输入“两次不一致重新输入”语音ID小于32767。"
 Const E056 = "请输入“提示用户修改成功”语音ID小于32767。"
 Const E057 = "请输入正确的时间值。"
 Const E058 = "请选择变量类型。"
 Const E059 = "请输入“时间段转移节点ID”大于255，小于32767。"
 Const E060 = "注意：当前流程尚未指定根节点！根节点是呼叫流程的起始节点，没有根节点流程将无法工作。"
 Const E061 = "请输入节点延时时长。"
 Const E062 = "请输入“语言选择结果记录”大于0，小于256。"
 Const E063 = "流程同步失败，请确认目标数据库是否可访问！"
 Const E064 = "资源同步失败，请确认目标数据库是否可访问！"
 Const E065 = "请输入“休息日转节点ID”大于255，小于32767。"
 Const E066 = "时间段次序错误。"
 Const E067 = "请输入“转移节点ID”大于255，小于32767。"
 Const E068 = "不是合法的资源文件或文件被破坏，不能导入！"
 Const E069 = "。"
 Const E070 = "失败转入节点：请输入内容大于255，小于32767。"
 Const E071 = "请输入“发送变量”小于256。"
 Const E072 = "子节点：请输入内容大于255，小于32767。"
 Const E073 = "附件内容：输入字符超出长度。"
 Const E074 = "请输入“接收者”。"
 Const E075 = "请输入“数据包类型”小于256。"
 Const E076 = "请输入“播放语音”小于32767。"
 Const E077 = "请输入“按键长度”小于256。"
 Const E078 = "长度超出范围。"
 Const E079 = "变量名称长度超出范围。"
 Const E080 = "请输入“使用变量ID”小于256。"
 Const E081 = "请输入“按键长度”小于256。"
 Const E082 = "请输入“按键最大间隔”小于256。"
 Const E083 = "请输入“按键记录”小于256。"
 'Const E084 = "请输入“录音时间长”小于256。"
 Const E084 = "请输入“录音时间长”小于65536(18 Hours)。"    'Michael Modified
 Const E085 = "请输入“座席组ID”小于32767。"
 Const E086 = "请输入“等待播放次数”小于256。"
 Const E087 = "DLL文件ID不能大于32767。"
 Const E088 = "COM文件ID不能大于32767。"
 Const E089 = "请选择需要打开的流程编号。"
 Const E090 = "请输入“前导播放语音ID”小于32767。"
 Const E091 = "请输入“后续播放语音ID”小于32767。"
 Const E092 = "请输入“提示语音ID”小于32767。"
 Const E093 = "请输入“传真文件ID”小于32767。"
 Const E094 = "请输入“LOGO文件ID”小于32767。"
 Const E095 = "请输入“表格格式文件ID”小于32767。"
 Const E096 = "请输入“转接提示语音ID”小于32767。"
 Const E097 = "请输入“等待循环语音ID”小于32767。"
 Const E098 = "请输入“没上班提示语音ID”小于32767。"
 Const E099 = "请输入“坐席忙提示语音ID”小于32767。"
 Const E100 = "请输入“按键转节点ID”大于255，小于32767。"
 Const E101 = "请输入“传真标题资源ID”小于32767。"
 Const E102 = "请输入“发出号码资源ID”小于32767。"
 Const E103 = "请输入“记录变量ID”小于256。"
 Const E104 = "请输入“TTS 播放资源ID”小于32767。"
 Const E105 = "请输入“替代播放语音ID”小于32767。。"
 Const E106 = "请输入“外拨组号”小于256。"
 Const E107 = "请输入“余额不足或卡无效提示语音”小于32767。"
 Const E108 = "请输入“最大静音时长”小于60。"
 Const E109 = "请输入“最小录音时长”小于180。"
 Const E110 = "请输入“转移节点ID”大于255，小于32767。"
 Const E111 = "请输入“没有上班节点ID”大于255，小于32767。"
 Const E112 = "请输入“工作组忙转节点ID”大于255，小于32767。"
 Const E113 = "请输入“转接成功转节点ID”大于255，小于32767。"
 Const E114 = "请输入“虚拟分机号”小于 2,147,483,647。"
 Const E115 = "请输入正确时间。"
 Const E116 = "请输入“被叫号码”长度应小于16。"
 Const E117 = "请输入“主叫座席ID”应小于256。"
 Const E118 = "请输入“优先级”应小于256。"
 Const E119 = "请输入“成功转接语音ID”小于32767。"
 Const E120 = "请输入“座席无应答提示语音ID”小于32767。"
 Const E121 = "请输入“无应答转节点ID”大于255，小于32767。"
 Const E122 = "请输入“拨分机前延时(秒)”大于0，小于256。"
 Const E123 = "请输入“通信超时(秒)”，数值在0到180之间。"
 Const E124 = "请输入“超时提醒(分钟)”，数值在0到60之间。"
 Const E125 = "按键定义不能重复，请重新设置。"
 Const E126 = "请输入“节点超时(秒)”，数值在0到3600之间。"
 Const E127 = "请输入新流程ID。"
 Const E128 = "该流程ID已经被使用，请重新选择。"
 Const E129 = "流程ID另存失败，请检查网络和数据库！"
 'Michael Added @7-4-07
 Const E130 = "请重新设置用户变量ID, 数值在1到255之间."
 Const E131 = "请输入“等待流程入口节点ID”，大于255, 小于32767。"
 'Michael Added @ July,9,07
 Const E132 = "请输入“转接座席ID记录变量”，大于0, 小于256。"
 Const E133 = "请输入“转接座席工号记录变量”，大于0, 小于256。"
 Const E134 = "查询字段名不可为空"
 Const E135 = "查询运算符不可为空"
 Const E136 = "查询表达式不可为空"
  'Michael Added @ 2007-11-26
 Const E137 = "驱动器不存在或路径错误,请输入绝对路径!"
 Const E138 = "资源ID不可为0或空"
 Const E139 = "资源编号重复"
 Const E140 = "此资源文件类型不被支持或资源文件名不完整,请确认后再重新合成..."
 Const E141 = "项目ID不可为空!"
 Const E142 = "项目名不可为空!"
 Const E143 = "项目ID必须为1到255之间的整数!"
 Const E144 = "语音文件路径不正确，且无法创建，请在启动后检查设置！见菜单“选项”->“常规”->“系统语音根目录”"
 Const E145 = "请输入“留言进行提示间隔(秒)”，数值不能大于“录音时间长(秒)”"
 Const E146 = ""
 Const E147 = ""
 Const E148 = "节点ID的值必须在256和32767之间"
 Const E149 = "语音资源ID需小于32767。"
 Const E150 = "资源ID需小于32767。"
 
 'Add End
 '********  Added End *******************
 Const E899 = "服务器数据更新出错！请检查网络连接和软件设置！"
 Const E999 = "错误消息代号！"

'**************************************************************
'Information message defination
'' < 100 for mesage box
 Const M001 = "该程序正在运行，请检查窗口是否最小化。"
 Const M002 = "操作者新增成功，缺省口令为‘888888’，请提醒该操作者及时修改口令。"
 Const M003 = "没有退出系统的权限。"
 Const M004 = "此记录为标志记录，请不要删除。"
 Const M005 = "该评价表已被锁定，不能进行编辑。"
 Const M006 = "打印作业已经完成！"
 Const M007 = "无此记录，请重新选择或输入。"
 Const M008 = "请指明希望在哪条记录之前插入新记录。"
 Const M009 = "数据已存在，请重新输入？"
 Const M010 = "流程已删除！"
 Const M011 = "流程同步已完成！"
 Const M012 = "数据源与数据目标不能相同！"
 Const M013 = "资源同步已完成！"
'' >= 100 AND < 200 for status and help information
 Const M100 = "正在进行系统初始化......"
 Const M101 = "完成"
 Const M102 = "正在发送数据……"
 Const M103 = "正在接收数据……"
 Const M104 = "出错!"
 Const M105 = "您没有更改数据的权限！"
 Const M106 = "正在传送功能列表......"
 Const M107 = "正在复位控制器......"
 Const M108 = "正在准备打印数据......"
 Const M109 = "正在更新服务器端成绩单布局......"
 Const M110 = "请输入适当数据"
 Const M111 = "提示：口令为3-8个字符或数字。"
 Const M112 = "提示：用鼠标点击带‘！’的标签，可弹出检索框。"
 Const M113 = "提示：从列表中选择分类或者用A-G键快速输入。"
 Const M114 = "提示：从列表中选择级别或者用数字键快速输入。"
 Const M115 = "提示：本模块用于设置串行通讯口，以便收发数据。"
 Const M116 = "提示：本模块用于值班员登录，以便进行更多操作。"
 Const M117 = "提示：本模块用于修改值班员口令。"
 Const M118 = "提示：本模块用于查看和设置工作日志。"
 Const M119 = "提示：本模块用于创建、修改和删除值班员。"
 Const M120 = "提示：本模块用于修改不同级别报警的上下限特征值。"
 Const M121 = "提示：本模块用于更改各报警事件的报警方式。"
 Const M122 = "提示：本模块用于安排值班员的值班时间表。"
 Const M123 = "提示：本模块用于动态监视参数变化情况。"
 Const M124 = "提示：本模块用于向指定人员发送通知。"
 Const M125 = "正在关闭系统......"
 Const M126 = "提示：流程已存在，请重新创建！"
 Const M127 = "提示：流程号不能为空，请重新输入！"
 Const M128 = "提示：流程创建完毕！"
 Const M129 = "提示：数据不存在！"
 Const M130 = "提示：创建流程出错！"
 Const M131 = "提示：请先创建或打开流程！"
 Const M132 = "流程文件导出成功！"
 Const M133 = "本节点为系统节点，不能删除！"
 Const M134 = "流程文件导入成功！"
 Const M135 = "不能删除根节点，除非先另外指定根节点！"
 Const M136 = "本节点为系统节点，不能进行拷贝、粘贴操作！"
 Const M137 = "已经是首月，月份不能再前移！"
 Const M138 = "当前月不在选定范围内！"
 Const M139 = "已经是末月，月份不能再后移！"
 Const M140 = "该变量已经在目标列表中存在！"
 Const M141 = "此页面不为空，不能删除！请先删除页面中的所有节点。"
 Const M142 = "流程已经被打开，不能重复打开。"
 Const M143 = "流程已经打开，处于编辑状态，请先关闭该流程。"
 Const M144 = "提示：目前尚未为流程指定主COM接口资源，请在‘编辑->流程属性’对话框中设置。"
 Const M145 = "流程另存成功。"
 Const M146 = "资源文件导出成功！"
 Const M147 = "导入资源文件成功！"
 Const M148 = "请先选择需要复制的资源，再进行复制操作。"
 Const M149 = "资源列表为空,请先添加添加资源再进行相应操作!"
 Const M150 = "您使用了项目复制附加选项,但没有选择目标资源替换类型,请确认。"
 Const M151 = "系统未安装中文TTS引擎,TTS中文转换将不能正常使用,如需使用请先安装TTS中文语音包"
 Const M152 = "节点ID小于255的节点统计标签不可被编辑!"
 Const M153 = ""
 Const M154 = ""
 Const M155 = ""
 Const M156 = ""
 Const M157 = ""
 Const M158 = ""
 Const M159 = ""

 '' >= 200 AND < 300 for Database operation status
 Const M200 = "Database OK!"
 Const M201 = " opening..."
 Const M202 = " closing..."
 Const M203 = " read..."
 Const M204 = " written..."
 Const M205 = " Connecting SQLServer..."
 Const M206 = " Disconnecting SQLServer..."
 Const M207 = " Reading SQLServer..."
 Const M208 = " Writing SQLServer..."
 Const M209 = " searching..."
 Const M210 = ""
 Const M299 = "Database Error!"

'**************************************************************
'Qustion message defination
 Const Q001 = "请确定是否要退出当前应用程序？"
 Const Q002 = "此操作将修改原有数据，是否继续？"
 Const Q003 = "非正常退出可能使数据丢失，是否继续？"
 Const Q004 = "该编号对应的内容不存在，是否新建？"
 Const Q005 = "当前的数据尚未保存，是否退出？"
 Const Q006 = "当前的数据尚未保存，是否保存？"
 Const Q007 = "当前数据库转换失败，是否继续进行？"
 Const Q008 = "确实要删除当前流程吗？"
 Const Q009 = "确实要注销吗？"
 Const Q010 = "确实要删除吗？"
 Const Q011 = "确实要清除所有事件吗？"
 Const Q012 = "该编号资源已经存在，是否覆盖？"
 Const Q013 = "请先创建流程"
 Const Q014 = "数据已存在，请重新输入？"
 Const Q015 = "提示：数据已保存！是否继续？"
 Const Q016 = "数据不存在！"
 Const Q017 = "表内有多条记录存在！"
 Const Q018 = "此操作将恢复选定时段内所有月份的节假日缺省设定，是否继续？"
 Const Q019 = "此操作将清除选定时段内所有月份的节假日设定（包括周六、周日），是否继续？"
 'Michael Added @ 2007-11-27
 Const Q020 = "确实要删除当前项目吗？"
 Const Q021 = "文件夹不存在,需要创建吗?"

#ElseIf Language = 1 Then

'''**************************************************************

 Const E001 = "Network setting error, please check NIC and software setting"
 Const E002 = "Can not connect server, please check connection of network"
 Const E003 = "No response, please check connection of equipment"
 Const E004 = "Receive buffer overflow"
 Const E005 = "Transmission buffer full"
 Const E006 = "Parity error!"
 Const E007 = "Error of  port parameter setting, please set up again!"
 Const E008 = "Connection overtime, please try it again later."
 Const E009 = "System file not integrated, please locating or reinstall."
 Const E010 = "Password incorrect! Please check key CapsLock, and input again."
 Const E011 = "Illegal operator! Login cancelled."
 Const E012 = "The new password inconsistency  with the confirm password, please input again."
 Const E013 = "Password should be length of 3-8 character, please input again."
 Const E014 = "The table is full can not add, please delete useless record first."
 Const E015 = "Please install printer first"
 Const E016 = "Cannot open the database, please check data source setting."
 Const E017 = "Cannot close database, compulsion terminate."
 Const E018 = "Error of database and can not recovery, compulsion stop."
 Const E019 = "Security identify database error, please connect with system custodian."
 Const E020 = "This operator does not exist, please try again."
 Const E021 = "Lose of data buffer, cannot response previous data."
 Const E022 = "Data can not import server, please check network connection and software setting."
 Const E023 = "Cannot finish printing, please check printer setting"
 Const E024 = "Cannot connect server, please check network connection"
 Const E025 = "Please input the content smaller than 32767."
 Const E026 = "Error of deleting flow!"
 Const E027 = "Number range incorrect, please input again."
 Const E028 = "Cannot read data from server, please check network connect and software setting!"
 Const E029 = "Cannot find specified number source."
 Const E030 = "Cannot add new source here, please choose appropriately insert position."
 Const E031 = "Incorrect data type, please input again."
 Const E032 = "Node number can not smaller than 256,please input again."
 Const E033 = "node cannot edit."
 Const E034 = "Root node: please input content large than 255 and smaller than 32767."
 Const E035 = "Parent node: please input content large than 255 and smaller than 32767."
 Const E036 = "Please input user defined variable smaller than 256."
 Const E037 = "Please input  clue user input number voice ID smaller than 32767."
 Const E038 = "Please input clue user input password voice ID smaller than 32767."
 Const E039 = "Please input clue user to input again voice ID smaller than 32767."
 Const E040 = "Please input  COM port ID larger than 0 and smaller than 32767."
 Const E041 = "Please input switching node successful ID larger than 255, smaller than 32767."
 Const E042 = "Please input  failed switching node ID larger than 255, and smaller than 32767."
 Const E043 = "Please input user number record larger than 0 and smaller than 256."
 Const E044 = "Please input user password record larger than 0 and smaller than 256."
 Const E045 = "Please input check times log larger than 0 and smaller than 256."
 Const E046 = "Please input check result log larger than 0 and smaller than 256."
 Const E047 = "Please input maximum try times larger than 0 and smaller than 256"
 Const E048 = "Please input visited log larger than 0 and smaller than 16"
 Const E049 = "Please input user password maximum length larger than 0 and smaller than 256"
 Const E050 = "Please input user number maximum length larger than 0 and smaller than 256."
 Const E051 = "Please input node overtime larger than 0 and smaller than 256."
 Const E052 = "Please input  revise result log larger than 0 and smaller than 256"
 Const E053 = "Please input clue user input new password voice ID smaller than 32767."
 Const E054 = "Please input  clue user confirm again voice ID smaller than 32767."
 Const E055 = "Please input inconsistency for twice input ,please input again voice ID smaller than 32767."
 Const E056 = "Please input clue user successful revise voice ID smaller than 32767"
 Const E057 = "Please input corrected time value"
 Const E058 = "Please select variable type"
 Const E059 = "Please input  time transfer node ID larger than 255, smaller than 32767."
' Const E060 = "Attention: current flow has not specified root node! Root node the node to start calling flow, without root node flow can not work properly."
 Const E061 = "  Please input node delay"
 Const E062 = "Please input language select result record  larger than 0 and smaller than 256."
 Const E063 = "Failed to synchronize call flow! Please check the target database."
 Const E064 = "Failed to synchronize resource! Please check the target database."
 Const E065 = "Please input  holiday transfer point ID larger than 255 and smaller than 32767."
 Const E066 = "  Error of time segment sequence"
 Const E067 = "Please input transfer node ID  larger than 255 and smaller than 32767."
 Const E068 = "Invalid resource file!"
 Const E069 = "。"
 Const E070 = "Failed to switching to node: please input content larger than 255, and smaller than 32767"
 Const E071 = "Please input: sending variable smaller than 256"
 Const E072 = "子节点：请输入内容大于255，小于32767。."
 Const E073 = "Attachment content: input character exceed length"
 Const E074 = "Please input  receiver"
 Const E075 = "Please input data package type smaller than 256"
 Const E076 = "Please input play voice smaller than 32767."
 Const E077 = "Please input press length smaller than 256"
 Const E078 = "    length exceed range"
 Const E079 = "Variable name length exceed range"
 Const E080 = "Please input using variable ID smaller than 256"
 Const E081 = "Please input press length smaller than 256"
 Const E082 = "lease input press maximum interval  smaller than 256"
 Const E083 = "Please input press record smaller than 256"
 Const E084 = "Please input record time smaller than 256"
 Const E085 = "Please input seat group ID smaller than 256"
 Const E086 = "Please input waiting play times smaller than 32767"
 Const E087 = "DLL file ID cannot larger than 32767"
 Const E088 = " COM file ID cannot larger than 32767"
 Const E089 = "Please select the flow to open."
 Const E090 = "Please input play voice preface ID smaller than 32767"
 Const E091 = "Please input play voice forward ID smaller than 32767"
 Const E092 = "Please input clue voice ID smaller than 32767"
 Const E093 = "Please input fax file ID smaller 32767"
 Const E094 = "Please input LOGO file ID smaller than 32767"
 Const E095 = "Please input table format file ID smaller than 32767"
 Const E096 = "Please input switching clue voice ID smaller than 32767"
 Const E097 = "Please input waiting loop voice ID smaller than 32767"
 Const E098 = "Please input clue voice ID of absent from work  smaller than 32767"
 Const E099 = "Please input clue voice ID of busy seat smaller than 32767"
 Const E100 = "Please input transfer point ID of press key larger than 255 and smaller than 32767"
 Const E101 = "Please input fax title source ID smaller than 32767"
 Const E102 = "Please input sending number source ID smaller than 32767"
 Const E103 = "Please input 'Record variable ID' which must be smaller than 256"
 Const E104 = "Please input 'TTS Resource ID' which must be smaller that 32767."
 Const E105 = "Please input 'Alter Voice Resource ID' which must be smaller that 32767."
 Const E106 = "Please input 'OB Group' which must be smaller than 256"
 Const E107 = "Please input 'Out of money or invalid card voice' which must be smaller than 32767"
 Const E108 = "Please input 'Max. Silence Time' which must be smaller than 60."
 Const E109 = "Please input 'Min Record Length', which should less then 180 seconds."
 Const E110 = "Please input  transfer node ID larger than 255, smaller than 32767"
 Const E111 = "Please input  node ID of absent from work  larger than 255 and smaller than 32767"
 Const E112 = "Please input  transfer Node ID of busy working group larger than 255 and smaller than 32767"
 Const E113 = "Please input  successful switching transfer node ID larger than 255 and smaller than 32767"
 Const E114 = "Please input  virtual extension number  smaller than 2,147,483,647"
 Const E115 = "Please input correct time"
 Const E116 = "Please input  calling number length should shorter than 16"
 Const E117 = "Please input calling extension ID should be smaller than 256"
 Const E118 = "Please input  priority  should smaller than 256"
 Const E119 = "Please input successful transfer voice ID smaller than 32767"
 Const E120 = "Please input  clue voice ID of no response of extension smaller than 32767"
 Const E121 = "Please input < transfer node ID of no response > larger than 255 and smaller than 32767"
 Const E122 = "Please input delay seconds before dialing extension number, larger than 0 and smaller than 256."
 Const E123 = "Please input 'Wait Timeout(s)', between 0 to 180."
 Const E124 = "Please input 'Timeout Annouce(min)', between 0 to 60."
 Const E125 = "Key is already defined. Please select another key."
 Const E126 = "Please input 'Node Timeout(s)', between 0 to 3600."
 Const E127 = "Please enter the new call flow id."
 Const E128 = "Call flow id is used, please try another one."
 Const E129 = "Failed to execute 'Save As' operation, please check network and database."
 Const E130 = "User variable must be an integer between 1 to 255, Please Reset!"
 Const E131 = "Must be an integer bigger then 255 and less then 32767!"
 Const E132 = "Must be an integer bigger then 0 and less then 256!"
 Const E133 = "Must be an integer bigger then 0 and less then 256!"
 Const E134 = "Search keyword could not be NULL"
 Const E135 = "Search operator could not be NULL!"
 Const E136 = "Search express could not be NULL!"
 Const E137 = "Driver is not exist or incorrect file path ,enter absolutely path please!"
 Const E138 = "Resource ID could not be 0 or NULL"
 Const E139 = "Duplicate Resource ID"
 Const E140 = "The Type of resource file is not supported or the file name is not complete, Confirm Please ..."
 Const E141 = "Project ID could not be NULL!"
 Const E142 = "Project ID name could not be NULL!"
 Const E143 = "Project ID must be an integer between 1 to 255!"
 Const E144 = "Invalid voice path and failed to create it. Please check settings. Refer to menu 'Options'->'General'->'System voice path'"
 Const E145 = "Please input‘Recoring notify interval(s)’，which can't larger than 'Max Record Duration (sec)'"
 Const E146 = ""
 Const E147 = ""
 Const E148 = "Node ID should have value between 256 to 32767"
 Const E149 = "Voice resource ID should not acceed 32767"
 Const E150 = "Resource ID should not acceed 32767"
  
 Const E899 = "Error of updating data in server! Please check network connection and software setting!"
 Const E999 = "error code!"
'**************************************************************
'Information message defination
'' < 100 for mesage box
 Const M001 = "this program is executed now, please check if the window is minimum."
 Const M002 = "operator successful added, password <88888888> please reminder operator to revise password on time."
 Const M003 = " no limitation of exit system"
 Const M004 = " this record is marker record, please do not delete it."
 Const M005 = " this evaluation table has been locked, can not edit."
 Const M006 = " finish print"
 Const M007 = "this record does not exist. Please select or input again"
 Const M008 = " please point out before which record you want to insert new record."
 Const M009 = "data existed, please input again?"
 Const M010 = "flow deleted!"
 Const M011 = "Call flow synchronizing finished!"
 Const M012 = "Data Source must differ from Data Target!"
 Const M013 = "Resource synchronizing finished!"
'' >= 100 AND < 200 for status and help information
 Const M100 = " system initializing…."
 Const M101 = "finish"
 Const M102 = "sending data…"
 Const M103 = "receiving data…"
 Const M104 = "error"
 Const M105 = "you do not have the limitation to revise data"
 Const M106 = "sending functional listing...."
 Const M107 = "restore control function..."
 Const M108 = "preparing for printing data..."
 Const M109 = "updating server data......"
 Const M110 = "please input appropriate data."
 Const M111 = " Notes: password should be length of 3-8 character or integer."
 Const M112 = "Notes: use mouse to click ! label, then it will display retriever."
 Const M113 = "Notes: choose assort from listing or use A-G hotkey."
 Const M114 = "Notes：please select from list or key in directly"
 Const M115 = " Notes: this module is used for setting bunch telecommunicate port , in order to sending and receive data."
 Const M116 = "Notes: this module is for login of staff, in order to have more operation"
 Const M117 = " Notes: this module is used to revise password of staff"
 Const M118 = "Notes: this module is used to check and setting working journal"
 Const M119 = " Notes: this module is used to create, revise and delete staff."
 Const M120 = " Notes: this module is used to revise maximum and minumn limitation of warning from different level"
 Const M121 = "Notes: this module is used to modify the warning mode of all the warning events"
 Const M122 = " Notes: this module is used to arrange the schedule on duty for staff"
 Const M123 = "Notes: this module is used to supervise the changeable of parameter actively"
 Const M124 = "Notes: this module is used to send inform to specified staff."
 Const M125 = " closing system......"
 Const M126 = "  Notes: flow already existed, please create again."
 Const M127 = "Notes: flow number can not be empty, please input again."
 Const M128 = "Notes: finish create flow"
 Const M129 = "Notes: data does not existed"
 Const M130 = "Notes: error of creating flow"
 Const M131 = "Notes: please create or open a flow firstly."
 Const M132 = "Finished to export flow to file."
 Const M133 = "This node is system node ,can not be deleted."
 Const M134 = "Successful import of flow file"
 Const M135 = "Can not delete root node, unless specify other root node first."
 Const M136 = "This node is the system node, can not copy, paste"
 Const M137 = "It is already the first month, can not move forward in month"
 Const M138 = "Current month does not be included in select range."
 Const M139 = "It is already end of the month, can not move backward of month"
 Const M140 = "The variable is already in target list！"
 Const M141 = "Can not remove this page. Please delete all elements on it."
 Const M142 = "Call flow has already been opened!"
 Const M143 = "Call flow is editing, please close it firstly."
 Const M144 = "Notes: There's no COM Interface ID assigned with the flow, please set in 'Edit->Flow Property' dialog."
 Const M145 = "'Save as' operation finished."
 Const M146 = "Finished to export resource to file."
 Const M147 = "Finished to import resource from file."
 Const M148 = "Please select resource item before copying."
 Const M150 = "You have select Project Copy Option, but the destination resource type is NULL, Confirm Please!"
 Const M151 = "Please install TTS engine first."
 Const M152 = "NodeTag can't be edited when NodeID less then 255 !"
 Const M153 = ""
 Const M154 = ""
 Const M155 = ""
 Const M156 = ""
 Const M157 = ""
 Const M158 = ""
 Const M159 = ""
 
 '' >= 200 AND < 300 for Database operation status
 Const M200 = "Database OK!"   '确定数据库
 Const M201 = " opening..."   '正在打开数据库
 Const M202 = " closing..."  ' 正在关闭数据库
 Const M203 = " read..."    '正在读取数据库
 Const M204 = " written..." '  正在写入数据库
 Const M205 = " Connecting SQLServer..."  ' 正在连接SQL 服务器
 Const M206 = " Disconnecting SQLServer..." ' 断开联系SQL 服务器
 Const M207 = " Reading SQLServer..." ' 读取SQL 服务器
 Const M208 = " Writing SQLServer..." ' 写入SQL 服务器
 Const M209 = " searching..." ' 正在查询
 Const M210 = ""
 Const M299 = "Database Error!" '  数据库错误

'**************************************************************
'Qustion message defination
 Const Q001 = "Are you sure you want to quit the currently application program>"
 Const Q002 = " this operation will revise the existed data, still want continue?"
 Const Q003 = " Illegal exit may cause lost of data, continue?"
 Const Q004 = "the number correspondence to the content does not exist, do you want to create?"
 Const Q005 = " save changes to the data before exiting?"
 Const Q006 = "current data have not been saved , want to save?"
 Const Q007 = " failed to transfer currently database, want to continue?"
 Const Q008 = "  Are you sure to delete currently flow?"
 Const Q009 = " Are you sure to log off?"
 Const Q010 = "Are you sure to delete?"
 Const Q011 = " Are you sure to clear all the events?"
 Const Q012 = "Resource ID is used, are you sure you want to overwrite?"
 Const Q013 = "please create flow first."
 Const Q014 = "data already existed, please input again"
 Const Q015 = " Notes: data have been saved, want to continue?"
 Const Q016 = " data does not exist."
 Const Q017 = "表内有多条记录存在！"
 Const Q018 = "此操作将恢复选定时段内所有月份的节假日缺省设定，是否继续？"
 Const Q019 = "此操作将清除选定时段内所有月份的节假日设定（包括周六、周日），是否继续？"
 'Michael Added @ 2007-11-27
 Const Q020 = "Are you sure to delete currently Project?"
 Const Q021 = "Folder don't exist, Create?"

#End If

'**************************************************************
 Const Def_Message_Code_Error = E999
 Const Def_Dummy_Msg = M101
 Const Def_Dummy_Help = M110

'Show message box with specific parametre
'
'   E??? : error message with one button--"OK"
'   Q??? : question message with tow buttons--"Yes/No"
'   M??? : information message with one button--"OK"
'
Public Function Message(strMsgNo As String, Optional nDefaultButton As Integer = vbDefaultButton1)

Dim strMsgContens$
 
'缺省消息代号->消息代号错误
strMsgContens$ = Def_Message_Code_Error

'分析消息类型
Select Case UCase(Left(strMsgNo, 1))
  Case "M"
    Select Case Right(strMsgNo, 3)
      Case "001"
        strMsgContens$ = M001
      Case "002"
        strMsgContens$ = M002
      Case "003"
        strMsgContens$ = M003
      Case "004"
        strMsgContens$ = M004
      Case "005"
        strMsgContens$ = M005
      Case "006"
        strMsgContens$ = M006
      Case "007"
        strMsgContens$ = M007
      Case "008"
        strMsgContens$ = M008
      Case "009"
        strMsgContens$ = M009
      Case "010"
        strMsgContens$ = M010
      Case "011"
        strMsgContens$ = M011
      Case "012"
        strMsgContens$ = M012
      Case "013"
        strMsgContens$ = M013
      Case "014"
        strMsgContens$ = M014
      Case "015"
        strMsgContens$ = M015
      Case "016"
        strMsgContens$ = M016
      Case "017"
        strMsgContens$ = M017
      Case "018"
        strMsgContens$ = M018
      Case "019"
        strMsgContens$ = M019
      Case "020"
        strMsgContens$ = M020
      Case "021"
        strMsgContens$ = M021
      Case "022"
        strMsgContens$ = M022
      Case "023"
        strMsgContens$ = M023
      Case "024"
        strMsgContens$ = M024
      Case "025"
        strMsgContens$ = M025
      Case "026"
        strMsgContens$ = M026
      Case "027"
        strMsgContens$ = M027
      Case "028"
        strMsgContens$ = M028
      Case "029"
        strMsgContens$ = M029
      Case "030"
        strMsgContens$ = M030
      Case "031"
        strMsgContens$ = M031
      Case "032"
        strMsgContens$ = M032
      Case "033"
        strMsgContens$ = M033
      Case "034"
        strMsgContens$ = M034
      Case "035"
        strMsgContens$ = M035
      Case "036"
        strMsgContens$ = M036
      Case "037"
        strMsgContens$ = M037
      Case "038"
        strMsgContens$ = M038
      Case "039"
        strMsgContens$ = M039
      Case "040"
        strMsgContens$ = M040
      Case "041"
        strMsgContens$ = M041
      Case "042"
        strMsgContens$ = M042
      Case "043"
        strMsgContens$ = M043
      Case "044"
        strMsgContens$ = M044
      Case "045"
        strMsgContens$ = M045
      Case "046"
        strMsgContens$ = M046
      Case "047"
        strMsgContens$ = M047
      Case "048"
        strMsgContens$ = M048
      Case "049"
        strMsgContens$ = M049
      Case "050"
        strMsgContens$ = M050
      Case "051"
        strMsgContens$ = M051
      Case "052"
        strMsgContens$ = M052
      Case "053"
        strMsgContens$ = M053
      Case "054"
        strMsgContens$ = M054
      Case "055"
        strMsgContens$ = M055
      Case "056"
        strMsgContens$ = M056
      Case "057"
        strMsgContens$ = M057
      Case "058"
        strMsgContens$ = M058
      Case "059"
        strMsgContens$ = M059
      Case "060"
        strMsgContens$ = M060
      Case "061"
        strMsgContens$ = M061
      Case "062"
        strMsgContens$ = M062
      Case "063"
        strMsgContens$ = M063
      Case "064"
        strMsgContens$ = M064
      Case "065"
        strMsgContens$ = M065
      Case "066"
        strMsgContens$ = M066
      Case "067"
        strMsgContens$ = M067
      Case "068"
        strMsgContens$ = M068
      Case "069"
        strMsgContens$ = M069
      Case "070"
        strMsgContens$ = M070
      Case "071"
        strMsgContens$ = M071
      Case "072"
        strMsgContens$ = M072
      Case "073"
        strMsgContens$ = M073
      Case "074"
        strMsgContens$ = M074
      Case "075"
        strMsgContens$ = M075
      Case "076"
        strMsgContens$ = M076
      Case "077"
        strMsgContens$ = M077
      Case "078"
        strMsgContens$ = M078
      Case "079"
        strMsgContens$ = M079
      Case "080"
        strMsgContens$ = M080
      Case "081"
        strMsgContens$ = M081
      Case "082"
        strMsgContens$ = M082
      Case "083"
        strMsgContens$ = M083
      Case "084"
        strMsgContens$ = M084
      Case "085"
        strMsgContens$ = M085
      Case "086"
        strMsgContens$ = M086
      Case "087"
        strMsgContens$ = M087
      Case "088"
        strMsgContens$ = M088
      Case "089"
        strMsgContens$ = M089
      Case "090"
        strMsgContens$ = M090
      Case "091"
        strMsgContens$ = M091
      Case "092"
        strMsgContens$ = M092
      Case "093"
        strMsgContens$ = M093
      Case "094"
        strMsgContens$ = M094
      Case "095"
        strMsgContens$ = M095
      Case "096"
        strMsgContens$ = M096
      Case "097"
        strMsgContens$ = M097
      Case "098"
        strMsgContens$ = M098
      Case "099"
        strMsgContens$ = M099
      Case "100"
        strMsgContens$ = M100
      Case "101"
        strMsgContens$ = M101
      Case "102"
        strMsgContens$ = M102
      Case "103"
        strMsgContens$ = M103
      Case "104"
        strMsgContens$ = M104
      Case "105"
        strMsgContens$ = M105
      Case "106"
        strMsgContens$ = M106
      Case "107"
        strMsgContens$ = M107
      Case "108"
        strMsgContens$ = M108
      Case "109"
        strMsgContens$ = M109
      Case "110"
        strMsgContens$ = M110
      Case "111"
        strMsgContens$ = M111
      Case "112"
        strMsgContens$ = M112
      Case "113"
        strMsgContens$ = M113
      Case "114"
        strMsgContens$ = M114
      Case "115"
        strMsgContens$ = M115
      Case "116"
        strMsgContens$ = M116
      Case "117"
        strMsgContens$ = M117
      Case "118"
        strMsgContens$ = M118
      Case "119"
        strMsgContens$ = M119
      Case "120"
        strMsgContens$ = M120
      Case "121"
        strMsgContens$ = M121
      Case "122"
        strMsgContens$ = M122
      Case "123"
        strMsgContens$ = M123
      Case "124"
        strMsgContens$ = M124
      Case "125"
        strMsgContens$ = M125
      Case "126"
        strMsgContens$ = M126
      Case "127"
        strMsgContens$ = M127
      Case "128"
        strMsgContens$ = M128
      Case "129"
        strMsgContens$ = M129
      Case "130"
        strMsgContens$ = M130
      Case "131"
        strMsgContens$ = M131
      Case "132"
        strMsgContens$ = M132
      Case "133"
        strMsgContens$ = M133
      Case "134"
        strMsgContens$ = M134
      Case "135"
        strMsgContens$ = M135
      Case "136"
        strMsgContens$ = M136
      Case "137"
        strMsgContens$ = M137
      Case "138"
        strMsgContens$ = M138
      Case "139"
        strMsgContens$ = M139
      Case "140"
        strMsgContens$ = M140
      Case "141"
        strMsgContens$ = M141
      Case "142"
        strMsgContens$ = M142
      Case "143"
        strMsgContens$ = M143
      Case "144"
        strMsgContens$ = M144
      Case "145"
        strMsgContens$ = M145
      Case "146"
        strMsgContens$ = M146
      Case "147"
        strMsgContens$ = M147
      Case "148"
        strMsgContens$ = M148
      Case "149"
        strMsgContens$ = M149
      Case "150"
        strMsgContens$ = M150
      Case "151"
        strMsgContens$ = M151
      Case "152"
        strMsgContens$ = M152
    End Select
    '显示消息
    If Len(Trim(strMsgContens$)) = 0 Then
        Message = MsgBox(Def_Message_Code_Error, vbCritical + vbOKOnly, App.Title)
    Else
        Message = MsgBox(strMsgContens$, vbInformation + vbOKOnly, App.Title)
    End If
  Case "E"
    Select Case Right(strMsgNo, 3)
      Case "001"
        strMsgContens$ = E001
      Case "002"
        strMsgContens$ = E002
      Case "003"
        strMsgContens$ = E003
      Case "004"
        strMsgContens$ = E004
      Case "005"
        strMsgContens$ = E005
      Case "006"
        strMsgContens$ = E006
      Case "007"
        strMsgContens$ = E007
      Case "008"
        strMsgContens$ = E008
      Case "009"
        strMsgContens$ = E009
      Case "010"
        strMsgContens$ = E010
      Case "011"
        strMsgContens$ = E011
      Case "012"
        strMsgContens$ = E012
      Case "013"
        strMsgContens$ = E013
      Case "014"
        strMsgContens$ = E014
      Case "015"
        strMsgContens$ = E015
      Case "016"
        strMsgContens$ = E016
      Case "017"
        strMsgContens$ = E017
      Case "018"
        strMsgContens$ = E018
      Case "019"
        strMsgContens$ = E019
      Case "020"
        strMsgContens$ = E020
      Case "021"
        strMsgContens$ = E021
      Case "022"
        strMsgContens$ = E022
      Case "023"
        strMsgContens$ = E023
      Case "024"
        strMsgContens$ = E024
      Case "025"
        strMsgContens$ = E025
      Case "026"
        strMsgContens$ = E026
      Case "027"
        strMsgContens$ = E027
      Case "028"
        strMsgContens$ = E028
      Case "029"
        strMsgContens$ = E029
      Case "030"
        strMsgContens$ = E030
      Case "031"
        strMsgContens$ = E031
      Case "032"
        strMsgContens$ = E032
      Case "033"
        strMsgContens$ = E033
      Case "034"
        strMsgContens$ = E034
      Case "035"
        strMsgContens$ = E035
      Case "036"
        strMsgContens$ = E036
      Case "037"
        strMsgContens$ = E037
      Case "038"
        strMsgContens$ = E038
      Case "039"
        strMsgContens$ = E039
      Case "040"
        strMsgContens$ = E040
      Case "041"
        strMsgContens$ = E041
      Case "042"
        strMsgContens$ = E042
      Case "043"
        strMsgContens$ = E043
      Case "044"
        strMsgContens$ = E044
      Case "045"
        strMsgContens$ = E045
      Case "046"
        strMsgContens$ = E046
      Case "047"
        strMsgContens$ = E047
      Case "048"
        strMsgContens$ = E048
      Case "049"
        strMsgContens$ = E049
      Case "050"
        strMsgContens$ = E050
      Case "051"
        strMsgContens$ = E051
      Case "052"
        strMsgContens$ = E052
      Case "053"
        strMsgContens$ = E053
      Case "054"
        strMsgContens$ = E054
      Case "055"
        strMsgContens$ = E055
      Case "056"
        strMsgContens$ = E056
      Case "057"
        strMsgContens$ = E057
      Case "058"
        strMsgContens$ = E058
      Case "059"
        strMsgContens$ = E059
      Case "060"
        strMsgContens$ = E060
      Case "061"
        strMsgContens$ = E061
      Case "062"
        strMsgContens$ = E062
      Case "063"
        strMsgContens$ = E063
      Case "064"
        strMsgContens$ = E064
      Case "065"
        strMsgContens$ = E065
      Case "066"
        strMsgContens$ = E066
      Case "067"
        strMsgContens$ = E067
      Case "068"
        strMsgContens$ = E068
      Case "069"
        strMsgContens$ = E069
      Case "070"
        strMsgContens$ = E070
      Case "071"
        strMsgContens$ = E071
      Case "072"
        strMsgContens$ = E072
      Case "073"
        strMsgContens$ = E073
      Case "074"
        strMsgContens$ = E074
      Case "075"
        strMsgContens$ = E075
      Case "076"
        strMsgContens$ = E076
      Case "077"
        strMsgContens$ = E077
      Case "078"
        strMsgContens$ = E078
      Case "079"
        strMsgContens$ = E079
      Case "080"
        strMsgContens$ = E080
      Case "081"
        strMsgContens$ = E081
      Case "082"
        strMsgContens$ = E082
      Case "083"
        strMsgContens$ = E083
      Case "084"
        strMsgContens$ = E084
      Case "085"
        strMsgContens$ = E085
      Case "086"
        strMsgContens$ = E086
      Case "087"
        strMsgContens$ = E087
      Case "088"
        strMsgContens$ = E088
      Case "089"
        strMsgContens$ = E089
      Case "090"
        strMsgContens$ = E090
      Case "091"
        strMsgContens$ = E091
      Case "092"
        strMsgContens$ = E092
      Case "093"
        strMsgContens$ = E093
      Case "094"
        strMsgContens$ = E094
      Case "095"
        strMsgContens$ = E095
      Case "096"
        strMsgContens$ = E096
      Case "097"
        strMsgContens$ = E097
      Case "098"
        strMsgContens$ = E098
      Case "099"
        strMsgContens$ = E099
      Case "100"
        strMsgContens$ = E100
      Case "101"
        strMsgContens$ = E101
      Case "102"
        strMsgContens$ = E102
      Case "103"
        strMsgContens$ = E103
      Case "104"
        strMsgContens$ = E104
      Case "105"
        strMsgContens$ = E105
      Case "106"
        strMsgContens$ = E106
      Case "107"
        strMsgContens$ = E107
      Case "108"
        strMsgContens$ = E108
      Case "109"
        strMsgContens$ = E109
      Case "110"
        strMsgContens$ = E110
      Case "111"
        strMsgContens$ = E111
      Case "112"
        strMsgContens$ = E112
      Case "113"
        strMsgContens$ = E113
      Case "114"
        strMsgContens$ = E114
      Case "115"
        strMsgContens$ = E115
      Case "116"
        strMsgContens$ = E116
      Case "117"
        strMsgContens$ = E117
      Case "118"
        strMsgContens$ = E118
      Case "119"
        strMsgContens$ = E119
      Case "120"
        strMsgContens$ = E120
      Case "121"
        strMsgContens$ = E121
      Case "122"
        strMsgContens$ = E122
      Case "123"
        strMsgContens$ = E123
      Case "124"
        strMsgContens$ = E124
      Case "125"
        strMsgContens$ = E125
      Case "126"
        strMsgContens$ = E126
      Case "127"
        strMsgContens$ = E127
      Case "128"
        strMsgContens$ = E128
      Case "129"
        strMsgContens$ = E129
      Case "130"
        strMsgContens$ = E130
      Case "131"
        strMsgContens$ = E131
      Case "132"
        strMsgContens$ = E132
      Case "133"
        strMsgContens$ = E133
      Case "134"
        strMsgContens$ = E134
      Case "135"
        strMsgContens$ = E135
      Case "136"
        strMsgContens$ = E136
      Case "137"
        strMsgContens$ = E137
      Case "138"
        strMsgContens$ = E138
      Case "139"
        strMsgContens$ = E139
      Case "140"
        strMsgContens$ = E140
      Case "141"
        strMsgContens$ = E141
      Case "142"
        strMsgContens$ = E142
      Case "143"
        strMsgContens$ = E143
      Case "144"
        strMsgContens$ = E144
      Case "145"
        strMsgContens$ = E145
      Case "146"
        strMsgContens$ = E146
      Case "147"
        strMsgContens$ = E147
      Case "148"
        strMsgContens$ = E148
      Case "149"
        strMsgContens$ = E149
      Case "150"
        strMsgContens$ = E150
        
      Case "899"
        strMsgContens$ = E899
    End Select
    '显示消息
    If Len(Trim(strMsgContens$)) = 0 Then
        Message = MsgBox(Def_Message_Code_Error, vbCritical + vbOKOnly, App.Title)
    Else
        Message = MsgBox(strMsgContens$, vbCritical + vbOKOnly, App.Title)
    End If
  Case "Q"
    Select Case Right(strMsgNo, 3)
      Case "001"
        strMsgContens$ = Q001
      Case "002"
        strMsgContens$ = Q002
      Case "003"
        strMsgContens$ = Q003
      Case "004"
        strMsgContens$ = Q004
      Case "005"
        strMsgContens$ = Q005
      Case "006"
        strMsgContens$ = Q006
      Case "007"
        strMsgContens$ = Q007
      Case "008"
        strMsgContens$ = Q008
      Case "009"
        strMsgContens$ = Q009
      Case "010"
        strMsgContens$ = Q010
      Case "011"
        strMsgContens$ = Q011
      Case "012"
        strMsgContens$ = Q012
      Case "013"
        strMsgContens$ = Q013
      Case "014"
        strMsgContens$ = Q014
      Case "015"
        strMsgContens$ = Q015
      Case "016"
        strMsgContens$ = Q016
      Case "017"
        strMsgContens$ = Q017
      Case "018"
        strMsgContens$ = Q018
      Case "019"
        strMsgContens$ = Q019
      Case "020"
        strMsgContens$ = Q020
      Case "021"
        strMsgContens$ = Q021
      Case "022"
        strMsgContens$ = Q022
      Case "023"
        strMsgContens$ = Q023
      Case "024"
        strMsgContens$ = Q024
      Case "025"
        strMsgContens$ = Q025
      Case "026"
        strMsgContens$ = Q026
      Case "027"
        strMsgContens$ = Q027
      Case "028"
        strMsgContens$ = Q028
      Case "029"
        strMsgContens$ = Q029
      Case "030"
        strMsgContens$ = Q030
      Case "031"
        strMsgContens$ = Q031
      Case "032"
        strMsgContens$ = Q032
      Case "033"
        strMsgContens$ = Q033
      Case "034"
        strMsgContens$ = Q034
      Case "035"
        strMsgContens$ = Q035
      Case "036"
        strMsgContens$ = Q036
      Case "037"
        strMsgContens$ = Q037
      Case "038"
        strMsgContens$ = Q038
      Case "039"
        strMsgContens$ = Q039
      Case "040"
        strMsgContens$ = Q040
      Case "041"
        strMsgContens$ = Q041
      Case "042"
        strMsgContens$ = Q042
      Case "043"
        strMsgContens$ = Q043
      Case "044"
        strMsgContens$ = Q044
      Case "045"
        strMsgContens$ = Q045
      Case "046"
        strMsgContens$ = Q046
      Case "047"
        strMsgContens$ = Q047
      Case "048"
        strMsgContens$ = Q048
      Case "049"
        strMsgContens$ = Q049
      Case "050"
        strMsgContens$ = Q050
      Case "051"
        strMsgContens$ = Q051
      Case "052"
        strMsgContens$ = Q052
      Case "053"
        strMsgContens$ = Q053
      Case "054"
        strMsgContens$ = Q054
      Case "055"
        strMsgContens$ = Q055
      Case "056"
        strMsgContens$ = Q056
      Case "057"
        strMsgContens$ = Q057
      Case "058"
        strMsgContens$ = Q058
      Case "059"
        strMsgContens$ = Q059
      Case "060"
        strMsgContens$ = Q060
      Case "061"
        strMsgContens$ = Q061
      Case "062"
        strMsgContens$ = Q062
      Case "063"
        strMsgContens$ = Q063
      Case "064"
        strMsgContens$ = Q064
      Case "065"
        strMsgContens$ = Q065
      Case "066"
        strMsgContens$ = Q066
      Case "067"
        strMsgContens$ = Q067
      Case "068"
        strMsgContens$ = Q068
      Case "069"
        strMsgContens$ = Q069
      Case "070"
        strMsgContens$ = Q070
      Case "071"
        strMsgContens$ = Q071
      Case "072"
        strMsgContens$ = Q072
      Case "073"
        strMsgContens$ = Q073
      Case "074"
        strMsgContens$ = Q074
      Case "075"
        strMsgContens$ = Q075
      Case "076"
        strMsgContens$ = Q076
      Case "077"
        strMsgContens$ = Q077
      Case "078"
        strMsgContens$ = Q078
      Case "079"
        strMsgContens$ = Q079
      Case "080"
        strMsgContens$ = Q080
      Case "081"
        strMsgContens$ = Q081
      Case "082"
        strMsgContens$ = Q082
      Case "083"
        strMsgContens$ = Q083
      Case "084"
        strMsgContens$ = Q084
      Case "085"
        strMsgContens$ = Q085
      Case "086"
        strMsgContens$ = Q086
      Case "087"
        strMsgContens$ = Q087
      Case "088"
        strMsgContens$ = Q088
      Case "089"
        strMsgContens$ = Q089
      Case "090"
        strMsgContens$ = Q090
      Case "091"
        strMsgContens$ = Q091
      Case "092"
        strMsgContens$ = Q092
      Case "093"
        strMsgContens$ = Q093
      Case "094"
        strMsgContens$ = Q094
      Case "095"
        strMsgContens$ = Q095
      Case "096"
        strMsgContens$ = Q096
      Case "097"
        strMsgContens$ = Q097
      Case "098"
        strMsgContens$ = Q098
      Case "099"
        strMsgContens$ = Q099
      Case "100"
        strMsgContens$ = Q100
      Case "101"
        strMsgContens$ = Q101
      Case "102"
        strMsgContens$ = Q102
      Case "103"
        strMsgContens$ = Q103
      Case "104"
        strMsgContens$ = Q104
      Case "105"
        strMsgContens$ = Q105
      Case "106"
        strMsgContens$ = Q106
      Case "107"
        strMsgContens$ = Q107
      Case "108"
        strMsgContens$ = Q108
      Case "109"
        strMsgContens$ = Q109
      Case "110"
        strMsgContens$ = Q110
      Case "111"
        strMsgContens$ = Q111
      Case "112"
        strMsgContens$ = Q112
      Case "113"
        strMsgContens$ = Q113
      Case "114"
        strMsgContens$ = Q114
      Case "115"
        strMsgContens$ = Q115
      Case "116"
        strMsgContens$ = Q116
      Case "117"
        strMsgContens$ = Q117
      Case "118"
        strMsgContens$ = Q118
      Case "119"
        strMsgContens$ = Q119
    End Select
    
    '显示消息
    If Len(Trim(strMsgContens$)) = 0 Then
        Message = MsgBox(Def_Message_Code_Error, vbCritical + vbOKOnly, App.Title)
    Else
        Message = MsgBox(strMsgContens$, vbQuestion + vbYesNo + nDefaultButton, App.Title)
    End If
    
  Case Else
    '显示消息
    Message = MsgBox(Def_Message_Code_Error, vbCritical + vbOKOnly, App.Title)
End Select

End Function

