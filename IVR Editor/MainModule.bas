Attribute VB_Name = "MainMod"
''--------------------------------------------------------------------------------------
'' Visual Editor Revision Notes:
''
'' 版本                    发布时间                 发布人          备注
''
'' Version 6.10.5         2014-01-29                Tony
''                        1、节点062（发起会议）的增加属性var_waitansto（等待坐席应答超时变量），以增强waitansto（应答超时）控制
''
'' Version 6.10.4         2012-11-23                Tony
''                        1、节点096（异步通信）增加属性：是否收到通信响应是否继续异步放音（carryonasynplay）
''                        2、节点006（无条件转移）增加属性：延时（sleep），单位为毫秒
''
'' Version 6.10.3         2012-06-26                Tony
''                        1、节点090（外拨）增加属性：设置主叫号码（setANI）
''
'' Version 6.10.2 Update_1  2012-05-07              Tony
''                        1、文件（File）主菜单中增加“最近访问文件”（MRU）菜单
''                        2、增加挂机入口节点，在执行挂机操作前系统自动转入该节点（通常用于收尾工作）
''
'' Version 6.10.1 Update_1  2012-04-18              Tony
''                        1、节点062（增强转坐席）改为“发起会议”
''                        2、节点040（建立留言）：增加属性“通知PL录音系统”，如果选中此项，则留言开始时发录音启动消息给PCS；留言结束时发录音停止消息；
''                        3、节点069（转虚拟分机）：增加属性“最大尝试次数”和“尝试间隔”
''
''
'' Version 6.9.8 Update_1  2012-01-17              Tony
''                        节点090（外拨）增加属性：
''                        1、外拨前是否强制摘机
''
'' Version 6.9.7 Update_1  2009-7-24               Tony
''                        节点040（录音）增加属性：
''                        1、是否录音开始提示（Tone）
''                        2、录音过程中提示（Tone）间隔
''                        3、录音实际时长记录变量(秒)
''
'' Version 6.9.6 Update_3  2008-8-29                Michael
''                        1,[Improved]流程页面支持鼠标滚轮
''                        2,[Fixed 2008-9-9]修复导出包含英文逗号的资源项目时产生错误CSV文件,导致导入错误的问题
''                        3,[Fixed 2008-10-7]修复添加资源项时不同语言的资源编号不能重复的问题.
''
'' Version 6.9.6 Update_2  2008-8-28                Michael
''                        1,[Fixed]修改打开多个流程后,流程参数->"设置"中字段为最后一次打开的流程设置
''
'' Version 6.9.6 Update_1  2008-8-21                Michael
''                        1,[Fixed]修复"查找语音资源"对话框不显示数字文件夹名称
''
'' Version 6.9.6          2008-08-20                Michael
''                       1,[Improved]删除变量后同时清空变量结构
''                       2,[Improved]修复打开流程后,不能立刻删除流程中的全部变量
''                       3,[Improved]修复在"选项"中设置"显示系统节点"后,被现实的系统节点标签未同步显示的问题
''
'' Version 6.9.5          2008-07-08                Michael
''                       1,[NEW]添加日志模块, VE在运行中的错误和操作被记入日志文件
''
'' Version 6.9.4 Update1  2008-06-30               Michael
''                       1,[FIX]修复运行在Oracle数据库环境中时不能新建资源项目的问题
''                       2,[FIX]重新编写判断是否是新增资源项目的代码,判断依据是在记录集中比较
''                              数据而不是和资源列表中的数据进行对比
''                       3,[FIX]重新编写删除资源列表中数据后设置选中行的代码.规则如下
''                              3.1:删除一条资源项时,如果该项目不是最后一项则选中行移至下一行
''                              3.2:删除一条资源项时,如果该项目是最后一项则选中行移至上一行
''                              3.3:同时删除多条资源时,则选中行移至被删除的最后一条资源的下一行
''
'' Version 6.9.4          2008-05-27               Michael
''                        转接座席(060),转接座席组(061)节点增加编号长度设置项, 取值范围0-10;0:不限制;默认4位,
''
'' Version 6.9.3          2008-05-16              Tony
''                        节点070（路由点查询）增加9-18查询参数
''
'' Version 6.9.2 Update2  2008-4-25                  Michael
''                      1,[FIX]修复修改流程属性的流程名称后,界面标题栏文字未及时更新的问题
'' Version 6.9.2 Update1  2008-4-24                 Michael
''                      1,[FIX]修复流程列表里删除流程时只删除第一行而不是选中行
''                      2,[FIX]修复从节点变量列表中转到变量列表是,选中行总是第一行而不是节点变量的所在行
'' Version 6.9.2       2008-2-19                Tony Sun
''                      1、修改节点日志默认值：
''                          007身份验证: 3
''                          008修改口令: 3
''                          022放音等待按键: 2
''                      2、流程里加一个全局开关，控制是否写节点日志(SData2_000.LogSwitch)，在“流程属性”->“流程参数”->“设置”中修改
''                      3、在节点清单中增加“日志”一项，可排序，如果节点“日志”属性、而且大于0则显示日志级别
''
'' Version 6.9.1       2008-2-18                Tony Sun
''                      1、流程导入、导出支持Tag(节点标签)字段；
''                      2、将所有图片资源包括在Exe中，不再需要BMP目录
''                      3、初次运行时，如果VoxPath不存在且创建失败时(比如指定到不存在的分区)，会弹出“不期望的错误”提示
''
'' Version 6.9.0       2008-1-25 ~ 2008-1-31     Michael
''                      1, [2008-1-25 ~ 29] 节点列表增加"统计标签"列,用于显示/编辑节点统计标签
''                      2, [2008-1-30 ~ 31] 所有节点界面增加"节点标签"按钮,用于编辑节点统计标签
''
'' Version 6.8.8       2007-11-29                Michael
''                      1, [2007-11-29] 增加TTS设置界面                  ---- New
''                      2, [2007-12-10] 选项卡中添加使用默认页面大小按钮   ---- New
''
'' Version 6.8.7       2007-11-26 ~ 2007-11-28   Michael
''                      1, 页面增加翻页按钮                     --- New
''                      2, 修改61节点的变量选择列表框            ---  Modified
''                      3, 国际化6.6.5之后发布版本,和老版本兼容  --- Improved
''                      4, 复制包含资源的项目时,可以让用户修改   --- Improved
''                          资源的前缀和后缀名
''                      5, 禁止资源号为0的资源被创建             --- Improved
''
'' Version 6.8.6       2007-11-01              Michael
''                      1, 更新Vox播放控件,解决程序崩溃           --- Improved
''                      2, 资源修改界面  -> 启用工具栏按钮        --- Fixed
''                      3, 录音界面      -> 优化状体标签         --- Improved
''                      4, 资源修改界面  -> 复制或新增资源时如果ID=0出现提示  --- Improved
''                      5, 播放语音文件后缀判别代码优化           --- Improved
''
'' Version 6.8.5       2007-10-20              Tony Sun
''                      6. 修改版本号
''                       . Debug：在新窗口打开页面后，原窗口会被锁住
''                       . 画面启动/切换慢，是因为ShowMarginLine中调用F_GetPrintPageScale查找打印机所致，
''                          修改ShowMarginLine过程，增加两个全局变量PageHeight，PageWidth，
''                          仅在程序启动以及选择打印时更改。
''
'' Version 6.8.0      July,18,2007 ~ Jul,24,2007    Michael
''                      1,项目管理列表,双击后进入此项目资源       --- New
''                      2,资源菜单 -> 打印 功能实现              --- New
''                      3,资源编辑 -> TTS批量语音合成            --- New
''                      4,资源编辑 -> 资源文件录音               --- New
''                      5,资源编辑 -> 播放资源文件               --- Modify,调用Windows API
''
'' Version 6.7        July,16,2007 ~ Jul,17,2007     Michael
''                      1,完成项目管理菜单,实现项目列表工具栏.
''                        方向按钮,新增按钮,更新按钮             --- New
''                      2,资源清单 - 删除资源                    --- Modify
''                      3,资源清单 - 查询                        --- New
''                      4,Node61播报方式添加 (秒)               --- Modify
''
'' Version 6.6.5        2007-06                 Michael
''                      1、[2007-06-25 Fix]选项-常规,增加语音文件格式,留言文件格式选项；
''                                         节点040（建立留言）增加留言文件分类选项；
''                                         节点040(建立留言)录音时长,最大时长由原来的255秒限制该为65535秒(18小时)
''                      2、[2007-07-01 Fix]节点061(转接座席组),新增usevar,switchtype,waitmethod,nd_wait 4个属性
''                      3、[2007-07-09 Fix]节点061(转接座席组),新增readEWT,var_userid,var_loginid  3个属性
''
''
'' Version 6.6          2007-03                 孙宝石
''                      1、[2007-03-20 New]节点022（放音等待按键）增加MaxTryTime属性；
''                                         设置COM时，如果没有指定主COM接口则提示；
''                                         节点040（建立留言）增加MinRecLength属性；
''                      2、[2007-03-30 New]与EasyResource合并
''                      3、[2007-04-08 New]修改资源导出文件(csv)格式，分三段：
''                              Section1: IVR服务加载资源列表，与以前导出文件兼容，3字段用逗号分割
''                                      L_ID, R_ID, R_Path
''                              Section2: 资源项目信息，10字段（比tbRefIVR表多一个Flag字段）用制表键分割
''                                      Flag(始终为-1)
''                              Section3: 资源详细清单，10字段（比tbResource表多一个Flag字段）用制表键分割
''                                      Flag(始终为-2)
''                      4、[2007-04-12 New]修改节点070（查询路由点）增加var_result属性，增加EWT查询参数选择
''                                         修改节点071（查询座席状态），querytype属性中增加查询工号状态选择
''                      5、[2007-04-18 New]增加TTS合成功能，以便于快速创建、测试流程
''
'' Version 6.5          2006-05-24              孙宝石
''                      1、[2005-11-10 New]多流程编辑功能（MDI）
''                      2、[2006-01-06 Fix]修正在“选项”中设置资源项目后没有保存到节点中的错误；
''                         [2006-01-06 Fix]修正在工作日设、定节点（010）年份不能显示去年的问题，该问题会导致旧的日期设定混乱；
''                      3、[2006-01-19 Fix]修正新建节点时节点编号取值错误，修正导出流程文件时流程ID清0错误；
''                      4、[2006-01-27 Fix]修正更新数据库连接参数时在当前工作页面打开流程连接旧数据源的问题；
''                      5、[2006-02-06 New]更新察看留言节点041，增加留言存档功能，
''                      可按全部、有效（新的和存档的留言）、新留言（缺省）、存档留言、删除留言分别察看，
''                      并可转换类型新留言->存档/删除；存档<->删除；
''                         [2006-02-06 New]复制节点时自动更新相关节点的连接；
''                      6、[2006-02-10 New]更新节点023：放音转移，增加变量放音功能；
''                      7、[2006-04-21 Fix]切换流程视图时应同时切换根、资源、语言、页码等信息；
''                      8、[2006-04-26 Fix]纠正工作日设定节点（010）中，月历只显示5周的问题；
''                      9、[2006-05-24 Fix]纠正新增变量：节点001，002，不能马上显示而必须重新打开流程的问题（原因：新建变量节点下标与节点编号不一致）；
''                      10、[2006-06-20 Fix]记录数据库时如果字节长度不足，则用0代替
''                      11、[2006-12-31 Fix]增加传真资源接收（050）、发送接点（055）；
''                      12、[2007-02-28 New]支持语音信箱功能（留言灯、分机邮箱、话务员邮箱）；
''                          修改 - 节点040：建立留言
''                          修改 - 节点041：察看留言
''
'' Version 6.2.0        2005-10-30              孙宝石
''                      1、[2005-08-18]修改数据库访问方式
''                      2、[2005-09-14]留言节点（040）增加最大静音时长(秒)
''                      3、[2005-10-25]节点多选功能
''                      4、[2005-10-30]多节点复制、删除、移动功能
''
'' Version 6.1.9        2005-08-05              孙宝石
''                      1、增加“流程同步”功能
''                      2、节点069,070增加使用变量功能
''                      3、节点070增加队列、组、小组、座席查询功能
''
'' Version 6.1.8        2005-07-15              孙宝石
''                      1、修改建立留言节点（040），增加“录音附加字段1，2，3”和“录音文件名记录变量”属性
''
'' Version 6.1.7        2005-06-27              孙宝石
''                      1、增加查询路由点状态节点（070）
''
'' Version 6.1.6        2005-06-20              孙宝石
''                      1、节点016（条件分支）增加字符串长度判断
''                      2、节点021（放音继续）增加日期类型（YYYYMMDD）
''
'' Version 6.1.5        2005-05-26              孙宝石
''                      1、增加Caling Card节点（091）
''                      2、变量定义节点中去掉类型设定，缺省为字符型（frm_buffer）
''
'' Version 6.1.4        2005-05-10              孙宝石
''                      1、复制家点时备注同时复制
''
'' Version 6.1.3        2005-03-15              孙宝石
''                      1、修改"外拨节点"
''                      2、增加"异步通信节点"(096)
''
'' Version 6.1.2        2005-03-10              孙宝石
''                      1、frmOptions中，页面数量控制增加验证
''
'' Version 6.1.1        2004-12-30              孙宝石
''                      1、增加节点028（TTS放音）
''                      2、增加条件编译参数Language，0 － 中文；1 － 英文
''                      3、在工具栏中增加节点连线
''                      4、将‘无操作’、‘发送数据’由控制类移动到系统类中
''                      5、增加节点016（条件分支）
''                      6、修改节点061，增加ACD、RoutePoint查询
''                      7、增强节点018（发送数据），每个流程可设定自己缺省的数据格式，转接前亦可设定特定格式
''                      8、将所有变量使用、设定组合框改为变量名列表
''                      9、增强节点102（变量修改），增加字符串处理
''
'' Version 6.0.1        2004-08-01              孟勇
''                      1、支持SQL和Oracle数据库
''                      2、通过条件编译Programnew控制节点数据存储格式，0 － 多数据项；1 － 单数据项
''
''--------------------------------------------------------------------------------------

Option Explicit

'' 流程版本，Sun added 2007-03-25
Public Const Def_CallFlow_MajorVersion = 2
Public Const Def_CallFlow_MinorVersion = 0

Sub Main()
On Error Resume Next

    'Michael Added @ 2008-7-1 for Error Log Module
    LogFile
    setLogfile
    WriteLogMessage 0, enu_Information, App.Title & " Started...", "Version:" & App.Major & "." & App.Minor & "." & App.Revision
    
    ' System Initialize
    SysInit
    
    '' Sun added 2012-05-07
    ''' Read MRU Data
    Call ReadMRUMData
    
    ' Show Welcome Infomation
    frmSplash.Show
    
    ' Show Main From
    frmMain.Show
    
    '' Sun added 2007-10-20
    ''' 打印机默认纸张尺寸
    Call F_GetPrintPageScale(gSystem.intPageWidth, gSystem.intPageHeight)
    
End Sub

'' Sun added 2002-04-02
' Call Me when gintSoundResourceID is changed
'
Public Sub SoundResourceIDChanged()

With frmMain
    If gintSoundResourceID > 0 Then
        .tbToolBar.Buttons("拨放语音").Enabled = True
        .tbToolBar.Buttons("停止放音").Enabled = True
        .tbToolBar.Buttons("拨放语音").ToolTipText = LoadNationalResString(1523) & " " & Trim(Str(gintSoundResourceID))
    Else
        .tbToolBar.Buttons("拨放语音").Enabled = False
        .tbToolBar.Buttons("停止放音").Enabled = False
        .tbToolBar.Buttons("拨放语音").ToolTipText = LoadNationalResString(1523)
    End If
End With
    
    
End Sub

Public Sub SetMainFormItemsEnableWhenPropertyShow(ByVal f_Enabled As Boolean)
    
    Call frmMain.SetActiveFormItemsEnable(f_Enabled)

End Sub
