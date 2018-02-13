Attribute VB_Name = "BitOpera"
 Option Explicit
  '说明----------------------------------------
  '增强   vb   的位操作功能的模块，主要包含
  '有左右移位，取字节，字节连接等通用例程
  '兼容性：VB5.0   ,6.0
  '--------------------------------------------
  'api函数     拷贝内存
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)
   
   Public bBorrowFlag As Byte
   
  '-----------------------下面这些例程实现整型变量的拆分，合并操作-------------
  
  
  Public Function Con(ByVal HiByte As Byte, ByVal LoByte As Byte) As Integer
    '把两个字节   (Byte)   连成一个字   （word）
    'INPUT--------------------------------------------------------------------
            'HiByte             参与连结的高字节
            'LoByte             参与连结的低字节
    'OUTPUT-------------------------------------------------------------------
            '返回值             连结的结果
    Dim iRet     As Integer
     
    '用到的函数   varptr()   说明：取一个变量的地址。
    CopyMemory ByVal VarPtr(iRet), LoByte, 1
    CopyMemory ByVal VarPtr(iRet) + 1, HiByte, 1
     
    Con = iRet
     
  End Function
   
  Public Function ConWord(ByVal HiWord As Integer, ByVal LoWord As Integer) As Long
  '把两个字（Word）连成一个双字（DWord）
  'INPUT--------------------------------------------------------------------
          'HiWord             参与连结的高位字
          'LoWord             参与连结的低位字
  'OUTPUT-------------------------------------------------------------------
          '返回值             连结的结果
  Dim lRet     As Long
   
  CopyMemory ByVal VarPtr(lRet), LoWord, 2
  CopyMemory ByVal VarPtr(lRet) + 2, HiWord, 2
   
  ConWord = lRet
   
  End Function
   
  Public Function Hi(ByVal Word As Integer) As Byte
  '取一个字（Word）的高字节（Byte）
  'INPUT-------------------------------------------
          'Word             字（Word）
  'OUTPUT------------------------------------------
          '返回值           Word参数的高字节
  Dim bytRet     As Byte
   
  CopyMemory bytRet, ByVal VarPtr(Word) + 1, 1
   
  Hi = bytRet
   
  End Function
   
  Public Function Lo(ByVal Word As Integer) As Byte
  '取一个字（Word）的低字节（Byte）
  'INPUT-------------------------------------------
          'Word             字（Word）
  'OUTPUT------------------------------------------
          '返回值           Word参数的低字节
  Dim bytRet     As Byte
   
  CopyMemory bytRet, ByVal VarPtr(Word), 1
   
  Lo = bytRet
   
  End Function
   
  Public Function HiWord(ByVal DWord As Long) As Integer
  '取一个双字（DWord）的高位字
  'INPUT-------------------------------------------
          'DWord             双字
  'OUTPUT------------------------------------------
          '返回值           DWord参数的高位字
  Dim intRet     As Integer
   
  CopyMemory intRet, ByVal VarPtr(DWord) + 2, 2
   
  HiWord = intRet
   
  End Function
   
  Public Function LoWord(ByVal DWord As Long) As Integer
  '取一个双字（DWord）的低位字
  'INPUT-------------------------------------------
          'DWord             双字
  'OUTPUT------------------------------------------
          '返回值           DWord参数的低位字
  Dim intRet     As Integer
   
  CopyMemory intRet, ByVal VarPtr(DWord), 2
   
  LoWord = intRet
   
  End Function
  
   
  '-------------------------下面这些例程实现整形变量的移位-------------------
   
  Public Function ShLB(ByVal Byt As Byte, Optional ByVal BitsNum As Long = 1) As Byte
  '字节的左移函数
  'INPUT-----------------------------
          'Byt   源操作数
          'BitsNum   移位的位数
  'OUTPUT----------------------------
          '返回值     移位结果
  Dim i&
   
  For i = 1 To BitsNum
          Byt = ShLB_By1Bit(Byt)
  Next i
   
  ShLB = Byt
   
  End Function
   
  Public Function ShRB(ByVal Byt As Byte, Optional ByVal BitsNum As Long = 1) As Byte
  '字节的右移函数
  'INPUT-----------------------------
          'Byt   源操作数
          'BitsNum   移位的位数
  'OUTPUT----------------------------
          '返回值     移位结果
  'last   updated   by   Liu   Qi   2004-3-23
  Dim i&
   
  For i = 1 To BitsNum
          Byt = ShRB_By1Bit(Byt)
  Next i
   
  ShRB = Byt
   
  End Function
   
  Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte
  '把字节左移一位的函数,为   ShlB   服务.
  'INPUT-----------------------------
          'Byt   源操作数
  'OUTPUT----------------------------
          '返回值     移位结果
     
  '(Byt   And   &H7F):   屏蔽最高位.     *2:左移一位
  ShLB_By1Bit = (Byt And &H7F) * 2
   
  'ShlB_By1Bit   =   Byt   *   2'溢出测试
   
  End Function
  Private Function ShRB_By1Bit(ByVal Byt As Byte) As Byte
  '把字节右移一位的函数,为   ShrB   服务.
  'INPUT-----------------------------
          'Byt   源操作数
  'OUTPUT----------------------------
          '返回值     移位结果
     
  '/2:右移一位
  ShRB_By1Bit = Fix(Byt / 2)
   
  End Function
   
   
  Public Function ShLW(ByVal Word As Integer, Optional ByVal BitsNum As Long = 1) As Integer
  '字的左移函数
  'INPUT-------------------------------
          'Word   源操作数
          'BitsNum   移位的位数
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim i&
   
  For i = 1 To BitsNum
          Word = ShLW_By1Bit(Word)
  Next i
   
  ShLW = Word
   
  End Function
   
  Public Function ShRW(ByVal Word As Integer, Optional ByVal BitsNum As Long = 1) As Integer
  '字的右移函数
  'INPUT-------------------------------
          'Word   源操作数
          'BitsNum   移位的位数
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim i&
   
  For i = 1 To BitsNum
          Word = ShRW_By1Bit(Word)
  Next i
   
  ShRW = Word
  End Function
  Private Function ShLW_By1Bit(ByVal Word As Integer) As Integer
  '把一个字左移一位的函数
  'INPUT-------------------------------
          'Word   源操作数
           
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim HiByte     As Byte, LoByte       As Byte
   
  '把字拆分为字节
  HiByte = Hi(Word):       LoByte = Lo(Word)
  '把高字节左移一位,保证把低字节的最高位移入高字节的最低位
  HiByte = ShLB_By1Bit(HiByte) Or IIf((LoByte And &H80) = &H80, &H1, &H0)
  LoByte = ShLB_By1Bit(LoByte)       '低字节左移一位
  '把移位后的字节再重新组合成字
  ShLW_By1Bit = Con(HiByte, LoByte)
   
  End Function
   
  Private Function ShRW_By1Bit(ByVal Word As Integer) As Integer
  '把一个字右移一位的函数
  'INPUT-------------------------------
          'Word   源操作数
           
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim HiByte     As Byte, LoByte       As Byte
   
  '把字拆分为字节
  HiByte = Hi(Word):       LoByte = Lo(Word)
   
  '低字节右移一位,保证把高字节的最低位移入低字节的最高位
  LoByte = ShRB_By1Bit(LoByte) Or IIf((HiByte And &H1) = &H1, &H80, &H0)
   
  '把高字节右移一位,
  HiByte = ShRB_By1Bit(HiByte)
   
  '把移位后的字节再重新组合成字
  ShRW_By1Bit = Con(HiByte, LoByte)
   
   
  End Function
   
   
  Public Function ShLD(ByVal DWord As Long, Optional ByVal BitsNum As Long = 1) As Long
  '把一个双字左移的函数
  'INPUT-------------------------------
          'DWord   源操作数
          'BitsNum   移位的位数
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim i&
   
  For i = 1 To BitsNum
          DWord = ShLD_By1Bit(DWord)
  Next i
   
  ShLD = DWord
   
  End Function
   
  Public Function ShRD(ByVal DWord As Long, Optional ByVal BitsNum As Long = 1) As Long
  '把一个双字右移的函数
  'INPUT-------------------------------
          'DWord   源操作数
          'BitsNum   移位的位数
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim i&
   
  For i = 1 To BitsNum
          DWord = ShRD_By1Bit(DWord)
  Next i
   
  ShRD = DWord
  End Function
  Public Function ShLD_By1Bit(ByVal DWord As Long) As Long
  '把一个双字左移一位的函数，为   ShlD()   服务
  'INPUT-------------------------------
          'DWord   源操作数
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim iHiWord%, iLoWord%
   
  '把双字拆分为两个单字
  iHiWord = HiWord(DWord):       iLoWord = LoWord(DWord)
   
  '高位字左移一位,要把低位字的最高位移到高位字的最低位
  iHiWord = ShLW_By1Bit(iHiWord) Or IIf((iLoWord And &H8000) = &H8000, &H1, &H0)
   
  '低位字左移一位
  iLoWord = ShLW_By1Bit(iLoWord)
   
  ShLD_By1Bit = ConWord(iHiWord, iLoWord)         '重新连接成双字返回结果
   
  End Function
   
  Public Function ShRD_By1Bit(ByVal DWord As Long) As Long
  '把一个双字右移一位的函数,为   ShrD()   服务
  'INPUT-------------------------------
          'DWord   源操作数
  'OUTPUT------------------------------
            '返回值     移位结果
  Dim iHiWord%, iLoWord%
   
  '把双字拆分为两个单字
  iHiWord = HiWord(DWord):       iLoWord = LoWord(DWord)
   
  '把低位字右移一位，要把高位字的最低位移到低位字的最高位
  iLoWord = ShRW_By1Bit(iLoWord) Or IIf((iHiWord And &H1) = &H1, &H8000, &H0)
   
  '把高位字右移一位
  iHiWord = ShRW_By1Bit(iHiWord)
   
  ShRD_By1Bit = ConWord(iHiWord, iLoWord)         '重新连接成双字返回结果
   
  End Function
   
  Public Function ShLB_C_By1Bit(ByVal Byt As Byte) As Byte
  '把字节<<循环>>左移一位的函数.C   表示   Cycle,循环
  'INPUT-----------------------------
          'Byt   :源操作数
  'OUTPUT----------------------------
          '返回值   :   移位结果
   
  '(Byt   And   &H7F):   屏蔽最高位.     *2:左移一位
  ShLB_C_By1Bit = ((Byt And &H7F) * 2) Or IIf((Byt And &H80) = &H80, &H1, &H0)
  End Function
   
  Public Function ShRB_C_By1Bit(ByVal Byt As Byte) As Byte
  '把字节<<循环>>右移一位的函数。
  'INPUT-----------------------------
          'Byt   :源操作数
  'OUTPUT----------------------------
          '返回值   :   移位结果
   
  '(Byt   And   &H7F):   屏蔽最高位.     *2:左移一位
  ShRB_C_By1Bit = Fix(Byt / 2) Or IIf((Byt And &H1) = &H1, &H80, &H0)
  End Function
   
  

