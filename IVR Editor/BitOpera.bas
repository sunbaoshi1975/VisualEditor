Attribute VB_Name = "BitOpera"
 Option Explicit
  '˵��----------------------------------------
  '��ǿ   vb   ��λ�������ܵ�ģ�飬��Ҫ����
  '��������λ��ȡ�ֽڣ��ֽ����ӵ�ͨ������
  '�����ԣ�VB5.0   ,6.0
  '--------------------------------------------
  'api����     �����ڴ�
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)
   
   Public bBorrowFlag As Byte
   
  '-----------------------������Щ����ʵ�����ͱ����Ĳ�֣��ϲ�����-------------
  
  
  Public Function Con(ByVal HiByte As Byte, ByVal LoByte As Byte) As Integer
    '�������ֽ�   (Byte)   ����һ����   ��word��
    'INPUT--------------------------------------------------------------------
            'HiByte             ��������ĸ��ֽ�
            'LoByte             ��������ĵ��ֽ�
    'OUTPUT-------------------------------------------------------------------
            '����ֵ             ����Ľ��
    Dim iRet     As Integer
     
    '�õ��ĺ���   varptr()   ˵����ȡһ�������ĵ�ַ��
    CopyMemory ByVal VarPtr(iRet), LoByte, 1
    CopyMemory ByVal VarPtr(iRet) + 1, HiByte, 1
     
    Con = iRet
     
  End Function
   
  Public Function ConWord(ByVal HiWord As Integer, ByVal LoWord As Integer) As Long
  '�������֣�Word������һ��˫�֣�DWord��
  'INPUT--------------------------------------------------------------------
          'HiWord             ��������ĸ�λ��
          'LoWord             ��������ĵ�λ��
  'OUTPUT-------------------------------------------------------------------
          '����ֵ             ����Ľ��
  Dim lRet     As Long
   
  CopyMemory ByVal VarPtr(lRet), LoWord, 2
  CopyMemory ByVal VarPtr(lRet) + 2, HiWord, 2
   
  ConWord = lRet
   
  End Function
   
  Public Function Hi(ByVal Word As Integer) As Byte
  'ȡһ���֣�Word���ĸ��ֽڣ�Byte��
  'INPUT-------------------------------------------
          'Word             �֣�Word��
  'OUTPUT------------------------------------------
          '����ֵ           Word�����ĸ��ֽ�
  Dim bytRet     As Byte
   
  CopyMemory bytRet, ByVal VarPtr(Word) + 1, 1
   
  Hi = bytRet
   
  End Function
   
  Public Function Lo(ByVal Word As Integer) As Byte
  'ȡһ���֣�Word���ĵ��ֽڣ�Byte��
  'INPUT-------------------------------------------
          'Word             �֣�Word��
  'OUTPUT------------------------------------------
          '����ֵ           Word�����ĵ��ֽ�
  Dim bytRet     As Byte
   
  CopyMemory bytRet, ByVal VarPtr(Word), 1
   
  Lo = bytRet
   
  End Function
   
  Public Function HiWord(ByVal DWord As Long) As Integer
  'ȡһ��˫�֣�DWord���ĸ�λ��
  'INPUT-------------------------------------------
          'DWord             ˫��
  'OUTPUT------------------------------------------
          '����ֵ           DWord�����ĸ�λ��
  Dim intRet     As Integer
   
  CopyMemory intRet, ByVal VarPtr(DWord) + 2, 2
   
  HiWord = intRet
   
  End Function
   
  Public Function LoWord(ByVal DWord As Long) As Integer
  'ȡһ��˫�֣�DWord���ĵ�λ��
  'INPUT-------------------------------------------
          'DWord             ˫��
  'OUTPUT------------------------------------------
          '����ֵ           DWord�����ĵ�λ��
  Dim intRet     As Integer
   
  CopyMemory intRet, ByVal VarPtr(DWord), 2
   
  LoWord = intRet
   
  End Function
  
   
  '-------------------------������Щ����ʵ�����α�������λ-------------------
   
  Public Function ShLB(ByVal Byt As Byte, Optional ByVal BitsNum As Long = 1) As Byte
  '�ֽڵ����ƺ���
  'INPUT-----------------------------
          'Byt   Դ������
          'BitsNum   ��λ��λ��
  'OUTPUT----------------------------
          '����ֵ     ��λ���
  Dim i&
   
  For i = 1 To BitsNum
          Byt = ShLB_By1Bit(Byt)
  Next i
   
  ShLB = Byt
   
  End Function
   
  Public Function ShRB(ByVal Byt As Byte, Optional ByVal BitsNum As Long = 1) As Byte
  '�ֽڵ����ƺ���
  'INPUT-----------------------------
          'Byt   Դ������
          'BitsNum   ��λ��λ��
  'OUTPUT----------------------------
          '����ֵ     ��λ���
  'last   updated   by   Liu   Qi   2004-3-23
  Dim i&
   
  For i = 1 To BitsNum
          Byt = ShRB_By1Bit(Byt)
  Next i
   
  ShRB = Byt
   
  End Function
   
  Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte
  '���ֽ�����һλ�ĺ���,Ϊ   ShlB   ����.
  'INPUT-----------------------------
          'Byt   Դ������
  'OUTPUT----------------------------
          '����ֵ     ��λ���
     
  '(Byt   And   &H7F):   �������λ.     *2:����һλ
  ShLB_By1Bit = (Byt And &H7F) * 2
   
  'ShlB_By1Bit   =   Byt   *   2'�������
   
  End Function
  Private Function ShRB_By1Bit(ByVal Byt As Byte) As Byte
  '���ֽ�����һλ�ĺ���,Ϊ   ShrB   ����.
  'INPUT-----------------------------
          'Byt   Դ������
  'OUTPUT----------------------------
          '����ֵ     ��λ���
     
  '/2:����һλ
  ShRB_By1Bit = Fix(Byt / 2)
   
  End Function
   
   
  Public Function ShLW(ByVal Word As Integer, Optional ByVal BitsNum As Long = 1) As Integer
  '�ֵ����ƺ���
  'INPUT-------------------------------
          'Word   Դ������
          'BitsNum   ��λ��λ��
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim i&
   
  For i = 1 To BitsNum
          Word = ShLW_By1Bit(Word)
  Next i
   
  ShLW = Word
   
  End Function
   
  Public Function ShRW(ByVal Word As Integer, Optional ByVal BitsNum As Long = 1) As Integer
  '�ֵ����ƺ���
  'INPUT-------------------------------
          'Word   Դ������
          'BitsNum   ��λ��λ��
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim i&
   
  For i = 1 To BitsNum
          Word = ShRW_By1Bit(Word)
  Next i
   
  ShRW = Word
  End Function
  Private Function ShLW_By1Bit(ByVal Word As Integer) As Integer
  '��һ��������һλ�ĺ���
  'INPUT-------------------------------
          'Word   Դ������
           
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim HiByte     As Byte, LoByte       As Byte
   
  '���ֲ��Ϊ�ֽ�
  HiByte = Hi(Word):       LoByte = Lo(Word)
  '�Ѹ��ֽ�����һλ,��֤�ѵ��ֽڵ����λ������ֽڵ����λ
  HiByte = ShLB_By1Bit(HiByte) Or IIf((LoByte And &H80) = &H80, &H1, &H0)
  LoByte = ShLB_By1Bit(LoByte)       '���ֽ�����һλ
  '����λ����ֽ���������ϳ���
  ShLW_By1Bit = Con(HiByte, LoByte)
   
  End Function
   
  Private Function ShRW_By1Bit(ByVal Word As Integer) As Integer
  '��һ��������һλ�ĺ���
  'INPUT-------------------------------
          'Word   Դ������
           
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim HiByte     As Byte, LoByte       As Byte
   
  '���ֲ��Ϊ�ֽ�
  HiByte = Hi(Word):       LoByte = Lo(Word)
   
  '���ֽ�����һλ,��֤�Ѹ��ֽڵ����λ������ֽڵ����λ
  LoByte = ShRB_By1Bit(LoByte) Or IIf((HiByte And &H1) = &H1, &H80, &H0)
   
  '�Ѹ��ֽ�����һλ,
  HiByte = ShRB_By1Bit(HiByte)
   
  '����λ����ֽ���������ϳ���
  ShRW_By1Bit = Con(HiByte, LoByte)
   
   
  End Function
   
   
  Public Function ShLD(ByVal DWord As Long, Optional ByVal BitsNum As Long = 1) As Long
  '��һ��˫�����Ƶĺ���
  'INPUT-------------------------------
          'DWord   Դ������
          'BitsNum   ��λ��λ��
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim i&
   
  For i = 1 To BitsNum
          DWord = ShLD_By1Bit(DWord)
  Next i
   
  ShLD = DWord
   
  End Function
   
  Public Function ShRD(ByVal DWord As Long, Optional ByVal BitsNum As Long = 1) As Long
  '��һ��˫�����Ƶĺ���
  'INPUT-------------------------------
          'DWord   Դ������
          'BitsNum   ��λ��λ��
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim i&
   
  For i = 1 To BitsNum
          DWord = ShRD_By1Bit(DWord)
  Next i
   
  ShRD = DWord
  End Function
  Public Function ShLD_By1Bit(ByVal DWord As Long) As Long
  '��һ��˫������һλ�ĺ�����Ϊ   ShlD()   ����
  'INPUT-------------------------------
          'DWord   Դ������
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim iHiWord%, iLoWord%
   
  '��˫�ֲ��Ϊ��������
  iHiWord = HiWord(DWord):       iLoWord = LoWord(DWord)
   
  '��λ������һλ,Ҫ�ѵ�λ�ֵ����λ�Ƶ���λ�ֵ����λ
  iHiWord = ShLW_By1Bit(iHiWord) Or IIf((iLoWord And &H8000) = &H8000, &H1, &H0)
   
  '��λ������һλ
  iLoWord = ShLW_By1Bit(iLoWord)
   
  ShLD_By1Bit = ConWord(iHiWord, iLoWord)         '�������ӳ�˫�ַ��ؽ��
   
  End Function
   
  Public Function ShRD_By1Bit(ByVal DWord As Long) As Long
  '��һ��˫������һλ�ĺ���,Ϊ   ShrD()   ����
  'INPUT-------------------------------
          'DWord   Դ������
  'OUTPUT------------------------------
            '����ֵ     ��λ���
  Dim iHiWord%, iLoWord%
   
  '��˫�ֲ��Ϊ��������
  iHiWord = HiWord(DWord):       iLoWord = LoWord(DWord)
   
  '�ѵ�λ������һλ��Ҫ�Ѹ�λ�ֵ����λ�Ƶ���λ�ֵ����λ
  iLoWord = ShRW_By1Bit(iLoWord) Or IIf((iHiWord And &H1) = &H1, &H8000, &H0)
   
  '�Ѹ�λ������һλ
  iHiWord = ShRW_By1Bit(iHiWord)
   
  ShRD_By1Bit = ConWord(iHiWord, iLoWord)         '�������ӳ�˫�ַ��ؽ��
   
  End Function
   
  Public Function ShLB_C_By1Bit(ByVal Byt As Byte) As Byte
  '���ֽ�<<ѭ��>>����һλ�ĺ���.C   ��ʾ   Cycle,ѭ��
  'INPUT-----------------------------
          'Byt   :Դ������
  'OUTPUT----------------------------
          '����ֵ   :   ��λ���
   
  '(Byt   And   &H7F):   �������λ.     *2:����һλ
  ShLB_C_By1Bit = ((Byt And &H7F) * 2) Or IIf((Byt And &H80) = &H80, &H1, &H0)
  End Function
   
  Public Function ShRB_C_By1Bit(ByVal Byt As Byte) As Byte
  '���ֽ�<<ѭ��>>����һλ�ĺ�����
  'INPUT-----------------------------
          'Byt   :Դ������
  'OUTPUT----------------------------
          '����ֵ   :   ��λ���
   
  '(Byt   And   &H7F):   �������λ.     *2:����һλ
  ShRB_C_By1Bit = Fix(Byt / 2) Or IIf((Byt And &H1) = &H1, &H80, &H0)
  End Function
   
  

