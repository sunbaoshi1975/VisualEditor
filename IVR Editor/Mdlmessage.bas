Attribute VB_Name = "mdlMessage"

'**************************************************************
'Error message defination
 
#If Language = 0 Then

 Const E001 = "��������������������������������ã�"
 Const E002 = "�޷����ӷ������������������ӣ�"
 Const E003 = "������Ӧ�������豸���������"
 Const E004 = "���ջ����������"
 Const E005 = "���仺����������"
 Const E006 = "Parity ����"
 Const E007 = "�˿ڲ�������������������ã�"
 Const E008 = "���ӳ�ʱ�����Ժ����ԣ�"
 Const E009 = "ϵͳ�ļ�������������һ����°�װ��"
 Const E010 = "�����������CapsLock�������������룡"
 Const E011 = "�Ƿ������ߣ���¼ȡ����"
 Const E012 = "�¿�����ȷ�Ͽ��һ�£����������롣"
 Const E013 = "�������Ϊ3-8���ַ������������롣"
 Const E014 = "�������޷�����������ɾ�����õļ�¼��"
 Const E015 = "���Ȱ�װ��ӡ����"
 Const E016 = "���ݿ��޷��򿪣���������Դ���ã�"
 Const E017 = "���ݿ��޷��رգ�ǿ����ֹ��"
 Const E018 = "���ݿ�������޷��ָ���ǿ����ֹ��"
 Const E019 = "��ȫ��֤���ݿ��������ϵͳ����Ա��ϵ��"
 Const E020 = "û�д˲����ߣ������ԣ�"
 Const E021 = "���������ݶ�ʧ���޷�ȥ�ظ�ԭ�����ݣ�"
 Const E022 = "�����޷�д��������������������Ӻ�������ã�"
 Const E023 = "�޷���ɴ�ӡ��ҵ�������ӡ�����á�"
 Const E024 = "�޷����ӷ������������������ӣ�"
 Const E025 = "����������С��32767��"
 Const E026 = "ɾ������ʧ�ܣ�"
 Const E027 = "��ŷ�Χ����ȷ�����������롣"
 Const E028 = "�޷��ӷ�������ȡ���ݣ������������Ӻ�������ã�"
 Const E029 = "�Ҳ���ָ����ŵ���Դ��"
 Const E030 = "�����ڴ�����µ���Դ����ѡ���ʵ��Ĳ���λ�ã�"
 Const E031 = "�������Ͳ���ȷ�����������롣"
 Const E032 = "�ڵ��Ų���<256�����������롣"
 Const E033 = "�ڵ��޷��༭��"
 Const E034 = "���ڵ㣺���������ݴ���255��С��32767��"
 Const E035 = "���ڵ㣺���������ݴ���255��С��32767��"
 Const E036 = "�������û����������С��256��"
 Const E037 = "�����롰��ʾ�û�������롱����IDС��32767��"
 Const E038 = "�����롰��ʾ�û�����������IDС��32767��"
 Const E039 = "�����롰��ʾ�û��������롱����IDС��32767��"
 Const E040 = "�����롰COM�ӿ�ID������0��С��32767��"
 Const E041 = "�����롰�ɹ�ת�ڵ�ID������255��С��32767��"
 Const E042 = "�����롰ʧ��ת�ڵ�ID������255��С��32767��"
 Const E043 = "�����롰�û������¼������0��С��256��"
 Const E044 = "�����롰�û������¼������0��С��256��"
 Const E045 = "�����롰��֤������־������0��С��256��"
 Const E046 = "�����롰��֤�����־������0��С��256��"
 Const E047 = "�����롰����Դ���������0��С��256��"
 Const E048 = "�����롰��������־������0��С��16��"
 Const E049 = "�����롰�û�������󳤶ȡ�����0��С��256��"
 Const E050 = "�����롰�û�������󳤶ȡ�����0��С��256��"
 Const E051 = "�����롰�ڵ㳬ʱ(S)������0��С��256��"
 Const E052 = "�����롰�޸Ľ����־������0��С��256��"
 Const E053 = "�����롰��ʾ�û������¿������IDС��32767��"
 Const E054 = "�����롰��ʾ�û��ٴ�ȷ�ϡ�����IDС��32767��"
 Const E055 = "�����롰���β�һ���������롱����IDС��32767��"
 Const E056 = "�����롰��ʾ�û��޸ĳɹ�������IDС��32767��"
 Const E057 = "��������ȷ��ʱ��ֵ��"
 Const E058 = "��ѡ��������͡�"
 Const E059 = "�����롰ʱ���ת�ƽڵ�ID������255��С��32767��"
 Const E060 = "ע�⣺��ǰ������δָ�����ڵ㣡���ڵ��Ǻ������̵���ʼ�ڵ㣬û�и��ڵ����̽��޷�������"
 Const E061 = "������ڵ���ʱʱ����"
 Const E062 = "�����롰����ѡ������¼������0��С��256��"
 Const E063 = "����ͬ��ʧ�ܣ���ȷ��Ŀ�����ݿ��Ƿ�ɷ��ʣ�"
 Const E064 = "��Դͬ��ʧ�ܣ���ȷ��Ŀ�����ݿ��Ƿ�ɷ��ʣ�"
 Const E065 = "�����롰��Ϣ��ת�ڵ�ID������255��С��32767��"
 Const E066 = "ʱ��δ������"
 Const E067 = "�����롰ת�ƽڵ�ID������255��С��32767��"
 Const E068 = "���ǺϷ�����Դ�ļ����ļ����ƻ������ܵ��룡"
 Const E069 = "��"
 Const E070 = "ʧ��ת��ڵ㣺���������ݴ���255��С��32767��"
 Const E071 = "�����롰���ͱ�����С��256��"
 Const E072 = "�ӽڵ㣺���������ݴ���255��С��32767��"
 Const E073 = "�������ݣ������ַ��������ȡ�"
 Const E074 = "�����롰�����ߡ���"
 Const E075 = "�����롰���ݰ����͡�С��256��"
 Const E076 = "�����롰����������С��32767��"
 Const E077 = "�����롰�������ȡ�С��256��"
 Const E078 = "���ȳ�����Χ��"
 Const E079 = "�������Ƴ��ȳ�����Χ��"
 Const E080 = "�����롰ʹ�ñ���ID��С��256��"
 Const E081 = "�����롰�������ȡ�С��256��"
 Const E082 = "�����롰�����������С��256��"
 Const E083 = "�����롰������¼��С��256��"
 'Const E084 = "�����롰¼��ʱ�䳤��С��256��"
 Const E084 = "�����롰¼��ʱ�䳤��С��65536(18 Hours)��"    'Michael Modified
 Const E085 = "�����롰��ϯ��ID��С��32767��"
 Const E086 = "�����롰�ȴ����Ŵ�����С��256��"
 Const E087 = "DLL�ļ�ID���ܴ���32767��"
 Const E088 = "COM�ļ�ID���ܴ���32767��"
 Const E089 = "��ѡ����Ҫ�򿪵����̱�š�"
 Const E090 = "�����롰ǰ����������ID��С��32767��"
 Const E091 = "�����롰������������ID��С��32767��"
 Const E092 = "�����롰��ʾ����ID��С��32767��"
 Const E093 = "�����롰�����ļ�ID��С��32767��"
 Const E094 = "�����롰LOGO�ļ�ID��С��32767��"
 Const E095 = "�����롰����ʽ�ļ�ID��С��32767��"
 Const E096 = "�����롰ת����ʾ����ID��С��32767��"
 Const E097 = "�����롰�ȴ�ѭ������ID��С��32767��"
 Const E098 = "�����롰û�ϰ���ʾ����ID��С��32767��"
 Const E099 = "�����롰��ϯæ��ʾ����ID��С��32767��"
 Const E100 = "�����롰����ת�ڵ�ID������255��С��32767��"
 Const E101 = "�����롰���������ԴID��С��32767��"
 Const E102 = "�����롰����������ԴID��С��32767��"
 Const E103 = "�����롰��¼����ID��С��256��"
 Const E104 = "�����롰TTS ������ԴID��С��32767��"
 Const E105 = "�����롰�����������ID��С��32767����"
 Const E106 = "�����롰�Ⲧ��š�С��256��"
 Const E107 = "�����롰�������Ч��ʾ������С��32767��"
 Const E108 = "�����롰�����ʱ����С��60��"
 Const E109 = "�����롰��С¼��ʱ����С��180��"
 Const E110 = "�����롰ת�ƽڵ�ID������255��С��32767��"
 Const E111 = "�����롰û���ϰ�ڵ�ID������255��С��32767��"
 Const E112 = "�����롰������æת�ڵ�ID������255��С��32767��"
 Const E113 = "�����롰ת�ӳɹ�ת�ڵ�ID������255��С��32767��"
 Const E114 = "�����롰����ֻ��š�С�� 2,147,483,647��"
 Const E115 = "��������ȷʱ�䡣"
 Const E116 = "�����롰���к��롱����ӦС��16��"
 Const E117 = "�����롰������ϯID��ӦС��256��"
 Const E118 = "�����롰���ȼ���ӦС��256��"
 Const E119 = "�����롰�ɹ�ת������ID��С��32767��"
 Const E120 = "�����롰��ϯ��Ӧ����ʾ����ID��С��32767��"
 Const E121 = "�����롰��Ӧ��ת�ڵ�ID������255��С��32767��"
 Const E122 = "�����롰���ֻ�ǰ��ʱ(��)������0��С��256��"
 Const E123 = "�����롰ͨ�ų�ʱ(��)������ֵ��0��180֮�䡣"
 Const E124 = "�����롰��ʱ����(����)������ֵ��0��60֮�䡣"
 Const E125 = "�������岻���ظ������������á�"
 Const E126 = "�����롰�ڵ㳬ʱ(��)������ֵ��0��3600֮�䡣"
 Const E127 = "������������ID��"
 Const E128 = "������ID�Ѿ���ʹ�ã�������ѡ��"
 Const E129 = "����ID���ʧ�ܣ�������������ݿ⣡"
 'Michael Added @7-4-07
 Const E130 = "�����������û�����ID, ��ֵ��1��255֮��."
 Const E131 = "�����롰�ȴ�������ڽڵ�ID��������255, С��32767��"
 'Michael Added @ July,9,07
 Const E132 = "�����롰ת����ϯID��¼������������0, С��256��"
 Const E133 = "�����롰ת����ϯ���ż�¼������������0, С��256��"
 Const E134 = "��ѯ�ֶ�������Ϊ��"
 Const E135 = "��ѯ���������Ϊ��"
 Const E136 = "��ѯ���ʽ����Ϊ��"
  'Michael Added @ 2007-11-26
 Const E137 = "�����������ڻ�·������,���������·��!"
 Const E138 = "��ԴID����Ϊ0���"
 Const E139 = "��Դ����ظ�"
 Const E140 = "����Դ�ļ����Ͳ���֧�ֻ���Դ�ļ���������,��ȷ�Ϻ������ºϳ�..."
 Const E141 = "��ĿID����Ϊ��!"
 Const E142 = "��Ŀ������Ϊ��!"
 Const E143 = "��ĿID����Ϊ1��255֮�������!"
 Const E144 = "�����ļ�·������ȷ�����޷����������������������ã����˵���ѡ�->�����桱->��ϵͳ������Ŀ¼��"
 Const E145 = "�����롰���Խ�����ʾ���(��)������ֵ���ܴ��ڡ�¼��ʱ�䳤(��)��"
 Const E146 = ""
 Const E147 = ""
 Const E148 = "�ڵ�ID��ֵ������256��32767֮��"
 Const E149 = "������ԴID��С��32767��"
 Const E150 = "��ԴID��С��32767��"
 
 'Add End
 '********  Added End *******************
 Const E899 = "���������ݸ��³��������������Ӻ�������ã�"
 Const E999 = "������Ϣ���ţ�"

'**************************************************************
'Information message defination
'' < 100 for mesage box
 Const M001 = "�ó����������У����鴰���Ƿ���С����"
 Const M002 = "�����������ɹ���ȱʡ����Ϊ��888888���������Ѹò����߼�ʱ�޸Ŀ��"
 Const M003 = "û���˳�ϵͳ��Ȩ�ޡ�"
 Const M004 = "�˼�¼Ϊ��־��¼���벻Ҫɾ����"
 Const M005 = "�����۱��ѱ����������ܽ��б༭��"
 Const M006 = "��ӡ��ҵ�Ѿ���ɣ�"
 Const M007 = "�޴˼�¼��������ѡ������롣"
 Const M008 = "��ָ��ϣ����������¼֮ǰ�����¼�¼��"
 Const M009 = "�����Ѵ��ڣ����������룿"
 Const M010 = "������ɾ����"
 Const M011 = "����ͬ������ɣ�"
 Const M012 = "����Դ������Ŀ�겻����ͬ��"
 Const M013 = "��Դͬ������ɣ�"
'' >= 100 AND < 200 for status and help information
 Const M100 = "���ڽ���ϵͳ��ʼ��......"
 Const M101 = "���"
 Const M102 = "���ڷ������ݡ���"
 Const M103 = "���ڽ������ݡ���"
 Const M104 = "����!"
 Const M105 = "��û�и������ݵ�Ȩ�ޣ�"
 Const M106 = "���ڴ��͹����б�......"
 Const M107 = "���ڸ�λ������......"
 Const M108 = "����׼����ӡ����......"
 Const M109 = "���ڸ��·������˳ɼ�������......"
 Const M110 = "�������ʵ�����"
 Const M111 = "��ʾ������Ϊ3-8���ַ������֡�"
 Const M112 = "��ʾ������������������ı�ǩ���ɵ���������"
 Const M113 = "��ʾ�����б���ѡ����������A-G���������롣"
 Const M114 = "��ʾ�����б���ѡ�񼶱���������ּ��������롣"
 Const M115 = "��ʾ����ģ���������ô���ͨѶ�ڣ��Ա��շ����ݡ�"
 Const M116 = "��ʾ����ģ������ֵ��Ա��¼���Ա���и��������"
 Const M117 = "��ʾ����ģ�������޸�ֵ��Ա���"
 Const M118 = "��ʾ����ģ�����ڲ鿴�����ù�����־��"
 Const M119 = "��ʾ����ģ�����ڴ������޸ĺ�ɾ��ֵ��Ա��"
 Const M120 = "��ʾ����ģ�������޸Ĳ�ͬ���𱨾�������������ֵ��"
 Const M121 = "��ʾ����ģ�����ڸ��ĸ������¼��ı�����ʽ��"
 Const M122 = "��ʾ����ģ�����ڰ���ֵ��Ա��ֵ��ʱ���"
 Const M123 = "��ʾ����ģ�����ڶ�̬���Ӳ����仯�����"
 Const M124 = "��ʾ����ģ��������ָ����Ա����֪ͨ��"
 Const M125 = "���ڹر�ϵͳ......"
 Const M126 = "��ʾ�������Ѵ��ڣ������´�����"
 Const M127 = "��ʾ�����̺Ų���Ϊ�գ����������룡"
 Const M128 = "��ʾ�����̴�����ϣ�"
 Const M129 = "��ʾ�����ݲ����ڣ�"
 Const M130 = "��ʾ���������̳���"
 Const M131 = "��ʾ�����ȴ���������̣�"
 Const M132 = "�����ļ������ɹ���"
 Const M133 = "���ڵ�Ϊϵͳ�ڵ㣬����ɾ����"
 Const M134 = "�����ļ�����ɹ���"
 Const M135 = "����ɾ�����ڵ㣬����������ָ�����ڵ㣡"
 Const M136 = "���ڵ�Ϊϵͳ�ڵ㣬���ܽ��п�����ճ��������"
 Const M137 = "�Ѿ������£��·ݲ�����ǰ�ƣ�"
 Const M138 = "��ǰ�²���ѡ����Χ�ڣ�"
 Const M139 = "�Ѿ���ĩ�£��·ݲ����ٺ��ƣ�"
 Const M140 = "�ñ����Ѿ���Ŀ���б��д��ڣ�"
 Const M141 = "��ҳ�治Ϊ�գ�����ɾ��������ɾ��ҳ���е����нڵ㡣"
 Const M142 = "�����Ѿ����򿪣������ظ��򿪡�"
 Const M143 = "�����Ѿ��򿪣����ڱ༭״̬�����ȹرո����̡�"
 Const M144 = "��ʾ��Ŀǰ��δΪ����ָ����COM�ӿ���Դ�����ڡ��༭->�������ԡ��Ի��������á�"
 Const M145 = "�������ɹ���"
 Const M146 = "��Դ�ļ������ɹ���"
 Const M147 = "������Դ�ļ��ɹ���"
 Const M148 = "����ѡ����Ҫ���Ƶ���Դ���ٽ��и��Ʋ�����"
 Const M149 = "��Դ�б�Ϊ��,������������Դ�ٽ�����Ӧ����!"
 Const M150 = "��ʹ������Ŀ���Ƹ���ѡ��,��û��ѡ��Ŀ����Դ�滻����,��ȷ�ϡ�"
 Const M151 = "ϵͳδ��װ����TTS����,TTS����ת������������ʹ��,����ʹ�����Ȱ�װTTS����������"
 Const M152 = "�ڵ�IDС��255�Ľڵ�ͳ�Ʊ�ǩ���ɱ��༭!"
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
 Const Q001 = "��ȷ���Ƿ�Ҫ�˳���ǰӦ�ó���"
 Const Q002 = "�˲������޸�ԭ�����ݣ��Ƿ������"
 Const Q003 = "�������˳�����ʹ���ݶ�ʧ���Ƿ������"
 Const Q004 = "�ñ�Ŷ�Ӧ�����ݲ����ڣ��Ƿ��½���"
 Const Q005 = "��ǰ��������δ���棬�Ƿ��˳���"
 Const Q006 = "��ǰ��������δ���棬�Ƿ񱣴棿"
 Const Q007 = "��ǰ���ݿ�ת��ʧ�ܣ��Ƿ�������У�"
 Const Q008 = "ȷʵҪɾ����ǰ������"
 Const Q009 = "ȷʵҪע����"
 Const Q010 = "ȷʵҪɾ����"
 Const Q011 = "ȷʵҪ��������¼���"
 Const Q012 = "�ñ����Դ�Ѿ����ڣ��Ƿ񸲸ǣ�"
 Const Q013 = "���ȴ�������"
 Const Q014 = "�����Ѵ��ڣ����������룿"
 Const Q015 = "��ʾ�������ѱ��棡�Ƿ������"
 Const Q016 = "���ݲ����ڣ�"
 Const Q017 = "�����ж�����¼���ڣ�"
 Const Q018 = "�˲������ָ�ѡ��ʱ���������·ݵĽڼ���ȱʡ�趨���Ƿ������"
 Const Q019 = "�˲��������ѡ��ʱ���������·ݵĽڼ����趨���������������գ����Ƿ������"
 'Michael Added @ 2007-11-27
 Const Q020 = "ȷʵҪɾ����ǰ��Ŀ��"
 Const Q021 = "�ļ��в�����,��Ҫ������?"

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
 Const E069 = "��"
 Const E070 = "Failed to switching to node: please input content larger than 255, and smaller than 32767"
 Const E071 = "Please input: sending variable smaller than 256"
 Const E072 = "�ӽڵ㣺���������ݴ���255��С��32767��."
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
 Const E145 = "Please input��Recoring notify interval(s)����which can't larger than 'Max Record Duration (sec)'"
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
 Const M100 = " system initializing��."
 Const M101 = "finish"
 Const M102 = "sending data��"
 Const M103 = "receiving data��"
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
 Const M114 = "Notes��please select from list or key in directly"
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
 Const M140 = "The variable is already in target list��"
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
 Const M200 = "Database OK!"   'ȷ�����ݿ�
 Const M201 = " opening..."   '���ڴ����ݿ�
 Const M202 = " closing..."  ' ���ڹر����ݿ�
 Const M203 = " read..."    '���ڶ�ȡ���ݿ�
 Const M204 = " written..." '  ����д�����ݿ�
 Const M205 = " Connecting SQLServer..."  ' ��������SQL ������
 Const M206 = " Disconnecting SQLServer..." ' �Ͽ���ϵSQL ������
 Const M207 = " Reading SQLServer..." ' ��ȡSQL ������
 Const M208 = " Writing SQLServer..." ' д��SQL ������
 Const M209 = " searching..." ' ���ڲ�ѯ
 Const M210 = ""
 Const M299 = "Database Error!" '  ���ݿ����

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
 Const Q017 = "�����ж�����¼���ڣ�"
 Const Q018 = "�˲������ָ�ѡ��ʱ���������·ݵĽڼ���ȱʡ�趨���Ƿ������"
 Const Q019 = "�˲��������ѡ��ʱ���������·ݵĽڼ����趨���������������գ����Ƿ������"
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
 
'ȱʡ��Ϣ����->��Ϣ���Ŵ���
strMsgContens$ = Def_Message_Code_Error

'������Ϣ����
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
    '��ʾ��Ϣ
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
    '��ʾ��Ϣ
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
    
    '��ʾ��Ϣ
    If Len(Trim(strMsgContens$)) = 0 Then
        Message = MsgBox(Def_Message_Code_Error, vbCritical + vbOKOnly, App.Title)
    Else
        Message = MsgBox(strMsgContens$, vbQuestion + vbYesNo + nDefaultButton, App.Title)
    End If
    
  Case Else
    '��ʾ��Ϣ
    Message = MsgBox(Def_Message_Code_Error, vbCritical + vbOKOnly, App.Title)
End Select

End Function

