<%@ LANGUAGE=VBScript CodePage=936%>
<%
Option Explicit
Response.Buffer = True
Response.Charset = "gb2312"
Session.CodePage = 936		'utf=65001,gbk=936	

Dim Db,conn
Dim SqlNowString,team,Cache
dim Startime,MyDbPath,SqlQueryNum
Startime=timer()
SqlQueryNum = 0
Const WebSuperAdmin = "admin"		'Ĭ�����ù���Ա,
Const Set_cookies = 1				'�Ƿ����cookies,1Ϊ����,0Ϊ�ر�,����ռ俪��cookies����
Const IsForum=""					'�������ݱ�ǰ׺,�� " TM_ "
Const ManagePath="Manage/"			'�����ù����̨���ļ���·��
Const IsSqlDataBase = 0				'�������ݿ����0ΪAccess���ݿ⣬1ΪSQL���ݿ�
Const IsDeBug = 0					'��������ģʽ������Ϊ1����������Ϊ0,�����������Ϣ�����ڰ�ȫ
'============================================================================================
Const IPDate = "Data/ipdata.mdb"	'IP���ݿ�ĵ�ַ
Const LogDate = "Data/LOGs.mdb"		'�����¼���ݿ�
'=============================================================================================
If IsSqlDataBase = 1 Then			'sql���ݿ����Ӳ���
	Const SqlDatabaseName = "teams"		'���ݿ���(SqlDatabaseName)
	Const SqlPassword = ""		'�û�����(SqlPassword)
	Const SqlUsername = "sa"			'�û���(SqlUsername)
	Const SqlLocalName = "(local)"		'������(SqlLocalName)��������local�������IP��
	SqlNowString = "GetDate()"
Else
	Db = "Data/TEAM.mdb" '���ݿ��ַ,�����ʹ��team���̳ϵͳ,���뽫�˴������ݿ��ַ�����ݿ����ƽ����޸�.��ͨ��FTP����Կռ���������ݿ�ĵ�ַ�ʹ��Ŀ¼������ͬ���޸�. ��������ĵ�ַ������,�������"���ݿ����ӳ������������ִ���"����ʾ.
	SqlNowString = "Now()"
End If

'=============��̳���ݿ�=================================
Sub ConnectionDatabase
	Dim ConnStr
	If IsSqlDataBase = 1 Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & db)
	End If
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "���ݿ����ӳ������������ִ���"
		Response.End
	End If
End Sub
%>