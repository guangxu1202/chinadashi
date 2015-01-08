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
Const WebSuperAdmin = "admin"		'默认内置管理员,
Const Set_cookies = 1				'是否编码cookies,1为开启,0为关闭,国外空间开启cookies编码
Const IsForum=""					'定义数据表前缀,如 " TM_ "
Const ManagePath="Manage/"			'自设置管理后台的文件夹路径
Const IsSqlDataBase = 0				'定义数据库类别，0为Access数据库，1为SQL数据库
Const IsDeBug = 0					'定义运行模式，测试为1，正常运行为0,不输出错误信息有利于安全
'============================================================================================
Const IPDate = "Data/ipdata.mdb"	'IP数据库的地址
Const LogDate = "Data/LOGs.mdb"		'管理记录数据库
'=============================================================================================
If IsSqlDataBase = 1 Then			'sql数据库连接参数
	Const SqlDatabaseName = "teams"		'数据库名(SqlDatabaseName)
	Const SqlPassword = ""		'用户密码(SqlPassword)
	Const SqlUsername = "sa"			'用户名(SqlUsername)
	Const SqlLocalName = "(local)"		'连接名(SqlLocalName)（本地用local，外地用IP）
	SqlNowString = "GetDate()"
Else
	Db = "Data/TEAM.mdb" '数据库地址,如果是使用team搭建论坛系统,必须将此处的数据库地址和数据库名称进行修改.并通过FTP软件对空间里面的数据库的地址和存放目录进行相同的修改. 如果两处的地址不符合,将会出现"数据库连接出错，请检查连接字串。"的提示.
	SqlNowString = "Now()"
End If

'=============论坛数据库=================================
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
		Response.Write "数据库连接出错，请检查连接字串。"
		Response.End
	End If
End Sub
%>