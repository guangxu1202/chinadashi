<!-- #include file="../conn.asp" -->
<link rel="stylesheet" rev="stylesheet" href="../skins/teams/bbs.css" type="text/css" media="all" />
<%
Server.ScriptTimeout= 9999
MyDbPath = "../"
Select Case Request("menu")
	Case "update"
		Update()
	Case "upskins"
		Call Upskins
	Case Else
		Call main
End Select

Sub Update()
	Dim Rs,newx,i,p
	Response.Write"<body text=""#000000"" leftmargin=""10"" topmargin=""10""><Br><BR><div class=a3 style=""padding: 15px;width:600""><div class=a4>"
	If Not IsObject(Conn) Then ConnectionDatabase
	Set rs = conn.execute("select Allclass from Clubconfig")
	newx = Split(RS(0),"$$$")
	For i = 0 To UBound(newx)
		If p ="" then
			p = newx(0)
		ElseIf i = 8 Then
			p = P & "$$$TEAM 2.0.5 Release"
		Else
			p =  P & "$$$" & newx(i)
		End if
	Next
	rs.close
	conn.Execute("ALTER TABLE [Clubconfig] alter column  ClearMail text")
	conn.Execute("ALTER TABLE [bbsconfig] alter column  Board_Key text")
	Conn.execute("update [Clubconfig] set Allclass='"& p &"',ClearMail='[b]�װ���{$username}, ���� [/b]"& vbcrlf&""& vbcrlf&" ��ϲ���ɹ���ע������������, �ǳ���л��ʹ�� {$clubname} �ķ���,��������ע������ϣ������Ʊ��ܣ� "& vbcrlf&""& vbcrlf&"   * �����ʺ���:{$username} "& vbcrlf&"   * �����ǣ�{$userpass} "& vbcrlf&"   *{$isregkey}"& vbcrlf & vbcrlf & vbcrlf &" ���, �м���ע�����������μ� "& vbcrlf & vbcrlf &"    1�������ء��������Ϣ�������������ȫ��������취�����һ�й涨"& vbcrlf&"    2��ʹ�����ɶ������Ļ��⣬�����벻Ҫ�漰���Ρ��ڽ̵����л��⡣"& vbcrlf&"    3���е�һ����������Ϊ��ֱ�ӻ��ӵ��µ����»����·�������"& vbcrlf & vbcrlf & vbcrlf&"����̳������ {$clubname} �ṩ�� "& vbcrlf&"{$emailkey}'")
	Application.Contents.RemoveAll() 
	Response.Write " ����OK�����ڽ�ת����̳��ҳ�����Ժ󡣡�����<meta http-equiv=refresh content=0;url=../></div></body>"
End Sub



Dim mconn
Sub myConnectionDatabase
	Dim ConnStr
	ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & "skins/TM_Style.mdb")
	On Error Resume Next
	Set mconn = Server.CreateObject("ADODB.Connection")
	mconn.open ConnStr
	If Err Then
		err.Clear
		Set mconn = Nothing
		Response.Write "���ݿ����ӳ������������ִ���"
		Response.End
	End If
End Sub

Sub upskins
	Dim tRs
	On Error Resume Next
	If Not IsObject(Conn) Then ConnectionDatabase
	conn.execute("update [style] set StyleHid=0")
	If Not IsObject(mConn) Then myConnectionDatabase
	Dim i
	Dim InsertName,InsertValue
	Set TRs= mconn.Execute("select * from [Style] where id>0 ")
	Do while not TRs.eof
	InsertName=""
	InsertValue=""
		For i = 1 to TRs.Fields.Count-1
			InsertName=InsertName & TRs(i).Name
			InsertValue=InsertValue & "'" &checkStr(TRs(i)) & "'"
			If i<> TRs.Fields.Count-1 Then 
				InsertName	= InsertName & ","
				InsertValue	= InsertValue & ","
			End If
		Next
	conn.Execute("insert into [Style] ("&InsertName&") values ("&InsertValue&") ")
	TRs.movenext
	loop
	TRs.close
	set TRs=Nothing
	Application.Contents.RemoveAll()
	Response.Write"<body text=""#000000"" leftmargin=""10"" topmargin=""10""><Br><BR><div class=a3 style=""padding: 15px;width:600""><div class=a4><li>ģ�浼��ɹ������ڽ�ת����̳��ҳ<BR><meta http-equiv=refresh content=3;url=../></div></body>" 

End Sub

Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
End Function

Sub main

	%>
	<BR/><BR/>
	<body Style="background-color:#9EB6D8" text="#000000" leftmargin="10" topmargin="10">
	<div style="width:600px;background-color:#fff;padding:10px;text-align :left;">
	Team Board ��������

	<li> ����ע��ģ�� 
	<li> ����һЩ2.0.4���ڵ�BUG
	</div><BR>
	<div style="width:600px;background-color:#fff;padding:10px;text-align :left;">
	<ul>
		<li>�����������ӣ����Խ�����team 2.0.3/2.0.4 �汾������2.0.5�汾 </li>

		<li> ע���� ���ڴ˴����ݿ���������ģ����ڵ����⣬��Ҫ��ͬʱ����ģ���������������������ӽ������ݿ��������
	</ul>
	<ul>
		<li> <a href="?menu=update" Style='cursor:hand;color:red;'> TEAM ��̳�����ļ�2.0.5����-->> </a></li>
	</ul>
	<hr style="width:550px;text-align :center;">
	<ul>
		<li> ģ������ǰ��Ҫ���ٷ�ѹ����SkinsĿ¼�����TM_Style.mdb�ϴ�����̳��skinsĿ¼���棬����ԭ�ļ���</li>
	</ul>
	<ul>
		<li> <a href="?menu=upskins" Style='cursor:hand;color:red;'> �ٷ�ģ������ -->> </a></li>
	</ul>
	</div>
<%
End Sub
%>
