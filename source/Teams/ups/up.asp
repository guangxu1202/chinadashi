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
	Conn.execute("update [Clubconfig] set Allclass='"& p &"',ClearMail='[b]亲爱的{$username}, 您好 [/b]"& vbcrlf&""& vbcrlf&" 恭喜您成功地注册了您的资料, 非常感谢您使用 {$clubname} 的服务,以下是您注册的资料，请妥善保管： "& vbcrlf&""& vbcrlf&"   * 您的帐号是:{$username} "& vbcrlf&"   * 密码是：{$userpass} "& vbcrlf&"   *{$isregkey}"& vbcrlf & vbcrlf & vbcrlf &" 最后, 有几点注意事项请您牢记 "& vbcrlf & vbcrlf &"    1、请遵守《计算机信息网络国际联网安全保护管理办法》里的一切规定"& vbcrlf&"    2、使用轻松而健康的话题，所以请不要涉及政治、宗教等敏感话题。"& vbcrlf&"    3、承担一切因您的行为而直接或间接导致的民事或刑事法律责任"& vbcrlf & vbcrlf & vbcrlf&"本论坛服务由 {$clubname} 提供。 "& vbcrlf&"{$emailkey}'")
	Application.Contents.RemoveAll() 
	Response.Write " 更新OK，现在将转入论坛首页。请稍后。。。。<meta http-equiv=refresh content=0;url=../></div></body>"
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
		Response.Write "数据库连接出错，请检查连接字串。"
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
	Response.Write"<body text=""#000000"" leftmargin=""10"" topmargin=""10""><Br><BR><div class=a3 style=""padding: 15px;width:600""><div class=a4><li>模版导入成功，现在将转入论坛首页<BR><meta http-equiv=refresh content=3;url=../></div></body>" 

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
	Team Board 升级程序

	<li> 更新注册模块 
	<li> 修正一些2.0.4存在的BUG
	</div><BR>
	<div style="width:600px;background-color:#fff;padding:10px;text-align :left;">
	<ul>
		<li>点击下面的链接，可以将您的team 2.0.3/2.0.4 版本升级到2.0.5版本 </li>

		<li> 注明： 由于此次数据库升级修正模板存在的问题，需要您同时进行模板的升级，请点击下面的链接进行数据库的升级。
	</ul>
	<ul>
		<li> <a href="?menu=update" Style='cursor:hand;color:red;'> TEAM 论坛升级文件2.0.5升级-->> </a></li>
	</ul>
	<hr style="width:550px;text-align :center;">
	<ul>
		<li> 模板升级前需要将官方压缩包Skins目录下面的TM_Style.mdb上传到论坛的skins目录下面，覆盖原文件。</li>
	</ul>
	<ul>
		<li> <a href="?menu=upskins" Style='cursor:hand;color:red;'> 官方模版升级 -->> </a></li>
	</ul>
	</div>
<%
End Sub
%>
