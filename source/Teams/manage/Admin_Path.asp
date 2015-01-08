<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Public boards
Dim Admin_Class,Uid
Call Master_Us()
Uid = Cid(Request("uid"))
Header()
Admin_Class=",11,"
Call Master_Se()
team.SaveLog ("  统计信息")
Select Case Request("action")
	Case "discreteness"
		Call discreteness
	Case "statroom"
		Call statroom
	Case Else
		Call Main
End Select
Sub statroom	
	Dim fso,upfacedir,d,upphotosize,tolsize,totalBytes,upfacesize,upphotodir
	dim upfiledir,upfilesize,toldir
	set fso=server.createobject("scripting.filesystemobject")
	upfacedir=server.mappath("../images/upface")
	set d=fso.getfolder(upfacedir)
	upfacesize=d.size
	upfiledir=server.mappath("../images/upfile")
	set d=fso.getfolder(upfiledir)
	upfilesize=d.size
	toldir=server.mappath("../")
	set d=fso.getfolder(toldir)
	tolsize=d.size
	totalBytes=upfacesize+upphotosize+upfilesize+tolsize
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="10" width="95%" align="center" class="a2">
  <tr class="a1">
    <td colspan="2">统计占用空间 [总 <%totalsize(totalBytes)%>]</td>
  </tr>
  <tr class="a3">
	<td width="20%"> 上传头像占用空间 </td>
	<td>
		<div class="a2" Style="width:<%=Int(upfacesize/totalBytes*100)%>%;padding:3px" title="<%=Int(upfacesize/totalBytes*100)%>%"> <%totalsize(upfacesize)%> </div>
	 </td>
  </tr>
  <tr class="a4">
	<td> 上传附件占用空间 </td>
	<td>
		<div class="a2" Style="width:<%=Int(upfilesize/totalBytes*100)%>%;padding:3px" title="<%=Int(upfilesize/totalBytes*100)%>%"> <%totalsize(upfilesize)%> </div>
	 </td>
  </tr>
  <tr class="a3">
	<td> 目录总共占用空间 </td>
	<td>
		<div class="a2" Style="width:<%=Int(tolsize/totalBytes*100)%>%;padding:3px" title="<%=Int(tolsize/totalBytes*100)%>%"> <%totalsize(tolsize)%> </div>
	 </td>
  </tr>
</table>
<%
End Sub

Sub totalsize(size)
	if size<1024 then
		response.write size&" Bytes"
	elseif size<1048576 then
		response.write Int(size/1024)&" KB"
	else
		response.write Int(size/1024/1024)&" MB"
	end if
End sub
Sub discreteness	

	Echo " <body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10""> "
	Echo " <br>"
	Echo " <table cellspacing=""1"" cellpadding=""4"" width=""95%"" align=""center"" class=""a2""> "
	Echo "	<tr class=""a1"">"
    Echo "		<td width=""50%"">组件名称</td><td>详情</td> "
	Echo "	</tr>"
	Dim theInstalledObjects(17)
	theInstalledObjects(0) = "MSWC.AdRotator"
	theInstalledObjects(1) = "MSWC.BrowserType"
	theInstalledObjects(2) = "MSWC.NextLink"
	theInstalledObjects(3) = "MSWC.Tools"
	theInstalledObjects(4) = "MSWC.Status"
	theInstalledObjects(5) = "MSWC.Counters"
	theInstalledObjects(6) = "MSWC.PermissionChecker"
	theInstalledObjects(7) = "Scripting.FileSystemObject"
	theInstalledObjects(8) = "adodb.connection"
	theInstalledObjects(9) = "SoftArtisans.FileUp"
	theInstalledObjects(10) = "SoftArtisans.FileManager"
	theInstalledObjects(11) = "JMail.Message"
	theInstalledObjects(12) = "CDONTS.NewMail"
	theInstalledObjects(13) = "Persits.MailSender"
	theInstalledObjects(14) = "LyfUpload.UploadFile"
	theInstalledObjects(15) = "Persits.Upload.1"
	theInstalledObjects(16) = "w3.upload"
	theInstalledObjects(17) = "Persits.Jpeg"
	For i=0 to 17
		Response.Write "<tr class=""a4""><td class=""a2"">&nbsp;" & theInstalledObjects(i) & " &nbsp;"
		Select case i
			case 7
				Response.Write "(FSO 文本文件读写)"
			case 8
				Response.Write "(ACCESS 数据库)"
			case 9
				Response.Write "(SA-FileUp 文件上传)"
			case 10
				Response.Write "(SA-FM 文件管理)"
			case 11
				Response.Write "(JMail 邮件发送)"
			case 12
				Response.Write "(WIN虚拟SMTP 发信)"
			case 13
				Response.Write "(ASPEmail 邮件发送)"
			case 14
				Response.Write "(LyfUpload 文件上传)"
			case 15
				Response.Write "(ASPUpload 文件上传)"
			case 16
				Response.Write "(w3 upload 文件上传)"
			case 17
				Response.Write "(ASP水印)"
		end select
		Response.Write "  " & getver(theInstalledObjects(i)) & "  </td><td><div class=""a2"" style=""width:30;padding:3px;"">"
		If Not IsObjInstalled(theInstalledObjects(i)) Then
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<b>√</b> "
		End If
		Response.Write "</div></td></TR>" & vbCrLf
	Next
	Echo "</table>"
End Sub

Sub Main%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>本页面只做完参考使用，计算参数因为各服务器性能的不同而区别。
      </ul>
	  </td>
  </tr>
</table>
<br>
<form method="post" action="?action=onlinelistok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
      <td>测试项目</td>
      <td>测试值</td>
    </tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器的域名 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("server_name")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器的IP地址 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("LOCAL_ADDR")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器操作系统 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("OS")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器解译引擎 </td>
		<td bgcolor="#F8F8F8"> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器软件的名称及版本 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("SERVER_SOFTWARE")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器正在运行的端口 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("server_port")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器CPU数量 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器Application数量 </td>
		<td bgcolor="#F8F8F8"> <%=Application.Contents.Count%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器Session数量 </td>
		<td bgcolor="#F8F8F8"> <%=Session.Contents.Count%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 请求的物理路径 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("path_translated")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 请求的URL </td>
		<td bgcolor="#F8F8F8"> http://<%=Request.ServerVariables("server_name")%><%=Request.ServerVariables("script_name")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> 服务器当前时间 </td>
		<td bgcolor="#F8F8F8"> <%=Now()%> </td>
	</tr>

	<tr> 
		<td bgcolor="#FFFFFF"> 脚本连接超时时间 </td>
		<td bgcolor="#F8F8F8"> <%=Server.ScriptTimeout%> 秒 </td>
	</tr>
</table><BR>
 <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr align="left" class="a1"> 
		<td> 磁盘文件操作速度测试 </td>
	</tr>
	<tr class="a2"><td> <%
	Response.Flush

	If Not IsObjInstalled("Scripting.FileSystemObject") Then
		Response.Write "不支持FSO,无法测试文件读取"
	Else
	Response.Write "正在重复创建、写入和删除文本文件50次..."
	Dim thetime3,tempfile,iserr,t1,FsoObj,tempfileOBJ,t2
	Set FsoObj=Server.CreateObject("Scripting.FileSystemObject")
	iserr=False
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	For i=1 To 50
		Err.Clear

		Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		If Err <> 0 Then
			Response.Write "创建文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		If Err <> 0 Then
			Response.Write "写入文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		If Err <> 0 Then
			Response.Write "删除文件错误！"
			iserr=True
			Err.Clear
			Exit For
		end if
		Set tempfileOBJ=Nothing
	Next
	t2=timer
	If Not iserr Then
		thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
		Response.Write "...已完成！本服务器执行此操作共耗时 <font color=red>" & thetime3 & " 毫秒</font>"
	End If
	End If
%>
<BR>P4 2.4,2GddrEcc,SCSI36.4G*2 执行此操作需要 <font color=red>32～65</font> 毫秒</a>
<%Response.Flush%>
</td>
	</tr>
	</table>
	<BR> 
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr align="left" class="a1"> 
		<td>  ASP脚本解释和运算速度测试  </td>
	</tr>
	<tr class="a2">
		<td>
<%
	Response.Write "整数运算测试，正在进行50万次加法运算..."
	dim lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime & " 毫秒</font><br>"
	Response.Write "浮点运算测试，正在进行20万次开方运算..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime2 & " 毫秒</font><br>"
%>
	<BR>P4 2.4,2GddrEcc,SCSI36.4G*2 整数运算需要 <font color=red>171～203</font> 毫秒, 浮点运算需要 <font color=red>156～171</font> 毫秒
	</td></tr>
  </table>
<br>
<br>
<%
End Sub


Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
''''''''''''''''''''''''''''''
Function getver(Classstr)
	On Error Resume Next
	getver=""
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(Classstr)
	If 0 = Err Then getver=xtestobj.version
	Set xTestObj = Nothing
	Err = 0
End Function
footer()
%>
