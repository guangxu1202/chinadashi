<!-- #include file="Conn.asp" -->
<!-- #include file="Inc/Md5.asp" -->
<link href="skins/teams/bbs.css" rel=stylesheet type=text/css>
<%
Dim mconn
Dim MyCode,IsWrite,CodeStr,NextAction
'====================使用前请修改此代码==================
'上传前修改验证密码，防止被他人利用。
MyCode = "team_admin_key"
'文件初始值为1，需要修改为0才可以运行此文件。
IsWrite = 1
'========================================================
NextAction = 0
CodeStr = ""
If IsWrite = 1 Then
	Response.Write "<BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAM后台管理安装程序 >>> </td></tr><tr class=a4><td height=100 valign=top  colspan=3> <li><b> 说明: </b></li> <UL>安装文件未启用，请打开Install.asp文件，设置IsWrite = 0 ，然后重新上传原文件。</UL></td></tr></table><br><center><input type=Submit value='下一步' name=Submit></center>"
	Response.Write" <BR/><table border=""0"" cellspacing=""3"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4><td  colspan=3>论坛名称: TEAM 论坛<BR/>代码编写: DayMoon<BR/>论坛地址: <a href=http://WWW.TEAM5.CN>WWW.TEAM5.CN</a> </td></tr></table>"
	Response.End 
End if
Select Case Request("Menu")
	Case "Update"
		Update
	Case "Update1"
		Call Update1
	Case "upskins"
		Call upskins
	Case "upclubsys"
		Call upclubsys
	Case Else
		Main
End Select

Sub upclubsys
	ismyKey()
	Conn.execute("update [Clubconfig] set Allclass='0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0$$$0$$$0$$$论坛维护中,请稍后访问!$$$1$$$$$$1$$$0$$$TEAM 2.0.4 Release$$$$$$0$$$0.2$$$100$$$0$$$10$$$0$$$欢迎加入TEAM论坛!$$$1$$$1$$$20$$$10$$$20$$$15$$$3$$$0$$$1000$$$1$$$1$$$1$$$1$$$team论坛,asp,bbs,免费,急速$$$team论坛,asp,bbs,免费,急速$$$3$$$1$$$30$$$31$$$1$$$1$$$1$$$0$$$1$$$$$$1$$$0$$$60$$$60$$$0$$$3$$$0$$$1$$$10$$$30$$$10$$$500$$$5$$$$$$0$$$$$$$$$粤ICP备05004532号$$$5$$$1$$$1$$$1$$$0$$$0$$$TEAM BOARD$$$10000$$$10$$$1$$$0$$$200$$$100$$$rar|jpg|txt|gif|zip$$$999$$$0$$$images/Mypic/logo.gif$$$88$$$1$$$$$$$$$$$$宋体$$$0$$$$$$1$$$$$$$$$40$$$80$$$0$$$20$$$0$$$0$$$0$$$1$$$1$$$$$$1$$$1$$$51$$$$$$$$$$$$100$$$100$$$0$$$$$$120$$$120$$$0$$$0$$$$$$1$$$$$$1$$$0$$$$$$$$$0$$$$$$'")
	FSOlinewrite "Install.asp",10,"IsWrite = 1"
	Response.Write"<body text=""#000000"" leftmargin=""10"" topmargin=""10""><Br><BR><div class=a3 style=""padding: 15px;width:600""><div class=a4><li>模版导入成功，现在将转入论坛首页<BR><meta http-equiv=refresh content=3;url=../></div></body>" 
End sub

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
	ismyKey()
	On Error Resume Next
	If Not IsObject(Conn) Then ConnectionDatabase
	conn.execute("delete from [style]")
	If Not IsObject(mConn) Then myConnectionDatabase
	Dim i
	Dim InsertName,InsertValue
	Set TRs= mconn.Execute("select * from [Style] where id = 1 ")
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
	FSOlinewrite "Install.asp",10,"IsWrite = 1"
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

Sub Update()
	Dim Myname,MyPass,Forumname,ForumPass
	If Request.Form("NextAction") = 1 Then
		Myname=HTMLEncode(Request.Form("Myname"))
		MyPass=HTMLEncode(Request.Form("MyPass"))
		Forumname=HTMLEncode(Request.Form("Forumname"))
		ForumPass=HTMLEncode(Request.Form("ForumPass"))
		Response.Write"<div align=left id=UpFile class=a1><div class=a4>"
		ismyKey()
		ConnectionDatabase
		Conn.Execute("Update [User] Set Userpass ='"&MD5(ForumPass,16)&"',UserGroupID=1 Where  UserName='"&Forumname&"'")
		Conn.Execute("insert into admin (adminname,adminpass,adminclass,forumname) values ('"&Myname&"','"&MD5(MyPass,16)&"','1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36','"&Forumname&"')")
		Response.Write"<li> TEAM更新成功 </li>"
		Response.Write"<li> TEAM前台用户名称："&Forumname&"，密码："&ForumPass&"</li>"
		Response.Write"<li> TEAM后台用户名称："&Myname&"，密码："&MyPass&"</li>"
		Response.Write"<li> 更新成功,请在论坛根目录删除此文件! 文件名 [Install.asp] </li></div></div>"
		FSOlinewrite "Install.asp",10,"IsWrite = 1"
	Else
		Response.Write "请输入正确的验证密码。"
		Response.End 
	End if
End Sub

Sub Main
	Response.Write" <BR/><BR/><form name=myform method=post action='?Menu=Update1'><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAM后台管理安装程序 >>> </td></tr><tr class=a4><td height=100 valign=top  colspan=3> <li><b> 说明: </b></li> <UL>1. 此文件用于管理忘记后台登陆密码或错误操作后台用户权限时,从新更新后台管理使用. </UL> <UL>2. 请修改此文件的操作密码后再上传到论坛根目录!</UL><UL>3. 使用完文件以后请及时删除此文件! </UL><UL>4. 更新确认码为人工设置的代码,请打开Install.ASP参照说明设置! </UL></td></tr> <tr class=a3><td colspan=3>请输入更新确认码: <input size='15' name='CodeStr'></td></tr></table><br><center><input type=Submit value='下一步' name=Submit></center>"
	Response.Write" <BR/><table border=""0"" cellspacing=""3"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4><td  colspan=3>论坛名称: TEAM 论坛<BR/>代码编写: DayMoon<BR/>论坛地址: <a href=http://WWW.TEAM5.CN>WWW.TEAM5.CN</a> </td></tr></table>"
End Sub

Sub Update1
	ismyKey()
	ConnectionDatabase
	Response.Write" <BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAM系统恢复操作指引 >>> </td></tr></table>"
	Response.Write" <BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>模板恢复</td></tr>"
	Response.Write" <tr class=a4 align=center><td  colspan=3><li> 1. 请先确认在论坛skins目录下面存在TM_Style.MDB文件。<li> 2.  查看TM_Style.MDB文件的更新日期。参照官方提供的压缩包进行对比。<li>点击恢复功能将会删除论坛其他模板，并将模板恢复为论坛官方默认的模板。<li>确认了以上选项，请点击 <B><a href=""?Menu=upskins"">模板恢复</a></B></td></tr>"
	Response.Write" </table>"
		
	Response.Write" <BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>基本参数恢复</td></tr>"
	Response.Write" <tr class=a4 align=center><td  colspan=3>点击此选项，将论坛恢复到原始参数设置。 <B><a href=""?Menu=upclubsys"">基本参数恢复</a></B></td></tr>"
	Response.Write" </table>"

	Response.Write" <BR/><BR/><form name=myform method=post action='?Menu=Update'><input type=""hidden"" value="""&CodeStr&""" name=""CodeStr""><input type=""hidden"" value=""1"" name=""NextAction""><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAM后台管理安装程序 >>> </td></tr><tr class=a4><td height=100 valign=top  colspan=3> <li><b> 说明: </b></li> <UL>1. 此文件用于管理忘记后台登陆密码或错误操作后台用户权限时,从新更新后台管理使用. </UL> <UL>2. 请修改此文件的操作密码后再上传到论坛根目录!</UL><UL>3. 使用完文件以后请及时删除此文件! </UL><UL>4. 更新确认码为人工设置的代码,请打开Install.ASP参照说明设置! </UL><UL>5. 此操作将重新设置前台用户密码，并添加一个新的后台登陆帐号，请注意不要使用重复的后台密码，不然将导致用户无法登陆后台 . 请设置新的后台登陆帐号时，保证登陆名称不重复。 </UL></td></tr>"
	ConnectionDatabase
	Dim Rs
	Response.Write" <tr class=a1><td colspan=2>  目前存在的后台用户  </td></tr><tr class=a4><td colspan=2>"
	Set Rs=Conn.Execute("Select adminname,forumname from [admin]" )
	Do While Not Rs.Eof
		Response.Write " <li>后台用户名称： "& RS(0)&" -  [ 绑定的前台用户名： "& RS(1)&" ]</li> "
		Rs.MoveNext
	Loop
	Rs.close:Set Rs=nothing
	Response.Write" </td> </tr><tr class=a1><td colspan=2>  添加的后台用户  </td></tr>"
	Response.Write" <tr class=a4><td>前台用户名称: <input size='15' name='Forumname'> </td><td>前台用户密码: <input size='15' name='ForumPass'>  </td> </tr>"
	Response.Write"  <tr class=a3><td>后台用户名称: <input size='15' name='Myname'> </td><td>后台用户密码: <input size='15' name='MyPass'>  </td></tr>"
	Response.Write" </table><br><center><input type=Submit value='下一步' name=Submit></center> "
	Response.Write" <BR/><table border=""0"" cellspacing=""3"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4><td  colspan=3>论坛名称: TEAM 论坛<BR/>代码编写: DayMoon<BR/>论坛地址: <a href=http://WWW.TEAM5.CN>WWW.TEAM5.CN</a> </td></tr></table>"
End Sub

Sub ismyKey()
	Dim CodeStr
	CodeStr=Request.Form("CodeStr")
	If Trim(MyCode) <>  Trim(CodeStr) Then
		Response.Write "<BR/><BR/><BR><BR><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4 align=center><td  colspan=3>请输入正确的验证密码。 >>> </td></tr></table>"
		Response.End 
	End if
End Sub

Function FSOlinewrite(filename,lineNum,Linecontent)
	if linenum < 1 then exit function
	dim fso,f,temparray,tempCnt
	set fso = server.CreateObject("scripting.filesystemobject")
	if not fso.fileExists(server.mappath(filename)) then exit function
	set f = fso.opentextfile(server.mappath(filename),1)
	if not f.AtEndofStream then
		tempcnt = f.readall
		f.close
		temparray = split(tempcnt,chr(13)&chr(10))
		if lineNum>ubound(temparray)+1 then
			exit function
		else
		temparray(lineNum-1) = lineContent
		end if
		tempcnt = join(temparray,chr(13)&chr(10))
		set f = fso.createtextfile(server.mappath(filename),true)
		f.write tempcnt
	end if
	f.close
	set f = nothing
End Function

Function HTMLEncode(fString)
	If fString="" or IsNull(fString) Then 
		Exit Function
	Else
		If Instr(fString,"'")>0 Then 
			fString = replace(fString, "'","&#39;")
		End If
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32)," ")
		fString = Replace(fString, CHR(9)," ")
		fString = Replace(fString, CHR(34),"&quot;")
		fString = Replace(fString, CHR(13),"")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10),"<BR>")
		fString = Replace(fString, CHR(39),"&#39;")
		HTMLEncode = fString
	End If
End Function
%>
