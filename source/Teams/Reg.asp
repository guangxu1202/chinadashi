<!-- #include file="Conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
Dim username,errorchar,i
team.Headers Team.Club_Class(1)& "- 用户注册"
If team.Forum_setting(4) = 0 then 
	team.error team.Forum_setting(5)
End If
UserName = HRF(2,1,"username")
Dim X1,x2,Fid
If team.UserLoginED Then
	team.error " 欢迎您回来，"&TK_UserName&"。现在将转入首页。<meta http-equiv=refresh content=3;url=""Default.asp""> "
End if
Select Case Request("action")
	Case "callreg"
		CallReg
	Case "myCheck"
		Call myCheck
	Case Else
		Call Main()
End Select

Sub Main()
	If team.Forum_setting(17)=0 Then
		Call myCheck
	Else
		Call RegMain()
	End If
End Sub

Sub RegMain()
	Echo "<form method=""post"" action=""?action=myCheck"">"
	Echo " <div id=""regtop""> 继续注册前请先阅读【"& team.Club_Class(1) &" － TEAM's】注册服务条款和声明 </div>"
	Echo " <div id=""reginfo""> "& Replace(Replace(Ubb_Code(team.Club_Class(13)),"{$clubname}",team.Club_Class(1)),vbcrlf,"<BR>") &" </div>"
    Echo " <div id=""regfoot"">"
	Echo " <li><a href=help.asp?action=custom>互联网电子公告服务管理规定</a></li>"
	Echo " <li><a href=help.asp?action=custom5>联网信息服务管理办法 </a></li>"
	Echo " </div><BR><input type=submit value=""同 意""></form><BR>"
End Sub

Sub myCheck
	Dim Tmp,Checkname,i,dmail
	dmail = 1
	If CID(team.Forum_setting(6)) = 0 Then dmail = 0
	%>
		<script src="Js/calendar.js" type="text/javascript"></script>
		<!--导航模块开始ElseHtml (4)-->
		<table border="0" width="98%" align="center" cellspacing="0" cellpadding="0">
		   <tr>
				<td class="a4"> <span class="bold">
				   <a href="default.asp"><%=team.Club_Class(1)%></a>  &raquo;  
				   </a> &raquo;  <A href="Reg.asp">注册协议</a></td>
				</td>
		   </tr>
		</table><br />
		<script language="JavaScript"> var dmail='<%=dmail%>';</script>
		<script language="JavaScript" src="Js/Reg.js"></script>
		<form method="Post" name="tform" id="tform" action="?action=callreg" onSubmit="return validate(this)">
		<input type="hidden" value="upreg" name="action">
		<input type="hidden" value="<%=Session.SessionID%>" name="formhash">
		<table cellspacing="1" cellpadding="4" width="98%" align="center" class="a2">
			<tr>
				<td colspan="3" class="a1">注册 - 必填内容</td>
			</tr>
			<tr title="用户名可以由中文，英文字母（不区分大小写）、数字（0-9）、下划线、连字符号组成，必须大于等于3个字符。输入用户名后系统会自动检测用户名是否合法。" class="a4">
				<td width="20%"> 用户名：(<a class="red">*</a>) </td>
				<td width="38%"> <input size="25" type="text" name="username" value="" onfocus="setfocus('ukey')" onBlur="setblur('ukey');checkfocus('ukey');"></td>
				<td width="42%"><span id="ukey" class="blur">用户名长度为3-20个字符。</span></td>
			</tr>
			<tr>
				<td> 密　码：(<a class="red">*</a>) </td>
				<td> <input size="25" type="password" name="password" value="" onfocus="setfocus('pkey')" onBlur="setblur('pkey');checkpass('pkey');"></td>
				<td><span id="pkey" class="blur">密码长度3-20位，字母请区分大小写，请使用字母加数字的组合。</span></td>
			</tr>
			<tr class="a4">
				<td> 重复密码：(<a class="red">*</a>) </td>
				<td> <input size="25" type="password" name="password2" value="" onfocus="setfocus('pkey1')" onBlur="setblur('pkey1');checkpass1('pkey1');"></td>
				<td><span id="pkey1" class="blur">请再输入一遍您上面输入的密码</span></td>
			</tr>
			<tr>
				<td> 邮件地址：(<a class="red">*</a>) </td>
				<td> <input size="25" type="text" name="email" value="" onfocus="setfocus('emailkey')" onBlur="setblur('emailkey');checkemail('emailkey');"></td>
				<td><span id="emailkey" class="blur">不可更改，请认真填写。遗忘密码时，可通过此邮箱取回</span></td>
			</tr>	
			<tr class="a4">
				<td> 安全提问：</td>
				<td> <input size="25" type="text" name="questionid" value="" onfocus="setfocus('questkey')" onBlur="setblur('questkey');">
					<select onchange="$('questionid').value=this.value">
						<option value="">无安全提问</option>
						<option value="母亲的名字">母亲的名字</option>
						<option value="爷爷的名字">爷爷的名字</option>
						<option value="父亲出生的城市">父亲出生的城市</option>
						<option value="您其中一位老师的名字">您其中一位老师的名字</option>
						<option value="您个人计算机的型号">您个人计算机的型号</option>
						<option value="您最喜欢的餐馆名称">您最喜欢的餐馆名称</option>
						<option value="驾驶执照的最后四位数字">驾驶执照的最后四位数字</option>
					</select></td>
				<td><span id="questkey" class="blur">请选择安全提问，当密码被遗忘或丢失时，用于找回密码</span></td>
			</tr>
			<tr>
				<td> 回　答：</td>
				<td> <input size="25" type="text" name="answer" value="" onfocus="setfocus('answerkey')" onBlur="setblur('answerkey');checkquest('answerkey');"></td>
				<td><span id="answerkey" class="blur">请填写上面问题的答案。</span></td>
			</tr>
			<tr class="a4">
				<td> 验证码：(<a class="red">*</a>) </td>
				<td> <input size="25" type="text" name="code" value="" onfocus="setfocus('codekey')" onBlur="setblur('answerkey');checkcode('codekey');" onclick="get_Code();">&nbsp;&nbsp;<span id="imgid" style="color:red">点击获取验证码</span></td>
				<td><span id="codekey"  class="blur">请输入右边的数字，如果看不清楚，请请点击刷新验证码</span></td>
			</tr>
			<tr>
				<td>高级选项:</td>
				<td colspan="2"><input name="advshow" type="checkbox"  value="1" onclick="showadv()" class="radio">
				<span id="advance">显示高级用户设置选项</span> </td>
			</tr>
			<tbody style="display:none" id="adv">
			<tr class="a4">
				<td> 性别： </td>
				<td>  
					<input type="radio" name="usersex" class="radio" value="1">男 &nbsp;
					<input type="radio" name="usersex" class="radio" value="2">女 &nbsp;
					<input type="radio" name="usersex" class="radio" value="0" checked> 保密  
				</td>
				<td></td>
			</tr>
			<tr>
				<td> 生日： </td>
				<td><input onclick="showcalendar(event, this)"  value="" name="birthday" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> QQ： </td>
				<td><input type="text" name="qq" size="25"></td>
				<td></td>
			</tr>
			<tr>
				<td> ICQ： </td>
				<td> <input type="text" name="icq" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> Yahoo： </td>
				<td> <input type="text" name="yahoo" size="25"></td>
				<td></td>
			</tr>
			<tr>
				<td> MSN： </td>
				<td><input type="text" name="msn" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> 淘宝旺旺： </td>
				<td> <input type="text" name="taobao" size="25"></td>
				<td></td>
			</tr>
			<tr>
				<td> 支付宝账号： </td>
				<td> <input type="text" name="alipay" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> 来自： </td>
				<td> <input type="text" name="usercity" size="25"></td>
				<td></td>
			</tr>
			</tbody>
		</table>
		<BR><input type="submit" id="lsubmit" class="button" value="提&nbsp;交"/></form>
<%
End Sub 

Sub callreg
	Dim i,Code,UserGroupID,UserInfo,ExtCredits,MustOpen
	Dim username,password,password2,email,questionid,answer,UserMail
	Dim usersex,birthday,City,site,qq,Icq,yahoo,msn,taobao,alipay,Sign,Levelname
	team.ChkPost()
	If Request.Form("formhash")<>Session.Sessionid then
		team.Error "您提交的参数错误,请重新返回刷新后再试 "
	End If
	'判断同一IP注册间隔时间
	If Not Isnull(Session("regtime")) Or CID(team.Forum_setting(10)) > 0 Then
		If DateDiff("s",Session("regtime"),Now()) < CID(team.Forum_setting(10)) Then
			team.Error "系统设置了同一个IP在 "&team.Forum_setting(10)&" 秒内只能注册一次,请误重复提交!"
			Exit Sub
		End If
	End If
	Code = Trim(Request.Form("code"))
	If CID(team.Forum_setting(48)) > 0 Then
		If CID(session("loginnum")) > CID(team.Forum_setting(48)) Then
			if Not team.CodeIsTrue(code) Then
				team.error "验证码错误，请刷新后重新填写。"
			End If
		End If
	End If
	session("loginnum") = session("loginnum") +1
	UserName = LCase(team.Checkstr(Trim(Request.Form("username"))))
	Password = team.Checkstr(Trim(Request.Form("password")))
	Password2 = team.Checkstr(Trim(Request.Form("password2")))
	If UserName = "" or IsNull(UserName) Then
		team.error "用户名不能为空 !"
	End If
	If Not IstrueName(UserName) Then 
		team.Error " 您的用户名有错误的字符。 "
	End If
	TestUName(UserName)
	UserMail = HtmlEncode(Request.Form("email"))
	If Not IsValidEmail(UserMail) Then
		team.error2 "邮件格式错误 !"
	End If
	If Password <> Password2 Then
		team.error2 "两次输入的密码不相同,请重新输入! "
	End If
	Questionid = team.Checkstr(Trim(Request.Form("questionid")))
	Answer = team.Checkstr(Trim(Request.Form("answer")))
	If Questionid<>"" and Not IsNull(Questionid) Then
		If Answer="" then team.Error  "你设置了安全提问，请填写必要的答案选项。"
	End If
	If team.Forum_setting(7)>=1 Then
		UserGroupID = 5
		Levelname="未激活用户||||||0||0"
	Else
		UserGroupID = 27
		Levelname="附小一年级||||||0||0"
	End If
	If Not team.execute("Select * From ["&Isforum&"User] Where UserName='"&UserName&"'").Eof Then
		team.Error " 用户名重复，请重新输入一个用户名。"
	End If
	Dim Mybirthday
	If Len(Request.Form("birthday"))>4 Then
		If Not IsDate(Request.Form("birthday")) Then
			team.Error "生日必须为日期格式"
		Else
			Mybirthday = HtmlEncode(Request.Form("birthday"))
		End If
	Else
		Mybirthday = ""
	End If
	UserInfo = team.Checkstr( Request.Form("qq") &"|"& Request.Form("Icq") &"|"& Request.Form("yahoo") &"|"& Request.Form("msn") &"|"& Request.Form("taobao") &"|"& Request.Form("alipay") )
	ExtCredits= Split(team.Club_Class(21),"|")
	'team.Execute( "insert into ["&Isforum&"User] (UserName,Userpass,UserGroupID,Members,Levelname,RegIP,Usermail,Userhome,UserCity,Question,Answer,Birthday,UserSex,Newmessage,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Postblog,UserInfo,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,UserUp) values('"&username&"','"&MD5(Password,16)&"',"&UserGroupID&",'注册用户','"&Levelname&"','"&Remoteaddr&"','"&UserMail&"','"&HtmlEncode(Request.Form("site"))&"','"&HtmlEncode(Request.Form("usercity"))&"','"&Questionid&"','"&Answer&"','"& Mybirthday &"',"&CID(Request.Form("usersex"))&",0,0,0,0,0,"&SqlNowString&","&SqlNowString&",0,'"&UserInfo&"',"&Cid(Split(ExtCredits(0),",")(2))&","&Cid(Split(ExtCredits(1),",")(2))&","&Cid(Split(ExtCredits(2),",")(2))&","&Cid(Split(ExtCredits(3),",")(2))&","&Cid(Split(ExtCredits(4),",")(2))&","&Cid(Split(ExtCredits(5),",")(2))&","&Cid(Split(ExtCredits(6),",")(2))&","&Cid(Split(ExtCredits(7),",")(2))&",'0|"&Now&"')" )
	Dim AdRs,SQL,RegNum
	'用户短信
	If team.Forum_setting(15) = 1 Then
		Set AdRs= team.execute("Select Top 1 UserName From ["&Isforum&"User] Where UserGroupID = 1")
		If Not (adRs.Eof And AdRs.Bof) Then
			SQL = "insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&Adrs(0)&"','"&UserName&"','"&team.Forum_setting(16)&"',"&SqlNowString&",'注册系统消息')"
			team.Execute(SQL)
			team.execute("Update ["&Isforum&"User] set Newmessage=Newmessage+1 where UserName='"&UserName&"'")
		End If
		AdRs.close:Set AdRs=Nothing
	End If
	'更新最后注册用户
	team.execute("Update ["&Isforum&"Clubconfig] Set Newreguser='"&UserName&"',UserNum=UserNum+1")
	Application(CacheName&"_UserNum") = Application(CacheName&"_UserNum")+1
	Cache.DelCache("Club_Class")
	'发送邮件通知/邮件激活
	RegNum = team.Createpass()
	Dim CallUser,LoginNums,Rs,Mailtopic,Body
	team.Forum_setting(7) = 2
	Select Case CID(team.Forum_setting(7))
		Case 1
			team.execute("Update ["&Isforum&"User] set RegNum='"&RegNum &"' where UserName='"&UserName&"'")
			Mailtopic="用户名注册成功"
			Body = Replace(team.Club_Class(23),"{$clubname}",team.Club_Class(1))
			Body = Replace(Body,"{$username}",UserName)
			Body = Replace(Body,"{$userpass}",PassWord)
			Body = Replace(Body,"{$isregkey}","")
			Body = Replace(Body,"{$emailkey}","论坛程序由 Team5.Cn (DayMoon) 开发&设计，版权所有。")
			Call Emailto(UserMail,Mailtopic,Body)
			CallUser = "<li>注册信息已经发送到您注册的邮箱地址! "
		Case 2
			team.execute("Update ["&Isforum&"User] set RegNum='"&RegNum &"' where UserName='"&UserName&"'")
			Mailtopic="用户名注册成功"
			Body = Replace(team.Club_Class(23),"{$clubname}",team.Club_Class(1))
			Body = Replace(Body,"{$username}",UserName)
			Body = Replace(Body,"{$userpass}",PassWord)
			Body = Replace(Body,"{$isregkey}","    * <a href="&team.Club_Class(2)&"/GetUserInfo.asp?getname="&UserName&"&getid="& RegNum &">点击此链接进行激活</a>")
			Body = Replace(Body,"{$emailkey}","论坛程序由 Team5.Cn (DayMoon) 开发&设计，版权所有。")
			Call Emailto(UserMail,Mailtopic,Body)
			CallUser = "<li>注册信息已经发送到您注册的邮箱地址,请点击信息码激活您的用户信息 ! "
		Case 3
			CallUser = "<li>用户名注册成功,请等待管理员审核您的申请!  "
		Case Else
			CallUser = "<li>用户名：<font color=red>"&Username&"</font><li>密码：<font color=red>"&Password&"</font> "
			'判断Cookies更新目录
			Dim cookies_path_s,cookies_path_d,cookies_path
			cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
			cookies_path_d=ubound(cookies_path_s)
			cookies_path="/"
			For i=1 to cookies_path_d-1
				cookies_path=cookies_path&cookies_path_s(i)&"/"
			Next
			LoginNums = team.Createpass()
			Response.Cookies(Forum_sn)("username")=CodeCookie(username)
			Response.Cookies(Forum_sn)("userpass")=md5(password,16)
			Response.Cookies(Forum_sn)("LoginNum") = LoginNums
			Set Rs=team.execute("Select Max(ID) From ["&IsForum&"User]")
			If Not Rs.Eof Then
				Response.Cookies(Forum_sn)("UserID") = Rs(0)
			End if
			Rs.Close:Set Rs=Nothing
			Response.Cookies(Forum_sn).path = cookies_path
			team.Execute("Update ["&Isforum&"User] Set LoginNum='"&LoginNums&"' Where UserName='"&UserName&"'")
	End Select
	Session("regtime")=now()
	CallUser = CallUser & "<li>注册新用户资料成功<li><a href=Default.asp>返回论坛首页</a>"
	team.error1 CallUser & "<meta http-equiv=refresh content=3;url=Default.asp>"
End Sub

Sub TestUName(s)
	Dim tmp,i,tmp1,u
	If IsNull(team.Club_Class(25)) Or team.Club_Class(25) = "" Then
		Exit sub
	Else
		If Instr(team.Club_Class(25),Chr(13)&Chr(10))>0 Then 
			tmp = Split(team.Club_Class(25),Chr(13)&Chr(10))
			For i = 0 To UBound(tmp)
				If InStr(tmp(i),"*") > 0 Then
					tmp1 = Split(tmp(i),"*")
					For u=0 To UBound(tmp1)
						If tmp1(u) <> "" Then If InStr(s,tmp1(u)) > 0 Then team.Error "您的用户名含有不允许注册的字符，请修改后重新提交"
					Next
				Else
					If InStr(s,tmp(i)) > 0 Then team.Error "您的用户名含有不允许注册的字符，请修改后重新提交"
				End if
			Next
		Else
			tmp = team.Club_Class(25)
			If InStr(tmp,"*") > 0 Then
				tmp1 = Split(tmp,"*")
				For u=0 To UBound(tmp1)
					If tmp1(u) <> "" Then If InStr(s,tmp1(u)) > 0 Then team.Error "您的用户名含有不允许注册的字符，请修改后重新提交"
				Next
			Else
				If InStr(s,tmp) > 0 Then team.Error "您的用户名含有不允许注册的字符，请修改后重新提交"
			End if
		End If
	End If
End sub

Team.Footer
%>
