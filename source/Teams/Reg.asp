<!-- #include file="Conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
Dim username,errorchar,i
team.Headers Team.Club_Class(1)& "- �û�ע��"
If team.Forum_setting(4) = 0 then 
	team.error team.Forum_setting(5)
End If
UserName = HRF(2,1,"username")
Dim X1,x2,Fid
If team.UserLoginED Then
	team.error " ��ӭ��������"&TK_UserName&"�����ڽ�ת����ҳ��<meta http-equiv=refresh content=3;url=""Default.asp""> "
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
	Echo " <div id=""regtop""> ����ע��ǰ�����Ķ���"& team.Club_Class(1) &" �� TEAM's��ע�������������� </div>"
	Echo " <div id=""reginfo""> "& Replace(Replace(Ubb_Code(team.Club_Class(13)),"{$clubname}",team.Club_Class(1)),vbcrlf,"<BR>") &" </div>"
    Echo " <div id=""regfoot"">"
	Echo " <li><a href=help.asp?action=custom>���������ӹ���������涨</a></li>"
	Echo " <li><a href=help.asp?action=custom5>������Ϣ�������취 </a></li>"
	Echo " </div><BR><input type=submit value=""ͬ ��""></form><BR>"
End Sub

Sub myCheck
	Dim Tmp,Checkname,i,dmail
	dmail = 1
	If CID(team.Forum_setting(6)) = 0 Then dmail = 0
	%>
		<script src="Js/calendar.js" type="text/javascript"></script>
		<!--����ģ�鿪ʼElseHtml (4)-->
		<table border="0" width="98%" align="center" cellspacing="0" cellpadding="0">
		   <tr>
				<td class="a4"> <span class="bold">
				   <a href="default.asp"><%=team.Club_Class(1)%></a>  &raquo;  
				   </a> &raquo;  <A href="Reg.asp">ע��Э��</a></td>
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
				<td colspan="3" class="a1">ע�� - ��������</td>
			</tr>
			<tr title="�û������������ģ�Ӣ����ĸ�������ִ�Сд�������֣�0-9�����»��ߡ����ַ�����ɣ�������ڵ���3���ַ��������û�����ϵͳ���Զ�����û����Ƿ�Ϸ���" class="a4">
				<td width="20%"> �û�����(<a class="red">*</a>) </td>
				<td width="38%"> <input size="25" type="text" name="username" value="" onfocus="setfocus('ukey')" onBlur="setblur('ukey');checkfocus('ukey');"></td>
				<td width="42%"><span id="ukey" class="blur">�û�������Ϊ3-20���ַ���</span></td>
			</tr>
			<tr>
				<td> �ܡ��룺(<a class="red">*</a>) </td>
				<td> <input size="25" type="password" name="password" value="" onfocus="setfocus('pkey')" onBlur="setblur('pkey');checkpass('pkey');"></td>
				<td><span id="pkey" class="blur">���볤��3-20λ����ĸ�����ִ�Сд����ʹ����ĸ�����ֵ���ϡ�</span></td>
			</tr>
			<tr class="a4">
				<td> �ظ����룺(<a class="red">*</a>) </td>
				<td> <input size="25" type="password" name="password2" value="" onfocus="setfocus('pkey1')" onBlur="setblur('pkey1');checkpass1('pkey1');"></td>
				<td><span id="pkey1" class="blur">��������һ�����������������</span></td>
			</tr>
			<tr>
				<td> �ʼ���ַ��(<a class="red">*</a>) </td>
				<td> <input size="25" type="text" name="email" value="" onfocus="setfocus('emailkey')" onBlur="setblur('emailkey');checkemail('emailkey');"></td>
				<td><span id="emailkey" class="blur">���ɸ��ģ���������д����������ʱ����ͨ��������ȡ��</span></td>
			</tr>	
			<tr class="a4">
				<td> ��ȫ���ʣ�</td>
				<td> <input size="25" type="text" name="questionid" value="" onfocus="setfocus('questkey')" onBlur="setblur('questkey');">
					<select onchange="$('questionid').value=this.value">
						<option value="">�ް�ȫ����</option>
						<option value="ĸ�׵�����">ĸ�׵�����</option>
						<option value="үү������">үү������</option>
						<option value="���׳����ĳ���">���׳����ĳ���</option>
						<option value="������һλ��ʦ������">������һλ��ʦ������</option>
						<option value="�����˼�������ͺ�">�����˼�������ͺ�</option>
						<option value="����ϲ���Ĳ͹�����">����ϲ���Ĳ͹�����</option>
						<option value="��ʻִ�յ������λ����">��ʻִ�յ������λ����</option>
					</select></td>
				<td><span id="questkey" class="blur">��ѡ��ȫ���ʣ������뱻������ʧʱ�������һ�����</span></td>
			</tr>
			<tr>
				<td> �ء���</td>
				<td> <input size="25" type="text" name="answer" value="" onfocus="setfocus('answerkey')" onBlur="setblur('answerkey');checkquest('answerkey');"></td>
				<td><span id="answerkey" class="blur">����д��������Ĵ𰸡�</span></td>
			</tr>
			<tr class="a4">
				<td> ��֤�룺(<a class="red">*</a>) </td>
				<td> <input size="25" type="text" name="code" value="" onfocus="setfocus('codekey')" onBlur="setblur('answerkey');checkcode('codekey');" onclick="get_Code();">&nbsp;&nbsp;<span id="imgid" style="color:red">�����ȡ��֤��</span></td>
				<td><span id="codekey"  class="blur">�������ұߵ����֣�������������������ˢ����֤��</span></td>
			</tr>
			<tr>
				<td>�߼�ѡ��:</td>
				<td colspan="2"><input name="advshow" type="checkbox"  value="1" onclick="showadv()" class="radio">
				<span id="advance">��ʾ�߼��û�����ѡ��</span> </td>
			</tr>
			<tbody style="display:none" id="adv">
			<tr class="a4">
				<td> �Ա� </td>
				<td>  
					<input type="radio" name="usersex" class="radio" value="1">�� &nbsp;
					<input type="radio" name="usersex" class="radio" value="2">Ů &nbsp;
					<input type="radio" name="usersex" class="radio" value="0" checked> ����  
				</td>
				<td></td>
			</tr>
			<tr>
				<td> ���գ� </td>
				<td><input onclick="showcalendar(event, this)"  value="" name="birthday" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> QQ�� </td>
				<td><input type="text" name="qq" size="25"></td>
				<td></td>
			</tr>
			<tr>
				<td> ICQ�� </td>
				<td> <input type="text" name="icq" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> Yahoo�� </td>
				<td> <input type="text" name="yahoo" size="25"></td>
				<td></td>
			</tr>
			<tr>
				<td> MSN�� </td>
				<td><input type="text" name="msn" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> �Ա������� </td>
				<td> <input type="text" name="taobao" size="25"></td>
				<td></td>
			</tr>
			<tr>
				<td> ֧�����˺ţ� </td>
				<td> <input type="text" name="alipay" size="25"></td>
				<td></td>
			</tr>
			<tr class="a4">
				<td> ���ԣ� </td>
				<td> <input type="text" name="usercity" size="25"></td>
				<td></td>
			</tr>
			</tbody>
		</table>
		<BR><input type="submit" id="lsubmit" class="button" value="��&nbsp;��"/></form>
<%
End Sub 

Sub callreg
	Dim i,Code,UserGroupID,UserInfo,ExtCredits,MustOpen
	Dim username,password,password2,email,questionid,answer,UserMail
	Dim usersex,birthday,City,site,qq,Icq,yahoo,msn,taobao,alipay,Sign,Levelname
	team.ChkPost()
	If Request.Form("formhash")<>Session.Sessionid then
		team.Error "���ύ�Ĳ�������,�����·���ˢ�º����� "
	End If
	'�ж�ͬһIPע����ʱ��
	If Not Isnull(Session("regtime")) Or CID(team.Forum_setting(10)) > 0 Then
		If DateDiff("s",Session("regtime"),Now()) < CID(team.Forum_setting(10)) Then
			team.Error "ϵͳ������ͬһ��IP�� "&team.Forum_setting(10)&" ����ֻ��ע��һ��,�����ظ��ύ!"
			Exit Sub
		End If
	End If
	Code = Trim(Request.Form("code"))
	If CID(team.Forum_setting(48)) > 0 Then
		If CID(session("loginnum")) > CID(team.Forum_setting(48)) Then
			if Not team.CodeIsTrue(code) Then
				team.error "��֤�������ˢ�º�������д��"
			End If
		End If
	End If
	session("loginnum") = session("loginnum") +1
	UserName = LCase(team.Checkstr(Trim(Request.Form("username"))))
	Password = team.Checkstr(Trim(Request.Form("password")))
	Password2 = team.Checkstr(Trim(Request.Form("password2")))
	If UserName = "" or IsNull(UserName) Then
		team.error "�û�������Ϊ�� !"
	End If
	If Not IstrueName(UserName) Then 
		team.Error " �����û����д�����ַ��� "
	End If
	TestUName(UserName)
	UserMail = HtmlEncode(Request.Form("email"))
	If Not IsValidEmail(UserMail) Then
		team.error2 "�ʼ���ʽ���� !"
	End If
	If Password <> Password2 Then
		team.error2 "������������벻��ͬ,����������! "
	End If
	Questionid = team.Checkstr(Trim(Request.Form("questionid")))
	Answer = team.Checkstr(Trim(Request.Form("answer")))
	If Questionid<>"" and Not IsNull(Questionid) Then
		If Answer="" then team.Error  "�������˰�ȫ���ʣ�����д��Ҫ�Ĵ�ѡ�"
	End If
	If team.Forum_setting(7)>=1 Then
		UserGroupID = 5
		Levelname="δ�����û�||||||0||0"
	Else
		UserGroupID = 27
		Levelname="��Сһ�꼶||||||0||0"
	End If
	If Not team.execute("Select * From ["&Isforum&"User] Where UserName='"&UserName&"'").Eof Then
		team.Error " �û����ظ�������������һ���û�����"
	End If
	Dim Mybirthday
	If Len(Request.Form("birthday"))>4 Then
		If Not IsDate(Request.Form("birthday")) Then
			team.Error "���ձ���Ϊ���ڸ�ʽ"
		Else
			Mybirthday = HtmlEncode(Request.Form("birthday"))
		End If
	Else
		Mybirthday = ""
	End If
	UserInfo = team.Checkstr( Request.Form("qq") &"|"& Request.Form("Icq") &"|"& Request.Form("yahoo") &"|"& Request.Form("msn") &"|"& Request.Form("taobao") &"|"& Request.Form("alipay") )
	ExtCredits= Split(team.Club_Class(21),"|")
	'team.Execute( "insert into ["&Isforum&"User] (UserName,Userpass,UserGroupID,Members,Levelname,RegIP,Usermail,Userhome,UserCity,Question,Answer,Birthday,UserSex,Newmessage,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Postblog,UserInfo,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,UserUp) values('"&username&"','"&MD5(Password,16)&"',"&UserGroupID&",'ע���û�','"&Levelname&"','"&Remoteaddr&"','"&UserMail&"','"&HtmlEncode(Request.Form("site"))&"','"&HtmlEncode(Request.Form("usercity"))&"','"&Questionid&"','"&Answer&"','"& Mybirthday &"',"&CID(Request.Form("usersex"))&",0,0,0,0,0,"&SqlNowString&","&SqlNowString&",0,'"&UserInfo&"',"&Cid(Split(ExtCredits(0),",")(2))&","&Cid(Split(ExtCredits(1),",")(2))&","&Cid(Split(ExtCredits(2),",")(2))&","&Cid(Split(ExtCredits(3),",")(2))&","&Cid(Split(ExtCredits(4),",")(2))&","&Cid(Split(ExtCredits(5),",")(2))&","&Cid(Split(ExtCredits(6),",")(2))&","&Cid(Split(ExtCredits(7),",")(2))&",'0|"&Now&"')" )
	Dim AdRs,SQL,RegNum
	'�û�����
	If team.Forum_setting(15) = 1 Then
		Set AdRs= team.execute("Select Top 1 UserName From ["&Isforum&"User] Where UserGroupID = 1")
		If Not (adRs.Eof And AdRs.Bof) Then
			SQL = "insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&Adrs(0)&"','"&UserName&"','"&team.Forum_setting(16)&"',"&SqlNowString&",'ע��ϵͳ��Ϣ')"
			team.Execute(SQL)
			team.execute("Update ["&Isforum&"User] set Newmessage=Newmessage+1 where UserName='"&UserName&"'")
		End If
		AdRs.close:Set AdRs=Nothing
	End If
	'�������ע���û�
	team.execute("Update ["&Isforum&"Clubconfig] Set Newreguser='"&UserName&"',UserNum=UserNum+1")
	Application(CacheName&"_UserNum") = Application(CacheName&"_UserNum")+1
	Cache.DelCache("Club_Class")
	'�����ʼ�֪ͨ/�ʼ�����
	RegNum = team.Createpass()
	Dim CallUser,LoginNums,Rs,Mailtopic,Body
	team.Forum_setting(7) = 2
	Select Case CID(team.Forum_setting(7))
		Case 1
			team.execute("Update ["&Isforum&"User] set RegNum='"&RegNum &"' where UserName='"&UserName&"'")
			Mailtopic="�û���ע��ɹ�"
			Body = Replace(team.Club_Class(23),"{$clubname}",team.Club_Class(1))
			Body = Replace(Body,"{$username}",UserName)
			Body = Replace(Body,"{$userpass}",PassWord)
			Body = Replace(Body,"{$isregkey}","")
			Body = Replace(Body,"{$emailkey}","��̳������ Team5.Cn (DayMoon) ����&��ƣ���Ȩ���С�")
			Call Emailto(UserMail,Mailtopic,Body)
			CallUser = "<li>ע����Ϣ�Ѿ����͵���ע��������ַ! "
		Case 2
			team.execute("Update ["&Isforum&"User] set RegNum='"&RegNum &"' where UserName='"&UserName&"'")
			Mailtopic="�û���ע��ɹ�"
			Body = Replace(team.Club_Class(23),"{$clubname}",team.Club_Class(1))
			Body = Replace(Body,"{$username}",UserName)
			Body = Replace(Body,"{$userpass}",PassWord)
			Body = Replace(Body,"{$isregkey}","    * <a href="&team.Club_Class(2)&"/GetUserInfo.asp?getname="&UserName&"&getid="& RegNum &">��������ӽ��м���</a>")
			Body = Replace(Body,"{$emailkey}","��̳������ Team5.Cn (DayMoon) ����&��ƣ���Ȩ���С�")
			Call Emailto(UserMail,Mailtopic,Body)
			CallUser = "<li>ע����Ϣ�Ѿ����͵���ע��������ַ,������Ϣ�뼤�������û���Ϣ ! "
		Case 3
			CallUser = "<li>�û���ע��ɹ�,��ȴ�����Ա�����������!  "
		Case Else
			CallUser = "<li>�û�����<font color=red>"&Username&"</font><li>���룺<font color=red>"&Password&"</font> "
			'�ж�Cookies����Ŀ¼
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
	CallUser = CallUser & "<li>ע�����û����ϳɹ�<li><a href=Default.asp>������̳��ҳ</a>"
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
						If tmp1(u) <> "" Then If InStr(s,tmp1(u)) > 0 Then team.Error "�����û������в�����ע����ַ������޸ĺ������ύ"
					Next
				Else
					If InStr(s,tmp(i)) > 0 Then team.Error "�����û������в�����ע����ַ������޸ĺ������ύ"
				End if
			Next
		Else
			tmp = team.Club_Class(25)
			If InStr(tmp,"*") > 0 Then
				tmp1 = Split(tmp,"*")
				For u=0 To UBound(tmp1)
					If tmp1(u) <> "" Then If InStr(s,tmp1(u)) > 0 Then team.Error "�����û������в�����ע����ַ������޸ĺ������ύ"
				Next
			Else
				If InStr(s,tmp) > 0 Then team.Error "�����û������в�����ע����ַ������޸ĺ������ύ"
			End if
		End If
	End If
End sub

Team.Footer
%>
