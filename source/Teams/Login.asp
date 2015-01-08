<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
Dim x1,x2,fid
fid = HRF(2,2,"fid")
Select Case Request("menu")
	Case "add"
		useradd
	Case "out"
		userout
	Case "eremite"
		eremite
	Case Else
		Userlogin	
End Select

Sub eremite
	'判断Cookies更新目录
	Dim cookies_path_s,cookies_path_d,cookies_path
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	Dim i
	For i=1 to cookies_path_d-1
		cookies_path=cookies_path&cookies_path_s(i)&"/"
	Next
	If HRF(2,2,"upline") = 1 Then
		Response.Cookies(Forum_sn)("Eremite") = 1
	Else
		Response.Cookies(Forum_sn)("Eremite") = 0
	End if
	Response.Cookies(Forum_sn).Expires=date+365
	Response.Cookies(Forum_sn).path = cookies_path
	Cache.DelCache("ShowLines"& CID(fid) )
	Response.Redirect Request.ServerVariables("http_referer")
End sub

Sub Userlogin
	Dim RS,SytyleID
	Dim tmp
	team.Headers("用户登陆")
	Set Rs=Team.Execute("Select StyleName,id From ["&Isforum&"Style] order by ID Asc")
	Do While Not RS.Eof
		SytyleID = SytyleID &  "<option value="&rs(1)&">"&rs(0)&"</option>"
		Rs.Movenext
	Loop
	RS.CLOSE:Set RS=Nothing
	X1="登陆论坛"
	X2 = ""
	tmp = Replace(team.ElseHtml(0),"{$weburl}",team.MenuTitle)
	tmp = Replace(tmp,"{$clubname}",Team.Club_Class(1))
	tmp = Replace(tmp,"{$session}",session.sessionid)
	tmp = Replace(tmp,"{$username}",TK_UserName)
	tmp = Replace(tmp,"{$HTTP_REFERER}",Request.ServerVariables("HTTP_REFERER"))
	tmp = Replace(tmp,"{$username}",TK_UserName)
	tmp = Replace(tmp,"{$SortShowForum}",iif(CID(team.Forum_setting(48))>0,iif(Cid(Session("loginnum"))> CID(team.Forum_setting(48)),"","display:none"),"display:none"))
	tmp = Replace(tmp,"{$sytyleid}",SytyleID)
	Response.Write tmp &"<BR />"
	team.footer
End sub

Sub useradd
	Dim Url,Eremite,styleurl,LoginNums,FUrl
	Dim username,userpass,CookieDate,code,Rs
	Dim cookies_path_s,cookies_path_d,cookies_path
	LoginNums = team.Createpass()
	Session("Ulogin") = 1
	Url = HRF(1,1,"url")
	Eremite = HRF(1,2,"eremite")
	styleurl = HRF(1,1,"styleurl")
	UserName = HRF(1,1,"username")
	UserPass = Md5(HRF(1,1,"userpass"),16)
	CookieDate = Int(Request("CookieDate"))
	Code = Trim(Request.Form("code"))
	'判断Cookies更新目录
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	Dim i
	For i=1 to cookies_path_d-1
		cookies_path=cookies_path&cookies_path_s(i)&"/"
	Next
	If UserName="" or IsNull(UserName) or StrLength(UserName)<2 then 
		team.error "请输入正确的用户名。"
	End if
	If Not IstrueName(UserName) Then 
		team.Error " 您的用户名有错误的字符。 "
	End If
	If CID(team.Forum_setting(54))>0 Then
		If (Cid(Session("Login")) >= Cid(team.Forum_setting(54))) or Request.Cookies(Forum_sn)("OpenLogin")=1 Then
			team.error "您已经连续 "&team.Forum_setting(54)&" 次输入错误密码，系统不允许您登陆。"
			Response.Cookies(Forum_sn)("OpenLogin") = 1
			Response.Cookies(Forum_sn).Expires=Date+1
		Else
			Session("Login") = Session("Login") +1
		End If
	End If 
	If CID(team.Forum_setting(48)) > 0 Then
		If Cid(session("loginnum"))> CID(team.Forum_setting(48)) then
			if Not Team.CodeIsTrue(code) Then
				team.error "验证码错误，请刷新后重新填写。"
			End If
		End If
	End if
	session("loginnum") = session("loginnum") +1
	Set Rs = team.execute("Select ID,UserPass,UserGroupID,Answer,Question,Levelname from ["&Isforum&"User] Where UserName='"&UserName&"'")
	If Rs.Eof and Rs.Bof Then
		team.error "此用户名还未 <a href=""Reg.asp?username="&UserName&""">注册</a> "
	Else
		If Len(Trim(UserPass))<>16 Then
			team.error "您输入的密码错误，您还有 "& 5 - Cid(Session("Login"))&" 次机会重新输入。 "
		ElseIf Len(UserPass)=16 Then
			If Trim(Rs(1)) <> Trim(UserPass) Then
				team.error "您输入的密码有误，您还有 "& 5 - Cid(Session("Login"))&" 次机会重新输入。 "
			End if
		Else
			Session("Login") = 0
		End If
		If Rs(4) <> "" Or Rs(3) <> "" Then
			If Trim(Rs(3))<>Trim(HRF(1,1,"answer")) or Trim(Rs(4)) <>Trim(HRF(1,1,"questionid")) Then
				team.Error "您输入的安全提问或答案错误，请返回后重新输入。"
			End If
		End if
		If Cid(Rs(2)) = 5 Then team.error " 您的帐号尚未激活。<meta http-equiv=refresh content=3;url=""GetUserInfo.asp"">"
		If StyleUrl <> "" Then
			Response.Cookies("Style")("skins") = StyleUrl
		End if
		Select Case CookieDate
	 		Case 1
				Response.Cookies(Forum_sn).Expires=Date+1
	 		Case 2
				Response.Cookies(Forum_sn).Expires=Date+30
	 		Case 3
				Response.Cookies(Forum_sn).Expires=Date+365
		End Select
		Response.Cookies(Forum_sn)("username") = CodeCookie(username)
		Response.Cookies(Forum_sn)("userpass") = UserPass
		Response.Cookies(Forum_sn)("UserID") = Rs(0)
		Response.Cookies(Forum_sn)("LoginNum") = LoginNums
		Response.Cookies(Forum_sn)("Eremite") = Eremite
		Response.Cookies(Forum_sn).path = cookies_path
		Session(CacheName&"_UserLogin") = ""
		team.Execute("Update ["&Isforum&"User] Set LoginNum='"&LoginNums&"' Where UserName='"&UserName&"'")
		If team.UserLoginED = False Then
			team.Execute("Delete From ["&Isforum&"Online] Where Sessionid ="&Ccur(Session.SessionID))
		End if
		Rs.Close:Set Rs=Nothing
		If Url = "" Then
			Url = "Default.asp"
		ElseIf Instr(Url,"Error.asp")>0 Then
			Url = "Default.asp"
		ElseIf Instr(Url,"Login.asp")>0 Then
			Url = "Default.asp"
		Else
			Url = Url
		End If
		team.error1 " 您已经成功登陆论坛，您可以选择进入以下的页面或等待系统自动返回先前访问的页面<li><a href=""default.asp"">论坛首页</a><li><a href="""& Url &""">快速进入</a><li><a href="""&Url&""">先前访问的页面</a><meta http-equiv=refresh content=3;url="""&Url&""">"
	End if
End Sub

Sub userout
	Dim Msg 
	Team.execute("delete from ["&IsForum&"online] where sessionid="& session.sessionid)
	Team.execute("delete from ["&IsForum&"online] where UserName='"&TK_UserName&"' ")
	'判断Cookies更新目录
	Dim cookies_path_s,cookies_path_d,cookies_path
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	Dim i
	For i=1 to cookies_path_d-1
		cookies_path=cookies_path&cookies_path_s(i)&"/"
	Next
	Response.Cookies(Forum_sn).path = cookies_path
	Response.Cookies(Forum_sn)("username")=""
	Response.Cookies(Forum_sn)("userpass")=""
	Response.Cookies(Forum_sn)("UserMember")=""
	Response.Cookies(Forum_sn)("UserID")="0"
	Session("UserMember")= ""
	Session("Admin_Pass")=""
	Cache.DelCache("ForumUserOnline")
	session.abandon()
	Team.Error1 "您已退出论坛，现在将以游客身份转入首页。<meta http-equiv=refresh content=3;url=Default.asp>"
End Sub

%>