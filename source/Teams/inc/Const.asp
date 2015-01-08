<!--#Include File="ClsMain.asp"-->
<!--#Include File="UbbCode.asp"-->
<!--#Include File="Function.asp"-->
<%
	Dim TK_UserName,TK_Userpass,RemoteAddr,Pages,Forum_sn,CacheName
	CacheName = "team_"& Lcase(Replace(Replace(Replace(Server.MapPath("Default.asp"),"Default.asp","",1,-1,1),":",""),"\",""))
	Forum_sn = Replace(CacheName,"_","")
	Set Cache = New Cls_Cache
	Set team = New Cls_Forum
	TK_UserName = DecodeCookie(team.Checkstr(Trim(Request.Cookies(Forum_sn)("UserName"))))
	TK_Userpass = team.Checkstr(Trim(Request.Cookies(Forum_sn)("Userpass")))
	RemoteAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If RemoteAddr = "" Then RemoteAddr = Request.ServerVariables("REMOTE_ADDR")
	RemoteAddr = team.CheckStr(Trim(Mid(RemoteAddr, 1, 30)))
	team.GetForum_Setting
	team.CheckUserLogin
%>