<!--#Include File="../INC/ClsMain.asp"-->
<!--#Include File="../INC/UbbCode.asp"-->
<!--#Include File="../INC/Function.asp"-->
<%
	Dim TK_UserName,TK_Userpass,RemoteAddr,Pages
	Dim iCacheName,mCacheName,iCache,CacheName,Forum_sn
	MyDbPath = "../"
	iCacheName = Server.MapPath("admin.asp")
	iCacheName = Split(iCacheName,"\")
	For iCache = 0 To Ubound(iCacheName)-2
		mCacheName = mCacheName & iCacheName(iCache)
	Next
	CacheName = "team_" & Replace(mCacheName,":","")
	Forum_sn = Replace(CacheName,"_","")
	Set Cache = New Cls_Cache
	Set team = New Cls_Forum
	TK_UserName = DecodeCookie(team.Checkstr(Trim(Request.Cookies(Forum_sn)("UserName"))))
	TK_Userpass = team.Checkstr(Trim(Request.Cookies(Forum_sn)("Userpass")))
	RemoteAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If RemoteAddr = "" Then RemoteAddr = Request.ServerVariables("REMOTE_ADDR")
	RemoteAddr = team.Checkstr(RemoteAddr)
	team.GetForum_Setting
	team.CheckUserLogin
%>