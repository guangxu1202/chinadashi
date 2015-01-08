<!--#Include File="../inc/ClsMain.asp"-->
<!--#Include File="../inc/Function.asp"-->
<%	
	Dim TK_UserName,TK_Userpass,RemoteAddr,i,Pages
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
	Session.Timeout = 1000		'Session过期
	'**************************************************
	'过程名：Error
	'作  用：显示错误提示信息
	'**************************************************
	'无条件转向
	Public Sub Error(Message)
		Response.Redirect "../Error.asp?Message="&SerVer.URLencode(Message)&""
	End Sub
	Public Sub Error1(Message)
		Response.Redirect "../Error.asp?Message="&SerVer.URLencode(Message)&""
	End Sub
	'弹出提示
	Public Sub Error2(Message)
		Response.Redirect "../Error.asp?Message2="&SerVer.URLencode(Message)&""
	End Sub
	'**************************************************
	'过程名：SuccessMsg
	'作  用：显示成功提示信息
	'参  数：无
	'**************************************************
	Sub SuccessMsg(Msg)
		Dim strSuccess,ComeUrl
		ComeUrl=trim(request.ServerVariables("HTTP_REFERER"))
		strSuccess = strSuccess & "<html><head><title>TEAM's 信息!</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
		strSuccess = strSuccess & "<link href='images/Admin.css' rel='stylesheet' type='text/css'></head><body Style='background-color:#8C8C8C' text='#000000' leftmargin='10' topmargin='10'><br><br>" & vbCrLf
		strSuccess = strSuccess & "<table cellpadding=3 cellspacing=1 border=0 width=80% class='a2' align=center>" & vbCrLf
		strSuccess = strSuccess & "  <tr class='a1'><td height='25'><strong>TEAM's 信息!</strong></td></tr>" & vbCrLf
		strSuccess = strSuccess & "  <tr class='a4'><td height='80' valign='top' align='center'><br> " & Msg & "</td></tr>" & vbCrLf
		strSuccess = strSuccess & "  <tr align='center' class='a3'><td>"
		If ComeUrl <> "" Then
			strSuccess = strSuccess & "<a href='" & ComeUrl & "'>&lt;&lt; 返回上一页</a>"
		Else
			strSuccess = strSuccess & "<a href='javascript:window.close();'>【关闭】</a>"
		End If
		strSuccess = strSuccess & "</td></tr>" & vbCrLf
		strSuccess = strSuccess & "</table>" & vbCrLf
		strSuccess = strSuccess & "</body></html>" & vbCrLf
		Response.Write strSuccess
		Footer()
		Response.end
	End Sub
	'判断权限
	Sub Master_Us()
		If Session("UserMember")="" or Not isNumeric(Session("UserMember")) Then
			Error " 您没有登陆后台的权限 !"
		Else
			If Session("UserMember") = 1 Or Session("UserMember") = 2 Then 
				team.IsMaster = True
			Else
				Error " 您没有登陆后台的权限 !"
			End If
		End If
	End Sub
	Sub Master_Se()
		'Session码错误,返回错误提示
		If InStr(","&Session("Admin_Pass")&",",Admin_Class)=0 Then
			Error("<li>您没有管理本页面的权限。")
		End If
	End Sub
	'顶部
	Sub Header()
		Response.Write"<link href=images/admin.css rel=stylesheet>"
		Response.Write"<title>POWER BY TEAM5.CN</title>"
		Response.Write"<META http-equiv=Content-Type content=text/html;charset=GB2312><script language=""javascript"" src=""../Js/Common.Js""></script>"
	End Sub
	Sub Footer()
		Response.Write"<br><br>"
		Response.Write"<hr size=0 noshade width=80% class=a2>"
		Response.Write"<center><font style='font-size: 11px; font-family: Tahoma, Verdana, Arial'>Powered by <a href=http://www.Team5.cn target=_blank style='color: #000000'><b>"&team.Forum_setting(8)&"</b></a> &nbsp;&copy; 2005, <b><a href=http://www.TEAM5.Cn target=_blank style='color: #000000'>TEAM5.Cn</a> Studio.</b></font></center>"
		Response.Write"</body>"
		Response.Write"</html>"
		Response.End
		Set Rs=Nothing
		Set Rs1=Nothing
		Set Conn=Nothing
	End Sub
%>