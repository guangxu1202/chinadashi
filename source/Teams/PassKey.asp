<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim fid
Fid = HRF(2,2,"fid")
team.Headers("密码验证页面")
Select Case Request("action")
	Case "Logins"
		Call Logins
	Case Else
		Call Main
End Select
team.footer
Sub Main
	Dim tmp 
	tmp = "<BR><BR><BR><form action=""?action=Logins&fid="&fid&""" method=""post""><table width=""80%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2""><tr class=""a1""><td> TEAM's提示 </td></tr><tr class=""a4""><td><li>因为本板块设置了密码，所以您需要输入管理员给与的密码才可以访问。</li><li>密码验证板块除了管理员和超级版主可以免除密码输入，其他所有等级用户都需要输入密码才可以访问。</li> </td></tr></table><BR><table width=""80%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2"">"
	tmp = tmp & "<tr class=""tab1""><td colspan=""2""> 密码验证 </td></tr>"
	tmp = tmp & "<tr><td width=""40%"" class=""altbg1""> 请输入密码 </td><td class=""altbg2""><INPUT name=""loginpass"" type=""password"" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';"" class=""colorblur""></td></tr>"
	tmp = tmp & "</table><BR><center><input class=""button"" type=""submit"" name=""submit"" value=""登 录""></center></form><BR><BR>"
	Echo tmp
End Sub

Sub Logins
	Dim Rs,MyPass
	MyPass = HRF(1,1,"loginpass")
	Set Rs=team.Execute("Select Pass From ["&IsForum&"Bbsconfig] Where ID = "& Fid)
	If Rs.Eof Then
		team.Error "板块参数错误。"
	Else
		If Trim(MyPass) = Trim(RS(0)) Then
			Response.Cookies("Class")("LoginKey"& fid) = "1"
			Response.Redirect "Forums.asp?fid="&fid&""
		End If
	End if
End Sub
%>