<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/MD5.asp" -->
<%
team.Headers(Team.Club_Class(1) &" - 密码找回")
Dim X1,X2,Fid,acc
X1="忘记密码"
Select Case Request("action")
	Case "edit"
		Call Edit
	Case Else
		Call Main
End Select

Sub Main
	Echo Replace(Team.ElseHtml (7),"{$weburl}",team.MenuTitle)
End Sub

Sub Edit
	Dim cookies_path,cookies_path_s,cookies_path_d
	Dim i,username,question,answer,UserMail,rs
	UserName = HRF(1,1,"username")
	Question = HRF(1,1,"question")
	Answer = HRF(1,1,"answer")
	UserMail = HRF(1,1,"email")
	If (Cid(Session("Login")) >= Cid(team.Forum_setting(54))) or Request.Cookies(Forum_sn)("OpenLogin")=1 Then
		team.error "您已经连续 "&team.Forum_setting(54)&" 次输入错误信息，系统不允许您在次尝试。"
		cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
		cookies_path_d=ubound(cookies_path_s)
		cookies_path="/"
		For i=1 to cookies_path_d-1
			cookies_path=cookies_path&cookies_path_s(i)&"/"
		Next
		Response.Cookies(Forum_sn)("OpenLogin") = 1
		Response.Cookies(Forum_sn).Expires=Date+1
		Response.Cookies(Forum_sn).path = cookies_path
	Else
		Session("Login") = Session("Login") +1
	End If
	If Trim(UserName) = "" Then
		team.Error "用户名不能为空"
	End if
	If Not IstrueName(UserName) Then 
		team.Error " 您的用户名有错误的字符。 "
	End If
	If Not IsValidEmail(UserMail) Then
		team.Error  "邮件格式错误 !"
	End If	
	Set Rs=Team.Execute("Select Answer,Question,UserMail,UserGroupID From ["&IsForum&"User] Where UserName='"&username&"' ")
	If RS.Eof or Rs.Bof Then
		Rs.Close:Set RS=Nothing
		Team.Error "系统不存在此用户 "
	Else
		If Rs(3) = 1 Or Rs(3) = 2  Then 
			Team.Error "此用户的等级无法通过密码找回功能!"
		End if
		If Trim(RS(2)) <> Trim(UserMail) Then 
			Team.Error "请填写正确的Email地址"
		Else
			Session("Login") = 0
			Dim num1,Title,Mailtopic,Body
			Randomize
			num1= Mid((Rnd*999999),1,6)
			If CID(team.Forum_setting(1))=0 Then
				If Rs(0)&""="" Or Rs(1)&""="" Then
					team.Error " 因为系统不支持邮件发送功能，所以只有用户设置安全提问的前提下才可以使用密码找回功能。"
				End if
				If Trim(Rs(1))<>question or Trim(RS(0))<>answer Then 
					Team.Error "错误的答案!"
				Else
					Team.Execute("Update ["&IsForum&"User] Set userpass='"&Md5(num1,16)&"' Where UserName='"&username&"'")
					team.Error "尊敬的用户 "&username&" ，你的密码已经被修改为[ "&num1&" ] <br> 系统将在 30 秒后自动转入登陆界面。<meta http-equiv=refresh content=30;url=Login.asp>"
				End If
			Else
				If Rs(1)<>"" Or Rs(2)<>"" Then 
					If Trim(Rs(1))<>question or Trim(RS(0))<>answer Then 
						Team.Error "错误的答案!"
					End If
				End if
				Team.Execute("Update ["&IsForum&"User] Set userpass='"&Md5(num1,16)&"' Where UserName='"&username&"'")
				Mailtopic="请查收您的密码! ["&team.Club_Class(1)&"系统消息-Power By Team Borad]"
				Body=""&vbCrlf&"亲爱的"&username&", 您好!"&vbCrlf&""&vbCrlf&"恭喜! 您已经成功地找回您的密码,"& vbCrlf &" 您的新密码为："&num1&" ，请登陆论坛后修改您的密码。 "&vbCrlf&"  非常感谢您使用"&team.Club_Class(3)&"的服务!"&vbCrlf&""&vbCrlf&"　最后, 有几点注意事项请您牢记"&vbCrlf&"1、请遵守《计算机信息网络国际联网安全保护管理办法》里的一切规定。"&vbCrlf&"2、使用轻松而健康的话题，所以请不要涉及政治、宗教等敏感话题。"&vbCrlf&"3、承担一切因您的行为而直接或间接导致的民事或刑事法律责任。"&vbCrlf&""&vbCrlf&""&vbCrlf&"论坛服务由 "&team.Club_Class(1)&"("&team.Club_Class(2)&") 提供　程序制作:TEAM5.CN [By DayMoon]"&vbCrlf&""&vbCrlf&""&vbCrlf&""
				Call Emailto(UserMail,Mailtopic,Body)
				Team.Error " 密码已经发到您注册的邮箱，请注意查收 。<meta http-equiv=refresh content=3;url=Default.asp> "
			End If
		End If
	End If
End Sub
Team.footer
%>