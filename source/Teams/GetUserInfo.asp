<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim Rs,uName,uNum,fid
team.Headers(Team.Club_Class(1) & "- 用户激活页面")
uName = HRF(0,1,"getname")
uNum = HRF(0,1,"getid")
If uName = "" Or IsNull(uName) Then
	team.Error "用户名不能为空。"
End If 
If uNum = "" Or Len(uNum)<16 Then
	team.Error "验证码不能为空。"
End If
Session("sLogin") = CID(Session("sLogin")) + 1
If CID(Session("sLogin"))>5 Then 
	team.Error "您已经连续5次输入了错误的验证密码。"
End If
Set Rs=team.execute("Select RegNum,UserGroupID From ["&IsForum&"User] Where UserName='"&uName&"'")
If Rs.Eof And Rs.Bof Then
	team.Error "系统不存在此用户。"
Else
	If Len(uNum)<>16 Then
		team.Error "您的验证密码错误。"
	ElseIf Trim(uNum) <> Trim(RS(0)) Then
		team.Error "您的验证密码错误。"
	Else
		If Int(Rs(1))=5 Then
			Session("sLogin") = 0
			team.execute("Update ["&IsForum&"User] Set UserGroupID=27,Levelname='附小一年级||||||0||0',Members='注册用户' Where UserName='"&uName&"'")
			team.Error1 "您的帐号已经激活，请等待系统自动返回到登陆页面。<meta http-equiv=refresh content=3;url=login.asp>"
		Else
			team.Error "您的帐号已经被激活，请勿多次运行此程序。<meta http-equiv=refresh content=3;url=login.asp>"
		End if
	End If 
End if
team.footer
%>
