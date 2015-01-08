<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim tID,Rs,i,ReList,fid
tID = HRF(2,2,"tid")
team.Headers(Team.Club_Class(1) & "- 查看回收箱或未审核的帖子内容")
Call Main()

Sub Main()
	ManageClass
	Set Rs = team.execute("Select ID,Topic,Content,posttime,UserName,ReList From ["&IsForum&"forum] Where ID=" & tID)
	If RS.eof And Rs.bof Then
		team.Error "系统不存在此帖子。"
	Else
		ReList = Rs(5)
		Echo "<table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=""a2""><tr class=""tab1""><td align=""left"">主题："&Rs(1)&" </td><td  align=""left"">Name: "&Rs(4)&" / Time: "&Rs(3)&"</td></tr><tr class=""a4""><td height=""50"" colspan=""2"">"
		Echo Ubb_Code(Rs(2))
		Echo "</td></tr></table><br />"
	End If
	Rs.Close :Set Rs = Nothing 
	Set Rs = team.execute("Select Username,Content,posttime,ReTopic,ID From  ["&IsForum & ReList &"] Where topicid=" & tID)
	Do While Not Rs.Eof
		Echo "<a name=""RID"&RS(4)&"""><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=""a2""><tr class=""tab1""><td  align=""left"">回复主题："&Rs(3)&" </td><td  align=""left"">Name: "&Rs(0)&" / Time: "&Rs(2)&"</td></tr><tr class=""a4""><td height=""50"" colspan=""2"">"
		Echo Ubb_Code(Rs(1))
		Echo "</td></tr></table><br />"
		Rs.MoveNext
	Loop
	Rs.Close :Set Rs = Nothing
End Sub

Sub ManageClass()
	If Not team.UserLoginED Then
		team.Error " 你未登陆论坛。<meta http-equiv=refresh content=3;url=login.asp> "
	End if
	If Not team.ManageUser Then
		team.Error " 您的权限不够，不能参与论坛管理 。"
	End if
End Sub

team.footer
%>
