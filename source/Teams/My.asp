<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim tID,fID,x1,x2
tID = HRF(2,2,"tid")
fID = HRF(2,2,"fid")
team.Headers(Team.Club_Class(1) & "- 我的收藏夹")

Select Case HRF(2,1,"action")
	Case "getnews"
		Call getnews
	Case "upnews"
		Call upnews
	Case "myopen"
		Call myopen
	Case Else
		Call Main()
End Select

Sub myopen()
	Dim Rs,sRoot,i
	x1="<A href=My.asp>我的快捷菜单</a> "
	Echo team.MenuTitle
	Echo "<div class=""left"" style=""width:20%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""a1""><td>我的快捷菜单</td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp"">我的主题</a></td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp?action=myopen"">我的回复</a></td></tr>"
	Echo "</table>"
	Echo "</div>"
	Echo "<div style=""margin-left:5px;width:100%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""tab1""><td width=""40%"">标题</td><td width=""15%"">论坛</td><td width=""25%"">最后发表</td><td width=""10%"">状态</td></tr>"
	Set Rs=Team.Execute("Select Top 50 B.id,B.Topic,B.Lasttime,B.Views,B.Replies,B.forumid,B.Locktopic,B.CloseTopic,B.Goodtopic,U.Topicid,U.ID,B.LastText From "&IsForum & team.Club_Class(11)&" U Inner Join ["&IsForum&"Forum] B On U.Topicid=B.ID Where B.Auditing=0 and U.UserName='"&TK_UserName&"' and B.UserName<>'"&TK_UserName&"' and B.Deltopic=0 Order By B.Lasttime Desc")
	sRoot = team.BoardList
	Do While Not Rs.eof
		Echo "	<tr class=""tab4""><td align=""left""><a href=""thread.asp?tid="& Rs(0) &"#"& RS(10) &""" target=""_blank"">"& Rs(1) &"</a> (查看:"& RS(3) &"/回复:"&RS(4)&")</td><td align=""left"">"
		For i = 0 To UBound(sRoot,2)
			If Int(sRoot(0,i)) = Int(RS(5)) Then Echo sRoot(1,i)
		Next
		Echo "</td><td align=""left"">"& Rs(2) &" by <a href=""Profile.asp?username="& Split(Rs(11),"$@$")(0) &""">"& Split(Rs(11),"$@$")(0) &"</a></td><td>"
		If Rs(6) = 1 Then
			Echo "关闭"
		ElseIf Rs(7) = 1 Then
			Echo "锁定"
		ElseIf Rs(8) = 1 Then
			Echo "已被加精"
		End if
		Echo "</td></tr>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo "</table>"
	Echo "</div>"
End Sub



Sub Main()
	Dim Rs,sRoot,i
	x1="<A href=My.asp>我的快捷菜单</a> "
	Echo team.MenuTitle
	Echo "<div class=""left"" style=""width:20%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""a1""><td>我的快捷菜单</td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp"">我的主题</a></td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp?action=myopen"">我的回复</a></td></tr>"
	Echo "</table>"
	Echo "</div>"
	Echo "<div style=""margin-left:5px;width:100%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""tab1""><td width=""40%"">标题</td><td width=""15%"">论坛</td><td width=""25%"">最后发表</td><td width=""10%"">状态</td></tr>"
	Set Rs=Team.Execute("Select top 50 ID,Topic,Views,Replies,forumid,Lasttime,Locktopic,CloseTopic,Goodtopic,LastText From ["&IsForum&"Forum] Where Auditing=0 and deltopic=0 and UserName='"& tk_UserName &"' Order By Lasttime Desc")
	sRoot = team.BoardList
	Do While Not Rs.eof
		Echo "	<tr class=""tab4""><td align=""left""><a href=""thread.asp?tid="& Rs(0) &""" target=""_blank"">"& Rs(1) &"</a> (查看:"& RS(2) &"/回复:"&RS(3)&")</td><td align=""left"">"
		For i = 0 To UBound(sRoot,2)
			If Int(sRoot(0,i)) = Int(RS(4)) Then Echo sRoot(1,i)
		Next
		Echo "</td><td align=""left"">"& Rs(5) &" by <a href=""Profile.asp?username="& Split(Rs(9),"$@$")(0) &""">"& Split(Rs(9),"$@$")(0) &"</a></td><td>"
		If Rs(6) = 1 Then
			Echo "关闭"
		ElseIf Rs(7) = 1 Then
			Echo "锁定"
		ElseIf Rs(8) = 1 Then
			Echo "已被加精"
		End if
		Echo "</td></tr>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo "</table>"
	Echo "</div>"
End Sub

Sub getnews
	Dim Rs
	If team.UserLoginED Then
		If tid = 0 Then
			team.Error "帖子参数错误"
		Else
			Set Rs = team.execute("select topic From ["& isforum &"Forum] Where id="& tid)
			If Rs.eof Then
				team.Error "不存在的帖子"
			Else
				If team.execute("select * From ["& isforum &"Favorites] Where ispub="& tid & " and username='"& tk_UserName &"' ").eof Then
					team.Execute("insert into ["&Isforum&"Favorites] (username,name,url,addtime,tag,ispub,look) values ('"& tk_UserName &"','"& Rs(0) &"','thread.asp?tid="& tid &"',"&SqlNowString&",'',"& tid &",0)")
					team.Error1 "您选择的主题已经成功收藏,请等待系统返回.<meta http-equiv=refresh content=3;url="""& Request.ServerVariables("http_referer") &""">"
				Else
					team.Error1 "您过去已经收藏过这个帖子,请等待系统返回.<meta http-equiv=refresh content=3;url="""& Request.ServerVariables("http_referer") &""">"
				End If 
			End If
		End If
	Else
		team.Error "您还未<a href=login.asp>登录</a>论坛"
	End if
End Sub

team.footer
%>
