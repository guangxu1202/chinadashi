<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim Thing,myreforum,isre
Call Master_Us()
Header()
Dim Admin_Class
Admin_Class=",4,"
Call Master_Se()
team.SaveLog ("更新论坛统计 [包括：更新论坛统计 ] ")
Myreforum = Request.form("Reforum")
Select Case Request("menu")
	Case "UP_clear"
		Application.Contents.RemoveAll()
		SuccessMsg("系统缓存已经被重建!")
	Case "updateids"
		Dim eid,oid
		Eid = Request.Form("u1")
		oid = Request.Form("u2")
		For i = Eid To oID
			Set Rs=team.execute("Select Count(ID) From Reforum Where TopicID="& i )
			If not Rs.Eof Then
				team.execute("Update ["&IsForum&"Forum] Set replies = "& Rs(0) &" Where ID="& i )
			End If
			Rs.Close:Set Rs=Nothing
		Next 
		SuccessMsg " 更新完成！ "
	Case "update1"
		Dim SQL,SQL1,SQL2,oldday
		SQL2 = "update ["&IsForum&"ClubConfig] set "
		If Request.Form("u1")=1 Then
			Set Rs=team.Execute("Select ID From ["&IsForum&"bbsconfig]")
			Do While not rs.eof
				SQL=Team.Execute("Select count(*) from ["&IsForum&"forum] where deltopic=0 and forumid="&rs(0))(0)
				If SQL>0 Then
					SQL1=Team.Execute("Select sum(replies) from ["&IsForum&"forum] where deltopic=0 and forumid="&rs(0))(0)
				Else
					SQL1=0
				End If
				Team.Execute("update ["&IsForum&"bbsconfig] set toltopic="&SQL&",tolrestore="&SQL+SQL1&" where ID="& rs(0))
				rs.movenext
			Loop
			SQL2 = SQL2 & "PostNum="&team.execute("select count(id) from "&IsForum&"forum where deltopic=0")(0)&","
			SQL2 = SQL2 & "RepostNum="&team.execute("select sum(replies) from "&IsForum&"forum where deltopic=0")(0)&","
		End If
		If Request.Form("u2")=1 Then
			Dim UserNum
			UserNum = team.Execute("Select count(id) from ["&IsForum&"user]")(0)
			SQL2 = SQL2 & "UserNum="&UserNum&","
		End If
		If Request.Form("u3")=1 Then
			If IsSqlDataBase = 1 Then
				today = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff(d,Posttime,"&SqlNowString&")=0 ")(0)
			Else
				today = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff('d',Posttime,"&SqlNowString&")=0 ")(0)
			End If
			If IsSqlDataBase = 1 Then
				today = int(today) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff(d,Posttime,"&SqlNowString&")=0")(0)
			Else
				today = int(today) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff('d',Posttime,"&SqlNowString&")=0")(0)
			End If
			SQL2 = SQL2 & "Today="&Today&""
		End If
		If Request.Form("u4")=1 Then
			If IsSqlDataBase = 1 Then
				Oldday = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff(d,Posttime,"&SqlNowString&")=1 ")(0)
				Oldday = int(Oldday) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff(d,Posttime,"&SqlNowString&")=1")(0)
			Else
				Oldday = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff('d',Posttime,"&SqlNowString&")=1 ")(0)
				Oldday = int(Oldday) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff('d',Posttime,"&SqlNowString&")=1")(0)
			End If
			SQL2 = SQL2 & ",Oldday="&Oldday&""
		End If
		If Request.Form("u5")=1 Then
			SQL2 = SQL2 & ",newreguser='"&team.Execute("Select Top 1 UserName From ["&IsForum&"User] Order by regtime Desc")(0)&"'"
		End If
		Team.Execute(SQL2)
		Application.Contents.RemoveAll()
		SuccessMsg("总论坛数据统计数据更新成功!")
	Case "upnew"
		dim uid,toltopic,tolrestore,rs1,rs,today,trs,rs3,ismytoday,p
		uid=int(request("uid"))
		toltopic=0:tolrestore=0
		Set Rs=team.execute("select toltopic,tolrestore from ["&IsForum&"bbsconfig] where id="& uid)
		If Rs.Eof Then
			SuccessMsg("错误的参数")
		Else
			p =0
			Set rs3 = team.execute("select ID from "&IsForum&"forum where deltopic=0 and forumid="& uid)
			do while not rs3.eof
				If IsSqlDataBase = 1 Then
					ismytoday = team.execute("select count(ID) from "&IsForum&""&team.Club_Class(11)&" where datediff(d,Posttime,"&SqlNowString&")=0 and topicid="& RS3(0))(0)
				else
					ismytoday = team.execute("select count(ID) from "&IsForum&""&team.Club_Class(11)&" where datediff('d',Posttime,"&SqlNowString&")=0 and topicid="& RS3(0))(0)
				end if
				p= p+ismytoday
				rs3.movenext
			loop
			toltopic = team.execute("select count(*) from ["&IsForum&"forum] where deltopic=0 and forumid="& uid)(0)
			tolrestore = team.execute("select sum(replies) from ["&IsForum&"forum] where deltopic=0 and forumid="& uid)(0)
			if tolrestore="" or not isNumeric(tolrestore) then tolrestore=0
			Team.Execute("update ["&IsForum&"bbsconfig] set today="&p&",toltopic="&toltopic&",tolrestore="&tolrestore&" where id="& uid)
			Set Trs=team.Execute("Select Top 1 id,topic,username,lasttime From ["&IsForum&"Forum] Where deltopic=0 and forumid="& uid&" Order by lasttime Desc")
			If Not (Trs.Eof And Trs.Bof) Then
				Team.Execute("update ["&IsForum&"bbsconfig] set Board_Last='<a href=Thread.asp?tid="&Trs(0)&">"&Trs(1)&"</a>$@$"&Trs(2)&"$@$"&Trs(3)&"' where id="& uid)
			Else
				Team.Execute("update ["&IsForum&"bbsconfig] set Board_Last='暂无帖子$@$ - $@$"&SqlNowString&"' where id="& uid)
			End If
		End if
		Application.Contents.RemoveAll()
		SuccessMsg("统计数据更新成功. 发贴数"&toltopic&",回帖数"&tolrestore&",今日贴数"&p&" .")
	Case Else
		UP_main
		Footer()
End Select

Sub UP_main
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing=1 cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1">TEAM's 提示</td>
  </tr>
  <tr>
    <td class="a4" height="30">
	<li>下面有的操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。
	<li>更新总论坛数据  = 将重新计算整个论坛的帖子主题和回复数,今日贴数等等统计数据，建议每隔一段时间运行一次。
	<li>更新分版面数据 = 这里将重新计算指定版面的帖子主题和回复数，最后回复信息等，建议每隔一段时间运行一次。
	</td>
  </tr>
  </table>
<br />

<form name="form1" method="POST" action="?menu=update1">
<table cellspacing="1" cellpadding="6" width="90%" border="0" class="a2" align="center">
 <tr class="a1"><td colspan="2">更新论坛总数据</td></tr>
 <tr class="a4">
    <td>
	<input type="checkbox" name="u1" value="1">主题/回帖数&nbsp;
	<input type="checkbox" name="u2" value="1">用户数&nbsp;
	<input type="checkbox" name="u3" value="1" checked>今日帖&nbsp;
	<input type="checkbox" name="u4" value="1" checked>昨日帖&nbsp;
	<input type="checkbox" name="u5" value="1">最后注册用户&nbsp;
	</td></tr>
	<tr class=a4>
	<td>
	<input size="10" name="Reforum" value="Reforum">&nbsp;&nbsp;<input type="submit" name="Submit"value="更新论坛总数据"> <BR>注:请书写当前回帖表的名称,并选定需要重新统计的选项<br> 
	</td>
  </tr>
</table><br /></form>

<form name="form1" method="POST" action="?menu=updateids">
<table cellspacing="1" cellpadding="6" width="90%" border="0" class="a2" align="center">
 <tr class="a1"><td colspan="2">更新单条帖子数据</td></tr>
 <tr class="a4">
    <td>
	请填写需要重新统计的ID：  <input type="text" name="u1" value="1">&nbsp; &nbsp;<input type="text" name="u2" value="1">
	</td></tr>
	<tr class="a4">
	<td>
	<input type="submit" name="Submit"value="开始更新">
	</td>
  </tr>
</table><br /></form>

<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1">更新分版面数据</td>
  </tr>
	<tr class="a4">
	<td>
		<table cellspacing="1" cellpadding="5" width="99%" border="0" class="a2">
		<%ForumList(0)%>
		</table>
		</td>
  </tr>
</table><br><br>
<form name="form1" method="POST" action="?menu=UP_clear">
<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1" colspan="2">重建系统缓存</td>
  </tr>
  <tr>
    <td class="a4" colspan="2"><li>系统共使用了Application对象 <%=Application.Contents.Count%> 个，Session对象 <%=Session.Contents.Count%> </td>
  </tr> 
	<%
For Each Thing in Application.Contents
	Response.Write "<tr class=a4><td>" & thing & "</td><td>状态："
	If isObject(Application.Contents(Thing)) Then
		Set Application.Contents(Thing) = Nothing
		Application.Contents(Thing) = null
		Response.Write "对象成功关闭"
	ElseIf isArray(Application.Contents(Thing)) Then
		Set Application.Contents(Thing) = Nothing
		Application.Contents(Thing) = null
		Response.Write "数组成功释放"
	Else
		Response.Write Application.Contents(Thing)
		Application.Contents(Thing) = null
	End If
	Response.Write "</td></tr>"
Next
%></table><br>
<center><input type="submit" name="Submit1" value=" 释放缓存 "></form> <br>


<%
End sub

dim ii
ii=0
Sub ForumList(V)
	Dim SQL,RS,Style,S,T,Sty
	Set Rs=team.Execute("Select ID,BbsName,SortNum,Hide,Board_Model From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		Select Case RS(3)
			Case 1
				T="只对游客隐藏"
			Case 2
				T="隐藏"
			Case Else
				T="正常"
		End Select
		If V=0 then	
			Response.Write"<tr class=a4><td width=""5%""></td><td><a target=_blank href=../Forums.asp?Fid="&RS(0)&">┝<b>"&RS(1)&"</b></a>  </td><td> [状态: <b>"&T&"</b>]</a>  </td><td> <a href=""?menu=upnew&uid="&RS(0)&""">更新统计</a></span></td></tr>"
		Else
			Response.Write"<tr class=a4><td width=""5%""></td><td>"&String(ii*2,"　")&" ┕<a target=_blank href=../Forums.asp?Fid="&RS(0)&"><b>"&RS(1)&"</b></a> </td><td> [状态: <b>"&T&"</b>]</a>  </td><td> <a href=""?menu=upnew&uid="&RS(0)&""">更新统计</a></span></td></tr>"
		End If
		ii=ii+1
		ForumList RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub
%>