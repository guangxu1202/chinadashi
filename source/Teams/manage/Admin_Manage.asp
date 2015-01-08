<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim ii,ID
Dim Admin_Class
Call Master_Us()
Header()
ii=0:ID=Request("id")
Admin_Class=",3,"
Call Master_Se()
team.SaveLog ("快捷管理 [包括：快捷管理，帖子审核] ")
Select Case Request("Action")
	Case "deltopicok"
		Call deltopicok
	Case "delforumok"
		Call delforumok
	Case "delusertopicok"
		Call delusertopicok
	Case "delliketopicok"
		Call delliketopicok
	Case "delretopicok"
		Call delretopicok
	Case "deluserretopicok"
		Call deluserretopicok
	Case "UniForum"
		Call UniForum	'合并版块
	Case "Forumsmerge"
		Call Forumsmerge
	Case "uniteok"
		Call uniteok
	Case "readkey"
		CAll readkey
	Case "readkeyok"
		Call readkeyok
	Case Else
		Call Main()
End Select

Sub readkeyok
	Dim ho
	For each ho in request.form("checktid")
		team.execute("Update ["&Isforum&"Forum] Set Auditing=0 Where ID="&ho)
	Next
	team.SaveLog ("帖子审核完成")
	SuccessMsg " 帖子审核完成，请等待系统自动返回到 <a href=Admin_Manage.asp?Action=readkey>帖子审核 </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp?Action=readkey>。 "	
End sub

Sub readkey%>
	<br>
	<br>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?action=readkeyok">
	<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr class="a1">
	 <td>技巧提示</td>
	</tr>
	<tr class="a4">
	 <td><br>
      <ul>
        <li>您可以在此审核用户发表的帖子或回帖。
		<li>只有在板块设置了<B>发帖审核</B>功能，贴子才需要审核。通过了审核的贴子才可以正常显示在板块里面 。
      </ul></td>
	 </tr>
	</table>
	<BR>
	<table cellspacing="1" cellpadding="3" border="0" width="95%" align="center" class="a2">
	<tr class="tab1"> 
		<td align="center" width="10%"><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio">审核</td> <td align="center">标题</td>
	</tr>
	<%
	Dim Rs
	Set Rs=team.execute("Select * From ["&IsForum&"Forum] Where Auditing=1 and Deltopic=0")
	Do While Not Rs.Eof
			Echo "<tr class=""a4"">"
			Echo "	<td align=""center""><input type=""checkbox"" name=""checktid"" class=""radio"" value="""&RS(0)&"""></td>"
			Echo "	<td> <a href=""../SeeDeltop.asp?tid="&Rs("ID")&""" target=""_blank""> "& Rs("topic") &"</a> </td>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo "</table><br><center><input type=""submit"" name=""onlinesubmit"" value=""提 交""></center></form>"
End Sub

Sub uniteok
	Dim Source,Target,UserTips
	Source = Request.Form("source")
	Target = Request.Form("target")
	If Source="" or Target="" Then Error2 "源论坛或目标论坛不能为空!"
	If CID(Source) = CID(Target) Then Error2 "源论坛或目标论坛不能相同!"
	if Request.Form("postname") <> "" then 
		UserTips =" and UserName='"&HtmlEncode(Trim(Request.Form("postname")))&"'"
	End If
	If IsSqlDataBase=1 Then
		team.Execute( "Update ["&Isforum&"forum] Set Forumid="&Target&" Where Forumid="&Source&" and Datediff(d,Posttime, " & SqlNowString & ") > " & Cid(Request.Form("posttime"))&" "&UserTips&" ")
	Else
		team.Execute( "Update ["&Isforum&"forum] Set Forumid="&Target&" Where Forumid="&Source&" and Datediff('d',Posttime, " & SqlNowString & ") > " & Cid(Request.Form("posttime"))&" "&UserTips&" ")
	End If
	SuccessMsg "论坛移动成功，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> "
End Sub

Sub Forumsmerge
	Dim Source,Target
	Source = Request.Form("source")
	Target = Request.Form("target")
	If Source="" or Target="" Then Error2 "源论坛或目标论坛不能为空!"
	If CID(Source) = CID(Target) Then Error2 "源论坛或目标论坛不能相同!"
	team.execute("Update "&IsForum&"forum Set Forumid="&Target&" Where Forumid="&Source)
	team.execute("Delete from "&IsForum&"Bbsconfig Where id="&Source)
	SuccessMsg "论坛合并成功,源论坛已经被删除，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> "
End Sub

Sub deluserretopicok
	Call Master_Se()
	Dim BbsID,FindBoard,BoardName,Rs
	BbsID = Request.Form("bbsid")
	If Request.Form("postname") = "" Then 
		Error2 "您没有输入用户名。"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then
			Set Rs=Team.execute("Select ID From "&Isforum&"Forum Where forumid="& BbsID)
			While Not RS.EOF
				team.Execute( "Delete From ["&Isforum&""& Request.Form("Reforumname") &"] Where UserName = '"&HtmlEncode(Trim(Request.Form("postname")))&"' and Topicid="&RS(0) )
				Rs.MoveNext
			Wend
		Else
			team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where UserName = '"&HtmlEncode(Trim(Request.Form("postname")))&"' ")
		End If
		
		SuccessMsg " 已经将回帖表 [ "&Request.Form("Reforumname")&" ] 里面用户"&Request.Form("postname")&" 发表的回帖删除，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delretopicok
	Call Master_Se()
	Dim BbsID,FindBoard,Rs
	BbsID = Request.Form("bbsid")
	If Request.Form("posttime") = "" or ( Not isNumeric(Request.Form("posttime")) ) Then 
		Error2 "日期必须为数字。"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			Set Rs=Team.execute("Select ID From "&Isforum&"Forum Where forumid="& BbsID)
			While Not RS.EOF
				If IsSqlDataBase=1 Then
					team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff(d,Posttime, " & SqlNowString & ") > " & Request.Form("posttime")&" And Topicid="&RS(0) )
				Else
					team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff('d',Posttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" And Topicid="& RS(0) )
				End If
				Rs.MoveNext
			Wend
		Else
			If IsSqlDataBase=1 Then
				team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff(d,Posttime, " & SqlNowString & ") > " & Request.Form("posttime")&" ")
			Else
				team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff('d',Posttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" ")
			End If
		End If
		SuccessMsg " 已经将回帖表 [ "&Request.Form("Reforumname")&" ]里面 "&request("posttime")&" 天以前的回帖删除，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delliketopicok
	Call Master_Se()
	Dim BbsID,FindBoard,BoardName
	BbsID = Request.Form("bbsid")
	If Request.Form("topic") = "" Then 
		Error2 "您没有输入字符。"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
		End If
		team.Execute( "Delete From ["&Isforum&"forum] Where Topic Like  '%"&HtmlEncode(Trim(Request.Form("topic")))&"%'  "& FindBoard )
		SuccessMsg " 已经将标题里包含有 "&Request.Form("topic")&"  的主题删除，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delusertopicok
	Call Master_Se()
	Dim BbsID,FindBoard,BoardName
	BbsID = Request.Form("bbsid")
	If Request.Form("postname") = "" Then 
		Error2 "您没有输入用户名。"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
			BoardName = " 在版块 <a href=../BoardList.asp?ID="& BbsID &" "
		End If
		team.Execute( "Delete From ["&Isforum&"forum] Where UserName = '"&HtmlEncode(Trim(Request.Form("postname")))&"'  "& FindBoard )
		SuccessMsg " 已经将 "&Request.Form("postname")&"  "&BoardName&" 发表的主题删除，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delforumok
	Call Master_Se()
	Dim BbsID,FindBoard
	BbsID = Request.Form("bbsid")
	If Request.Form("posttime") = "" or ( Not isNumeric(Request.Form("posttime")) ) Then 
		Error2 "日期必须为数字。"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
		End If
		If IsSqlDataBase=1 Then
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff(d,Lasttime, " & SqlNowString & ") > " & Request.Form("posttime")&" "& FindBoard )
		Else
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff('d',Lasttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" "& FindBoard )
		End If
		SuccessMsg " 已经将"&request("posttime")&"天没有更新过的主题删除，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub deltopicok
	Call Master_Se()
	Dim BbsID,FindBoard
	BbsID = Request.Form("bbsid")
	If Request.Form("posttime") = "" or ( Not isNumeric(Request.Form("posttime")) ) Then 
		Error2 "日期必须为数字。"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
		End If
		If IsSqlDataBase=1 Then
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff(d,Posttime, " & SqlNowString & ") > " & Request.Form("posttime")&" "& FindBoard  )
		Else
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff('d',Posttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" "& FindBoard  )
		End If
		SuccessMsg " 已经将"&request("posttime")&"天以前的主题删除，请等待系统自动返回到 <a href=Admin_Manage.asp>快捷管理</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub Main()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2" >
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
    <td><BR>
      <ul>
        <li>请在论坛关闭的情况下进行下列操作，并在操作完成后 <a href=Admin_Update.asp><B>更新论坛统计</B></a> 信息。</li>
      </ul>
      <ul>
        <li>此操作是不可逆转的，所以推荐在 <a href=Admin_MDBS.asp><B>备份好数据库</B></a> 以后操作，并确认您的每一个步骤。</li>
      </ul>
      <ul>
        <li>此操作的时间将因你的数据库大小而异，数据库越大，时间越长。</li>
      </ul>
      <ul>
        <li> 本论坛共有论坛数：<B><%=team.execute("Select count(id)from bbsconfig")(0)%></B> ，
          主贴数：<B><%=team.execute("Select count(id)from forum")(0)%></font></B> ，
          当前回帖表数：<B><%=team.execute("Select count(id)from ["&team.Club_Class(11)&"]")(0)%></B>，如果您论坛首页的统计和当前的统计有误差，请 <a href=Admin_Update.asp><B>更新论坛统计</B> 。</li>
      </ul></td>
  </tr>
</table>
<BR>
<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1" colspan="3">批量删除主题</td>
  </tr>
  <form method="post" action="?Action=deltopicok">
    <tr class="a4">
      <td width="40%"> 删除 <INPUT size="3" name="posttime" value="180"> 天以前的主题</td>
      <td width="40%"><select name="bbsid">
          <option value="">所有论坛</option>
          <%ForumList_Sel(0)%>
        </select></td>
      <td width="20%"><input type="submit" value=" 确 定 "></td>
  </tr>
</form>
<form method="post" action="?Action=delforumok">
    <tr class="a3">
      <td> 删除<INPUT size="3" name="posttime" value="180"> 天没有更新的主题</td>
      <td><select name="bbsid">
          <option value="">所有论坛</option>
          <%ForumList_Sel(0)%>
        </select></td>
      <td><input type="submit" value=" 确 定 "></td>
  </tr>
   </form>
  <form method="post" action="?Action=delusertopicok">
  <tr class="a4">
	<td>删除 <input size="10" name="postname"> 发表的所有主题 </td>
    <td> <select name="bbsid">
			<option value="">所有论坛</option>
			<%ForumList_Sel(0)%>
			</select>
	</td>  
    <td><input type="submit" value=" 确 定 "></td>
  </tr>
  </form>
  <form method="post" action="?Action=delliketopicok">
  <tr class="a3">
	<td>删除标题里包含有 <input size="10" name="topic"> 的所有主题</td>
    <td><select name="bbsid">
			<option value="">所有论坛</option>
			<%ForumList_Sel(0)%>
			</select>
	</td>
    <td><input type="submit" value=" 确 定 "></td>
   </tr>
</form>
</table>
<BR/>

<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr  class="a1">
    <td colspan="3"> 批量删除回帖 </td>
  </tr>
  <form method="post" action="?Action=delretopicok">
    <tr class="a4">
      <td width="40%"> 删除 <INPUT size="3" name="posttime" value="180"> 天以前的回帖</td>
      <td width="40%"><select name="bbsid">
						<option value="">所有论坛</option>
						 <%ForumList_Sel(0)%>
						 </select> - <select name="Reforumname">
										<option value="ReForum">请选择回帖表</option>
<%	Dim Value,i,Rs1
	Set Rs1 = Team.Execute(" Select id,TableName From TableList ")
	If Not Rs1.Eof Then
		Value = Rs1.GetRows(-1)
	End If
	Rs1.Close:Set Rs1=Nothing
	If IsArray(Value) Then
		For i=0 To Ubound(Value,2)
			Echo "<option value="&Value(1,i)&">"&Value(1,i)&"</option>"
		Next
	End If
	%></select>
      </td>
      <td width="20%"><input type="submit" value=" 确 定 "></td>
    </tr>
  </form>
  <form method="post" action="?Action=deluserretopicok">
  <tr Class="a3">
	<td>删除 <input size="10" name="postname"> 发表的所有回帖 </td> 
    <td>	<select name="bbsid">
				<option value="">所有论坛</option>
				<%ForumList_Sel(0)%>
			</select> -
			<select name="Reforumname">
				<option value="ReForum">请选择回帖表</option>
				<%
			If IsArray(Value) Then
				For i=0 To Ubound(Value,2)
					Echo "<option value="&Value(1,i)&">"&Value(1,i)&"</option>"
				Next
			End If	%>
		</select>
    </td>  
    <td><input type="submit" value=" 确 定 ">
</tr>
</form>
</table>
<BR>
  <center>
    <input type="submit" name="submit" value="提 交" onclick="{if(confirm('您确定要删除论坛么?')){return true;}return false;}">
  </center>
</form>
<form method="post" action="?Action=uniteok">
  <table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
    <tr class="a1">
      <td colspan="3">移动论坛 - 将指定论坛的帖子按照条件筛选转入目标论坛，同时保留源论坛</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">源论坛:</td>
      <td bgcolor="#FFFFFF" width="60%"  align="left">　
        <select name="source">
          <option value="">┝ 请选择</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">目标论坛:</td>
      <td bgcolor="#FFFFFF" width="60%"  align="left">　
        <select name="target">
          <option value="">┝ 请选择</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">日期设定:</td>
      <td bgcolor="#FFFFFF" width="60%" align="left">　仅移动
        <input size="2" name="posttime" value="0">
        天前的帖子</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">用户设定:</td>
      <td bgcolor="#FFFFFF" width="60%"  align="left">　仅移动
        <input size="8" name="postname">
        发表的帖子</td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="submit" value="提 交">
  </center>
</form>
<form method="post" action="?Action=Forumsmerge">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="3">合并论坛 - 源论坛的帖子全部转入目标论坛，同时删除源论坛</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">源论坛:</td>
      <td bgcolor="#FFFFFF" width="60%" align="left">　
        <select name="source">
          <option value="">┝ 请选择</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">目标论坛:</td>
      <td bgcolor="#FFFFFF" width="60%" align="left">　
        <select name="target">
          <option value="">┝ 请选择</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="submit" value="提 交">
  </center>
</form>
<br>
<%
End Sub

Sub DelForums%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?Action=ServerDelForum&ID=<%=request("ID")%>">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan=5>TEAM's 提示</td>
    </tr>
    <tr align="center">
      <td bgcolor="#FFFFFF"><br>
        <br>
        <br>
        本操作不可恢复，您确定要删除该论坛，清除其中帖子和附件吗?<br>
        注意: 删除论坛并不会更新用户发帖数和积分<br>
        <br>
        <br>
        <br>
        <input type="submit" name="forumsubmit" value=" 确 定 ">
        &nbsp;
        <input type="button" value=" 取 消 " onClick="history.go(-1);"></td>
    </tr>
  </table>
</form>
<br>
<%
End Sub

Sub ForumList_Sel(V)
	Dim SQL,ii,RS,W
	Set Rs=Team.Execute("Select ID,BBSname,Followid From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		W="　┕ "
		If V = 0 Then W="┝ "
		Response.Write "<option value="&RS(0)&""
		Response.Write ">"&String(ii,"　")&""&W&""&RS(1)&"</option>"
		ii=ii+1
		ForumList_Sel RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub
%>
