<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim Admin_Class,Page
Call Master_Us()
Header()
Admin_Class=",10,"
Call Master_Se()

If Cid(Session("UserMember")) <> 1 Then 
	SuccessMsg "对不起，只有管理员方可查看此版的内容 。 "
End if
team.SaveLog ("论坛维护")
Page = HRF(2,2,"Page")
Select Case Request("action")	
	Case "updates"
		Call updates
	Case "runquery"
		Call runquery
	Case "reforums"
		Call reforums
	Case "updatestb"
		Call updatestb
	Case "creattable"
		Call creattable
	Case "reforumdel"
		Call reforumdel
	Case "upfiles"
		Call upfiles
	Case "attachments"
		Call attachments
	Case "deleattachments"
		Call deleattachments
	Case "BakUserbf"
		Call BakUserbf
	Case "SQLUserReadme"
		Call SQLUserReadme
	Case "rebakuserdata"
		Call rebakuserdata
	Case "compressdata"	
		Call compressdata
	Case "clearmsg"
		Call clearmsg
	Case "delmsgok"
		Call delmsgok
	Case "savelog"
		Call savelog
	Case "dellogok"
		Call dellogok
	Case "reforumdelpass"
		Call reforumdelpass
	Case Else
		Call Main
End Select

Sub dellogok
	Dim lConnStr,lConn,ldb,ho
	ldb = MyDbPath & LogDate
	lConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ldb)
	Set lConn = Server.CreateObject("ADODB.Connection")
	lConn.Open lConnStr
	for each ho in Request.form("deleteid")
		lConn.execute("Delete from [SaveLog] Where ID="&ho)
	Next
	SuccessMsg " 选中的操作记录已经被删除，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=savelog>操作记录管理  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=savelog>。 "
End Sub

Sub savelog %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=dellogok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="6">操作记录管理</td>
    </tr>
    <tr class="tab3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio"> 删</td><td>操作人员</td><td>登陆IP</td><td>操作详情</td><td>操作时间</td>
    </tr>
	<%
	Dim Rs,tocou,Maxpage,PageNum,Shows
	Dim SQL,i
	Dim lConnStr,lConn,ldb
	ldb = MyDbPath & LogDate
	lConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ldb)
	Set lConn = Server.CreateObject("ADODB.Connection")
	lConn.Open lConnStr
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	tocou = lConn.execute("Select Count(ID) From [SaveLog]")(0)
	SQL = "Select ID,UserName,IP,Windows,Remark,Logtime From [SaveLog] Order By ID DESC"
	Rs.Open SQL,lConn,1,1,&H0001
	If Rs.Eof And Rs.Bof Then
		Echo "<tr class=""a4""><td colspan=""6"" align=""center""> 暂无内容操作记录 </td></tr></table>"
	Else
		Maxpage = 50
		PageNum = Abs(int(-Abs(tocou/Maxpage)))	'页数
		Page = CheckNum(Page,1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If	
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	For i=0 To Ubound(shows,2)
		Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""deleteid"" value="&Shows(0,i)&" class=""radio""></td><td> <a href=""../Profile.asp?username="& Shows(1,i) &""" target=""_blank"" alt=""点击查看"">"& Shows(1,i) &"</a> </td><td>  "& Shows(2,i) &" </td><td align=""left""> "& Shows(4,i) &" </td><td>"& Shows(5,i) &" </td></tr>"
	Next
	Echo "<tr class=""a4""><td colspan=""6"">"
	Echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td>"
	Echo "<script language=""JavaScript"">"
	Echo"		var pg = new showPages('pg');	"
	Echo"		pg.pageCount = "& PageNum &"	;	"
	Echo"		pg.dispCount = "& tocou &";	"
	Echo"		pg.argName = 'Page';"
	Echo"		pg.printHtml(1); "
	Echo "</script></td></tr></table></td></tr></table><BR/><center><input type=""submit"" name=""onlinesubmit"" value=""提 交""></center></form>"
	Set lConn = Nothing 
End Sub


Sub delmsgok
	Dim ho
	If Request.Form("chkallmsg") = 1 Then
		team.execute("Delete from ["&IsForum&"Message] ")
		SuccessMsg " 所有的短信已经被删除，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=clearmsg>短信管理  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=clearmsg>。 "
	Else
		for each ho in Request.form("deleteid")
			team.execute("Delete from ["&IsForum&"Message] Where ID="&ho)
		Next
		SuccessMsg " 选中的短信已经被删除，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=clearmsg>短信管理  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=clearmsg>。 "
	End if
End Sub

Sub clearmsg %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=delmsgok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="5">短信管理 (删除所有短信<input type="checkbox" name="chkallmsg" class="radio" value="1"> )</td>
    </tr>
    <tr class="tab3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio"> 删 </td><td>发送人</td><td>接受人</td><td>标题</td><td>发送时间</td>
    </tr>
	<%
	Dim Rs,tocou,Maxpage,PageNum,Shows
	Dim SQL,i
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	tocou = team.execute("Select Count(ID) From ["&IsForum&"Message]")(0)
	SQL = "Select ID,author,incept,msgtopic,Sendtime,isbak From ["&IsForum&"Message] Order By Sendtime asc"
	Rs.Open SQL,Conn,1,1,&H0001
	If Rs.Eof And Rs.Bof Then
		Echo "<tr class=""a4""><td colspan=""5"" align=""center""> 短信箱暂无内容 </td></tr></table>"
	Else
		Maxpage = 20
		PageNum = Abs(int(-Abs(tocou/Maxpage)))	'页数
		Page = CheckNum(Page,1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If	
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	For i=0 To Ubound(shows,2)
		Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""deleteid"" value="&Shows(0,i)&" class=""radio""></td><td> <a href=""../Profile.asp?username="& Shows(1,i) &""" target=""_blank"" alt=""点击查看"">"& Shows(1,i) &"</a> </td><td> <a href=""../Profile.asp?username="& Shows(2,i) &""" target=""_blank"">"& Shows(2,i) &"</a> </td><td align=""left""> <a href=""../Msg.asp?action=readmsg&sid="& Shows(0,i) &""" target=""_blank"">"& Shows(3,i) &"</a> "
		If Shows(5,i) = 1 Then
			Echo " - [草稿]"
		End if
		Echo "</td><td>"& Shows(4,i) &" </td></tr>"
	Next
	Echo "<tr class=""a4""><td colspan=""5"">"
	Echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td>"
	Echo "<script language=""JavaScript"">"
	Echo"		var pg = new showPages('pg');	"
	Echo"		pg.pageCount = "& PageNum &"	;	"
	Echo"		pg.dispCount = "& tocou &";	"
	Echo"		pg.argName = 'Page';"
	Echo"		pg.printHtml(1); "
	Echo "</script></td></tr></table></td></tr></table><BR/><center><input type=""submit"" name=""onlinesubmit"" value=""提 交""></center></form>"	
End Sub

Sub deleattachments
	Dim ho,mFso,fPath,Rs,fName
	fPath = "../Images/Upfile/"
	Set mFso = Server.CreateOBject("Scripting.FileSystemObject")
	for each ho in Request.form("deleteid")
		Set Rs = team.execute("Select FileName From ["&IsForum&"Upfile] Where FILEID="&ho)
		If Not Rs.Eof Then
			fName = fPath & Rs(0)
			If  mFso.FileExists(Server.mappath(fName)) Then
				'On Error Resume Next
				mFso.deletefile(Server.mappath(fName))
			End  If
		End if
		team.execute("Delete from ["&IsForum&"Upfile] Where FILEID="&ho)
	Next
	SuccessMsg " 选中的附件已经被删除，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=upfiles>附件管理  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=upfiles>。 "
End Sub

Sub attachments
	Dim inforum,dmincount,dmaxcount,upname,upsize
	Dim Twher,tocou,sql,Maxpage,PageNum,Rs,Shows
	Dim i,tids
	inforum = HRF(1,2,"inforum")
	tids = HRF(1,2,"tids")
	upname = HRF(1,1,"upname")
	upsize = HRF(1,1,"upsize")
	dmaxcount = HRF(1,2,"dmaxcount")
	dmincount = HRF(1,2,"dmincount")
	If upname&"" = "" Then
		Twher = " UserName <>'' "
	Else
		Twher = " UserName Like '% "& upname &" %' "
	End if
	If upsize <> "" Then
		Twher = Twher & " and FileName Like '% "& upsize &" %'"
	End if
	If dmaxcount > 0 Then
		Twher = Twher & " and Upcount>"& dmaxcount &" "
	End if
	If dmincount > 0 Then
		Twher = Twher & " and Upcount<"& dmincount &" "
	End If
	If inforum > 0 Then
		Twher = Twher & " and FID="& Int(inforum) &" "
	End If
	If tids > 0 Then
		Twher = Twher & " and ID="& Int(tids) &" "
	End if
	tocou = team.execute("Select Count(ID) From ["&IsForum&"Upfile] Where "&Twher&" ")(0)
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	SQL = "Select FILEID,ID,FID,UserName,FileName,Types,FileSize,Upcount,ByPowers,Lasttime From ["&IsForum&"Upfile] Where "&Twher&" Order By Lasttime Desc"
	Rs.Open SQL,Conn,1,1,&H0001
	Response.Write "<body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10""><br><br><form method=""post"" action=""?action=deleattachments""><table cellspacing=""1"" cellpadding=""5"" width=""95%"" align=""center"" border=""0"" class=""a2""><tr class=""a3""><td colspan=""8"" align=""center"">本次搜索共找到 <Font color=""red"">"& tocou &"</Font> 条相关附件记录</td></tr><tr class=""tab1""><td><input type=""checkbox"" name=""chkall"" onClick=""checkall(this.form)"" class=""radio""> 删</td><td> 附件名称</td><td>帖子链接</td><td>上传用户</td><td>上传时间</td><td>阅读权限</td><td>下载次数</td><td>主题状态</td></tr>"
	If Rs.Eof And Rs.Bof Then
		Echo "<tr class=""a4""><td colspan=""8"" align=""center""> 对不起，没有找到您要查询的内容 </td></tr></table>"
	Else
		Maxpage = 20
		PageNum = Abs(int(-Abs(tocou/Maxpage)))	'页数
		Page = CheckNum(Page,1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	For i=0 To Ubound(shows,2)
		Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""deleteid"" value="&Shows(0,i)&" class=""radio""></td><td> <a href=""../Images/Upfile/"& Shows(4,i) &""" target=""_blank"" alt=""点击查看"">"& Shows(4,i) &"</a> </td><td> <a href=""../Thread.asp?tid="& Shows(1,i) &""" target=""_blank"">帖子链接</a> </td><td>"& Shows(3,i) &"</td><td>"& Shows(9,i) &" </td><td>"& Shows(8,i) &"</td><td>"& Shows(7,i) &"</td><td>"
		Set Rs = team.execute("Select Deltopic from ["&IsForum&"Forum] Where ID="& Cid(Shows(1,i)))
		If Rs.Eof And Rs.Bof Then
			Echo "已删除"
		Else
			If CID(Rs(0)) = 1 Then
				Echo "已删除"
			Else
				Echo "正常"
			End If
		End if
		Echo "</td></tr>"
	Next
	Echo "<tr class=""a4""><td colspan=""8"">"
	Echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td>"
	Echo "<script language=""JavaScript"">"
	Echo"		var pg = new showPages('pg');	"
	Echo"		pg.pageCount = "& PageNum &"	;	"
	Echo"		pg.dispCount = "& tocou &";	"
	Echo"		pg.argName = 'inforum="&inforum&"&upname="&upname&"&upsize="&upsize&"&dmaxcount="&dmaxcount&"&dmincount="&dmincount&"&Page';"
	Echo"		pg.printHtml(1); "
	Echo "</script></td></tr></table></td></tr></table><BR/><center><input type=""submit"" name=""onlinesubmit"" value=""提 交""></center></form>"
End Sub


Sub upfiles	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=attachments">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="2">搜索附件 【模糊搜索】</td>
    </tr>
    <tr>
      <td class="altbg1">所在论坛:</td>
      <td class="altbg2" align="right">
		<select name="inforum">
			<option value="all"	selected="selected">&nbsp;&nbsp;> 全部</option>
			<%Call BBsList(0)%>
        </select>
      </td>
    </tr>
    <tr>
      <td class="altbg1">帖子ID:</td>
      <td class="altbg2" align="right"><input type="text" name="tids" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">上传用户名:</td>
      <td class="altbg2" align="right"><input type="text" name="upname" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">附件名称:</td>
      <td class="altbg2" align="right"><input type="text" name="upsize" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">被下载次数大于:</td>
      <td class="altbg2" align="right"><input type="text" name="dmaxcount" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">被下载次数小于:</td>
      <td class="altbg2" align="right"><input type="text" name="dmincount" size="40"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="searchsubmit" value="提 交">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub reforumdel
	Dim Rs,tablename
	Tablename = Replace(Request("tablename"),"'","''")
	If Tablename&"" = "" Then
		Successmsg " 请输入回帖表名称。"
	Else
		If not team.Execute("Select ReList From ["&isforum&"forum] where ReList='"&Tablename&"'" ).eof Then
			SuccessMsg " 该表中有对应的主题，您确定要删除么？ <a href=""?action=reforumdelpass&pname="& Tablename &""">如果确定请按下一步</a>"
		Else
			if Ucase(Trim(team.Club_Class(11))) = Ucase(Trim(Tablename)) then
				SuccessMsg("当前正在使用中的数据库不能删除。")
			End If
			team.execute " delete from ["&isforum&"TableList] where TableName='"&Tablename&"' "
			team.Execute " drop table "&Tablename&"  " 
			SuccessMsg " 选中的回帖表已经被删除，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=reforums>回帖表设置  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>。 "
		End if
	End if
End Sub

Sub reforumdelpass
	Dim pname
	pname = Replace(Request("pname"),"'","''")
	if Ucase(Trim(team.Club_Class(11))) = Ucase(Trim(pname)) then
		SuccessMsg("当前正在使用中的数据库不能删除。")
	End If
	team.execute " delete from ["&isforum&"TableList] where TableName='"&pname&"' "
	team.Execute " drop table "& pname &"  " 
	SuccessMsg " 选中的回帖表已经被删除，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=reforums>回帖表设置  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>。 "
End Sub 

Sub creattable
	Dim SQL,tablename
	Tablename = Replace(Request.Form("tablename"),"'","''")
	If Tablename&"" = "" Then
		Successmsg " 请输入回帖表名称。"
	Else
		Sql="CREATE TABLE "&isforum&""&tablename&" ("&_
			"id int IDENTITY (1, 1) NOT NULL ,"&_
			"topicid int NOT NULL ,"&_
			"username varchar(255) NOT NULL ,"&_
			"ReTopic varchar(255) NOT NULL ,"&_
			"content text NOT NULL ,"&_
			"posttime datetime Default "&SqlNowString&" NOT NULL ,"&_
			"postip varchar(255)  NOT NULL ,"&_
			"Reward int NOT NULL ,"&_
			"IsNoName int NOT NULL ,"&_
			"Auditing int NOT NULL ,"&_
			"lock int NULL"&_
			")"
		team.execute(sql)
		team.Execute("insert into ["&isforum&"TableList] (TableName) values ('"&tablename&"')" )
	End if
	SuccessMsg "新回帖表建立成功，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=reforums>回帖表设置  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>。 "
End Sub

Sub updatestb
	Dim tablename
	Tablename = Replace(Request.Form("tablename"),"'","''")
	If Tablename&"" = "" Then
		Successmsg " 请输入回帖表名称。"
	Else
		Cache.DelCache("club_class")
		team.execute("update ["&isforum&"Clubconfig] set ReForumName='"&Tablename&"'")
		Successmsg " 回帖表设置成功 ，请等待系统自动返回到 <a href=Admin_dbmake.asp?action=reforums>回帖表设置  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>。 "
	End if
End Sub

Sub reforums
	If IsSqlDataBase = 1 then
		Successmsg " <BR><BR><BR><div class=""a2"" style='height:50;width:80%'> <ul><BR><li>SQL版本无需设置回帖表。</li></ul></div>"
		Exit Sub
	End If
	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="altbg1">
    <td><br>
      <ul>
        <li>当回帖表数据大量增加时，会导致读取数据变慢，所以添加一个新的回帖表，可以有效加快速度。</li>
      </ul>
      <ul>
        <li> 使用ACCSEE数据库时，当数据库的容量大于100M以后，如果你发现就算添加更多的回帖表也不能显著改变速度，那么推荐您采用SQL数据库。</li>
      </ul></td>
  </tr>
</table>
<BR>
<form method="post" action="?action=updatestb">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan="4">回帖数据表管理 </td>
    </tr>
    <tr class="a3"  align="center">
      <td> 当前表名称 </td>
      <td> 数据量 </td>
      <td> 选定 </td>
      <td> 管理 </td>
    </tr>
    <%
			Dim Rs
			Set Rs=team.execute("select TableName from ["& isforum &"TableList] ")
			Do While Not RS.EOF
				Echo " <tr class=""a4"" align=""center"">"
				Echo " <td bgcolor=""#FFFFFF""> "&RS(0)&"</td>"
				Echo " <td bgcolor=""#F8F8F8""> "&team.execute("Select count(id)from ["&RS(0)&"]")(0)&" </td>"
				Echo " <td bgcolor=""#FFFFFF""> <input type=""radio"" "
				if Ucase(Trim(team.Club_Class(11))) = Ucase(Trim(Rs(0))) then 
					Echo " CHECKED "
				End if
				Echo " value="&RS(0)&" name=""tablename""> </td><td bgcolor=""#F8F8F8""> "
				If Ucase(Trim(Rs(0)))=Ucase("Reforum") Then
					Echo "默认表不能删除"
				Else
					Echo " <a href=""?action=reforumdel&tablename="&RS(0)&""">删除</a> "
				End if
				Echo " </td></tr>"
			RS.MoveNext
		Loop
		Rs.close:Set Rs = Nothing
		%>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="更 新">
  </center>
</form>
<form method="post" action="?action=creattable">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr class="a4">
      <td class="a1" colspan="2"> 添加新的回帖表 </td>
    </tr>
    <tr class="a4">
      <td width="60%"><B>添加新的数据表：</B><br>
        填写你新的回帖表名称，新添加的回帖表名称不能与已经存在的回帖表名称相同，回帖表的名称推荐使用英文字母。</td>
      <td width="40%"><input type="text" size="30" name="tablename" value="newreforum"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="提 交">
  </center>
</form>
<%
End Sub


Sub runquery
	Dim Sqlstr,Sqlstr1,i
	Sqlstr=Request.Form("queries")
	If Sqlstr="" Then
		Successmsg("请输入sql执行语句!")
		Exit Sub
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	On Error Resume Next
	If InStr(Sqlstr,Chr(13)&Chr(10))>0 Then
		Sqlstr1 = Split(Sqlstr,Chr(13)&Chr(10))
		For i=0 To UBound(Sqlstr1)
			Conn.Execute(Sqlstr1(i))
		Next
	Else
		Conn.Execute(Sqlstr)
	End if
	If Err Then
		Err.Clear
		Successmsg "您输入的sql语句有错误 。 <blockquote> "&Sqlstr&" </blockquote>"
	Else
		Successmsg " 成功执行SQL语句 。"
	End If
End Sub

Sub updates %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<table cellspacing="1" cellpadding="4" width="60%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="altbg1">
    <td><br>
      <ul>
        <li> 升级数据库操作有一定的危险性，请小心操作。</li>
		<li> 每行输入一句SQL语句，可以一次性输入多条SQL进行升级。</li>
      </ul>
	  </td>
  </tr>
</table>
<BR>
<form method="post" action="?action=runquery">
  <table cellspacing="1" cellpadding="4" width="60%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">TEAM's 数据库升级 - 请将数据库升级语句粘贴在下面 </td>
    </tr>
    <tr class="altbg1" align="center">
      <td valign="top"><textarea cols="85" rows="10" name="queries"></textarea>
        <br>
        <br>
        注意: 为确保升级成功，请不要修改 SQL 语句的任何部分。</td>
    </tr>
  </table>
  <br>
  <br>
  <center>
    <input type="submit" name="sqlsubmit" value="提 交">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub Main	
	If IsSqlDataBase = 1 then
		Call SQLUserReadme()
		Exit Sub
	End If
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="altbg1">
    <td><br>
      <ul>
        <li> 以下操作需要空间对FSO组件的支持，请查看<a href="Admin_Path.asp?action=discreteness"> <B>组件支持情况</B> </a>确认。</li>
      </ul>
      <ul>
        <li> 进行以下数据库的操作前必须先关闭论坛。</li>
      </ul>
      <ul>
        <li> 经常性的备份数据库可以有效防止因数据库损坏带来的影响（建议每个星期备份一次），备份数据库时必须修改默认的备份路径和备份文件名称，避免因采用默认数据库名称，而导致数据库被黑客下载，从而对密码进行破解的危险。</li>
      </ul>
      <ul>
        <li> 对数据库周期性的进行压缩，有效加快论坛的运行速度（建议每个月压缩一次） ，正确的压缩过程应该是先将数据库备份，然后对备份好的数据库进行压缩，压缩完成后再将压缩数据库还原为当前数据库。请勿将当前的数据库进行压缩，因为那样将存在损坏数据库的危险。 </li>
      </ul></td>
  </tr>
</table>
<BR>
<table cellspacing="1" cellpadding="3" width="95%" align="center">
  <tr>
    <td class="a2"><BR>
      <ul>
        <li> <FONt  COLOR="red">以下操作对数据库潜在危险性，操作失误将造成数据库的损坏，所以请在掌握相应的技巧后再对数据库进行设置。</FONt></li>
      </ul></td>
  </tr>
</table>
<BR>
<form method="post" action="?action=BakUserbf">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan="2">备份数据库 ( 需要FSO支持，FSO相关帮助请看微软网站 )</td>
    </tr>
    <tr class="a3">
      <td width="30%">当前数据库路径(相对路径)： </td>
      <td width="70%"><input type="text" size="60" name="DBpath" value="../<%=db%>"></td>
    </tr>
    <tr class="a4">
      <td width="30%">备份数据库目录(相对路径)：<br>
        如目录不存在，程序将自动创建</td>
      <td width="70%"><input type="text" size="60" name="bkfolder" value="../Databackup"></td>
    </tr>
    <tr class="a4">
      <td width="30%">备份数据库名称(填写名称)：<br>
        如备份目录有该文件，将覆盖，如没有，将自动创建</td>
      <td width="70%"><input type="text" size="60" name="bkDBname" value="teams_Backup.mdb"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="备 份">
  </center>
</form>
<br>
<form method="post" action="?action=rebakuserdata">
  <table cellspacing="1" cellpadding="2" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan=2>恢复数据库 ( 需要FSO支持，FSO相关帮助请看微软网站 )</td>
    </tr>
    <tr class="a3">
      <td width="30%">备份数据库路径(相对)： </td>
      <td width="70%"><input type=text size="60" name="DBpath" value="../DataBackup/teams_Backup.MDB"></td>
    </tr>
    <tr class="a4">
      <td width="30%">目标数据库路径(相对)：</td>
      <td width="70%"><input type=text size="60" name="backpath" value="../<%=db%>"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="恢 复">
  </center>
</form>
<BR>
<form action="?action=compressdata" method="post">
  <table cellspacing="1" cellpadding="2" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan="2">压缩数据库</td>
    </tr>
    <tr>
      <td class="a4" colspan="2"><b>注意：</b><br>
        输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作）</td>
    </tr>
    <tr class="a3">
      <td width="30%">数据库路径： </td>
      <td width="70%"><input size="60" value="../DataBackup/teams_Backup.MDB" name="dbpath"></td>
    </tr>
    <tr class="a4">
      <td width="30%">数据库格式：</td>
      <td width="70%"><input type="radio" value="true" name="boolIs97" id="boolIs97">
        <label for="boolIs97">Access 97</label>
        <input type="radio" value="" name="boolIs97" checked id="boolIs97_1">
        <label for="boolIs97_1">Access 2000、2002、2003</label></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="压 缩">
  </center>
</form>
<BR>
<%
End Sub

Sub compressdata()
	dim dbpath,boolIs97
	dbpath = request("dbpath")
	boolIs97 = request("boolIs97")
	If dbpath <> "" then
		dbpath = server.mappath(dbpath)
		response.write(CompactDB(dbpath,boolIs97))
	End If
End Sub

Sub BakUserbf
		Dim Dbpath,backpath,testConn,bkfolder,bkdbname,fso
		On error resume next
		Dim FileConnStr,Fileconn
		Dbpath=request.Form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.Form("bkfolder")
		bkdbname=request.Form("bkdbname")
		FileConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Dbpath
		Set Fileconn = Server.CreateObject("ADODB.Connection")
		Fileconn.open FileConnStr
		If Err then
			Response.Write Err.Description
			Err.Clear
			Set Fileconn = Nothing
			SuccessMsg("备份的文件并非合法的数据库!")
			Exit Sub
		Else
			Set Fileconn = Nothing
		End If
		Set Fso=server.createobject("scripting.filesystemobject")
		If Fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = true then
				Fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
				MakeNewsDir bkfolder
				Fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			SuccessMsg("备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname &" ")
		Else
			SuccessMsg("找不到您所需要备份的文件!")
		End if
End Sub
Sub rebakuserdata
	Dim Dbpath,backpath,testConn,fso
	Dbpath=request.form("Dbpath")
	backpath=request.form("backpath")
	if dbpath="" then
		SuccessMsg("请输入您要恢复成的数据库全名!")
	else
		Dbpath=server.mappath(Dbpath)
	end if
	backpath=server.mappath(backpath)
	Set testConn = Server.CreateObject("ADODB.Connection")
	On Error Resume Next
	testConn.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Dbpath
	If Err then
		Response.Write Err.Description
		Err.Clear
		Set testConn = Nothing
		SuccessMsg("备份的文件并非合法的数据库!")
		Response.End 
	Else
		Set testConn = Nothing
	End If
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(dbpath) then  					
		fso.copyfile Dbpath,Backpath
		SuccessMsg("数据库恢复成功!")
	else
		SuccessMsg("备份目录下并无您的备份文件!")
	end if
End Sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	Dim fso1
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = true
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f,fso1
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = true
    Set fso1 = nothing
End Function
'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
	Dim fso, Engine, strDBPath,JEt_3X
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(dbPath) then
		fso.CopyFile dbpath,strDBPath & "temp.mdb"
		Set Engine = CreateObject("JRO.JetEngine")
		If boolIs97 = "true" then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
			& "Jet OLEDB:Engine type=" & JEt_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		End If
		fso.CopyFile strDBPath & "temp1.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		fso.DeleteFile(strDBPath & "temp1.mdb")
		Set fso = nothing
		Set Engine = nothing
		SuccessMsg("你的数据库, " & dbpath & ", 已经压缩成功!") & vbCrLf
	Else
		SuccessMsg("数据库名称或路径不正确. 请重试!") & vbCrLf
	End If
End Function

Sub SQLUserReadme()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align="center" width="95%" class="a2">
  <tr>
    <td class="a1">&nbsp;&nbsp;SQL数据库数据处理说明</td>
  </tr>
  <tr>
    <td class="a4"><blockquote> <B>一、备份数据库</B> <BR>
        <BR>
        1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
        2、SQL Server组-->双击打开你的服务器-->双击打开数据库目录<BR>
        3、选择你的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择备份数据库<BR>
        4、备份选项选择完全备份，目的中的备份到如果原来有路径和名称则选中名称点删除，然后点添加，如果原来没有路径和名称则直接选择添加，接着指定路径和文件名，指定后点确定返回备份窗口，接着点确定进行备份 <BR>
        <BR>
        <B>二、还原数据库</B><BR>
        <BR>
        1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
        2、SQL Server组-->双击打开你的服务器-->点图标栏的新建数据库图标，新建数据库的名字自行取<BR>
        3、点击新建好的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择恢复数据库<BR>
        4、在弹出来的窗口中的还原选项中选择从设备-->点选择设备-->点添加-->然后选择你的备份文件名-->添加后点确定返回，这时候设备栏应该出现您刚才选择的数据库备份文件名，备份号默认为1（如果您对同一个文件做过多次备份，可以点击备份号旁边的查看内容，在复选框中选择最新的一次备份后点确定）-->然后点击上方常规旁边的选项按钮<BR>
        5、在出现的窗口中选择在现有数据库上强制还原，以及在恢复完成状态中选择使数据库可以继续运行但无法还原其它事务日志的选项。在窗口的中间部位的将数据库文件还原为这里要按照你SQL的安装进行设置（也可以指定自己的目录），逻辑文件名不需要改动，移至物理文件名要根据你所恢复的机器情况做改动，如您的SQL数据库装在D:\Program Files\Microsoft SQL Server\MSSQL\Data，那么就按照您恢复机器的目录进行相关改动改动，并且最后的文件名最好改成您当前的数据库名（如原来是bbs_data.mdf，现在的数据库是forum，就改成forum_data.mdf），日志和数据文件都要按照这样的方式做相关的改动（日志的文件名是*_log.ldf结尾的），这里的恢复目录您可以自由设置，前提是该目录必须存在（如您可以指定d:\sqldata\bbs_data.mdf或者d:\sqldata\bbs_log.ldf），否则恢复将报错<BR>
        6、修改完成后，点击下面的确定进行恢复，这时会出现一个进度条，提示恢复的进度，恢复完成后系统会自动提示成功，如中间提示报错，请记录下相关的错误内容并询问对SQL操作比较熟悉的人员，一般的错误无非是目录错误或者文件名重复或者文件名错误或者空间不够或者数据库正在使用中的错误，数据库正在使用的错误您可以尝试关闭所有关于SQL窗口然后重新打开进行恢复操作，如果还提示正在使用的错误可以将SQL服务停止然后重起看看，至于上述其它的错误一般都能按照错误内容做相应改动后即可恢复<BR>
        <BR>
        <B>三、收缩数据库</B><BR>
        <BR>
        一般情况下，SQL数据库的收缩并不能很大程度上减小数据库大小，其主要作用是收缩日志大小，应当定期进行此操作以免数据库日志过大<BR>
        1、设置数据库模式为简单模式：打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->双击打开数据库目录-->选择你的数据库名称（如论坛数据库Forum）-->然后点击右键选择属性-->选择选项-->在故障还原的模式中选择“简单”，然后按确定保存<BR>
        2、在当前数据库上点右键，看所有任务中的收缩数据库，一般里面的默认设置不用调整，直接点确定<BR>
        3、<font color="blue">收缩数据库完成后，建议将您的数据库属性重新设置为标准模式，操作方法同第一点，因为日志在一些异常情况下往往是恢复数据库的重要依据</font> <BR>
        <BR>
        <B>四、设定每日自动备份数据库</B><BR>
        <BR>
        <font color="red">强烈建议有条件的用户进行此操作！</font><BR>
        1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器<BR>
        2、然后点上面菜单中的工具-->选择数据库维护计划器<BR>
        3、下一步选择要进行自动备份的数据-->下一步更新数据优化信息，这里一般不用做选择-->下一步检查数据完整性，也一般不选择<BR>
        4、下一步指定数据库维护计划，默认的是1周备份一次，点击更改选择每天备份后点确定<BR>
        5、下一步指定备份的磁盘目录，选择指定目录，如您可以在D盘新建一个目录如：d:\databak，然后在这里选择使用此目录，如果您的数据库比较多最好选择为每个数据库建立子目录，然后选择删除早于多少天前的备份，一般设定4－7天，这看您的具体备份要求，备份文件扩展名一般都是bak就用默认的<BR>
        6、下一步指定事务日志备份计划，看您的需要做选择-->下一步要生成的报表，一般不做选择-->下一步维护计划历史记录，最好用默认的选项-->下一步完成<BR>
        7、完成后系统很可能会提示Sql Server Agent服务未启动，先点确定完成计划设定，然后找到桌面最右边状态栏中的SQL绿色图标，双击点开，在服务中选择Sql Server Agent，然后点击运行箭头，选上下方的当启动OS时自动启动服务<BR>
        8、这个时候数据库计划已经成功的运行了，他将按照您上面的设置进行自动备份 <BR>
        <BR>
        修改计划：<BR>
        1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->管理-->数据库维护计划-->打开后可看到你设定的计划，可以进行修改或者删除操作 <BR>
        <BR>
        <B>五、数据的转移（新建数据库或转移服务器）</B><BR>
        <BR>
        一般情况下，最好使用备份和还原操作来进行转移数据，在特殊情况下，可以用导入导出的方式进行转移，这里介绍的就是导入导出方式，导入导出方式转移数据一个作用就是可以在收缩数据库无效的情况下用来减小（收缩）数据库的大小，本操作默认为您对SQL的操作有一定的了解，如果对其中的部分操作不理解，可以咨询tEAM论坛相关人员或者查询网上资料<BR>
        1、将原数据库的所有表、存储过程导出成一个SQL文件，导出的时候注意在选项中选择编写索引脚本和编写主键、外键、默认值和检查约束脚本选项<BR>
        2、新建数据库，对新建数据库执行第一步中所建立的SQL文件<BR>
        3、用SQL的导入导出方式，对新数据库导入原数据库中的所有表内容<BR>
      </blockquote></td>
  </tr>
</table>
<%
end sub

Sub BBsList(V)
	Dim SQL,ii,RS,i
	Set Rs=Team.Execute("Select ID,BBSname,Followid From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		If RS(2)=0 Then 
			Echo "<optgroup label="""&Rs(1)&""">"
		Else
			Echo "<option value="&RS(0)&">"&String(ii,"　") & RS(1)&"</option>"
		End if
		ii=ii+1
		BBsList RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub

Footer()
%>
