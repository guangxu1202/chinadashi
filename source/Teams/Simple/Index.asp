<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Header()
Select Case Request("action")
	Case "show"
		show
	Case Else
		Main()
End Select
Footer()

Sub Header() 
	Echo "<link href=""../Skins/teams/bbs.css"" rel=""stylesheet"" type=""text/css"">"
	Echo "<script language=""javascript"" src=""../Js/common.js""></script>"
	Echo "<title>"& team.Club_Class(1) &" - 简约版本</title><center><div id=""divline"">"
	Echo "<div id=""rsstop""><div class=""a1""  style=""padding: 5px;"">"& team.Club_Class(1) &" - <a href=index.asp Style=""color:ffffff"">[简约版本]</a></div></div>"
End Sub

Sub Main()
	Echo "<div id=""rsstop""><div class=""tab1""  style=""padding: 5px;"">"& team.Club_Class(1) &"版块设置</div></div>"
	Echo "<div id=""rssinfo"">"
	Call ForumList()
	Echo "</div>"
End Sub

Sub show()
	Dim Rs,fid,iRs,Bbsname,SQL,Page,u
	Dim Maxpage,PageNum,IsPage,i
	fid = HRF(2,2,"fid")
	Set Rs = team.Execute("Select ID,Pass,Bbsname,toltopic,lookperm From ["&IsForum&"bbsconfig] where hide=0 and id="& fid)
	If ( Rs.Eof and RS.Bof) Then
		Error "此板块不存在或您没有查看此板块的权限"
	Else
		Bbsname = RS(2)
	End If
	If Rs(0) = 0 Then 
		Response.Redirect "../Default.asp?rootid="&fid
	End if
	If Not (RS(4) = ",") Then
		If Instr(RS(4),",") > 0 Then Response.Redirect "../Forums.asp?fid="& fid
	End If
	If Rs(1) <> "" Then
		Response.Redirect "../Forums.asp?fid="& fid
	End if
	RS.Close:Set RS=Nothing
	Echo "<div id=""rsstop""><div class=""tab1""  style=""padding: 5px;"">"& team.Club_Class(1) &" -  <a href='?action=show&fid="&fid&"'>"& iUbb_Code(BbsName) &"</a> - 帖子列表</div></div>"
	IsPage = team.execute("Select Count(*) from ["&IsForum&"Forum] Where deltopic=0 and (Toptopic=2 or Forumid="&fid&" ) ")(0)
	SQL="SELECT ID,Topic,replies,goodtopic,toptopic From ["&IsForum&"Forum] Where deltopic=0 and (Toptopic=2 or Forumid="&fid&") Order By Toptopic DESC,Lasttime DESC" 
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum=SqlQueryNum+1
		Maxpage = 20		'每页分页数
		PageNum = Abs(int(-Abs(IsPage/Maxpage)))	'页数
		Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		iRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
	If Page<2 Then 
		U=0
	Else
		U=page*Maxpage-Maxpage
	End If
	If Not Isarray(iRs) Then
		Echo "<div id=""rssshow""> 暂无帖子记录 </div>"
	Else
		For i=0 To Ubound(iRs,2)
			U= U+1
			Echo "<div id=""rssshow""><div id=""rssleft""> " & U &". </div><div id=""rssright""> "
			If iRs(3,i)=1 Then Echo "[精]" 
			If iRs(4,i)=1 Then Echo "[置顶]" 
			If iRs(4,i)=2 Then Echo "[总置顶]" 
			Echo IIF(CID(team.Forum_setting(65))=1," <a href=""Archiver-"&fid&"-" & iRs(0,i) &".html"">"," <a href=""Archiver.asp?fid="&fid&"&tid=" & iRs(0,i) &""">") & HTMLEncode(iRs(1,i)) &" </a>  - <Span Style=""Color:red"">回复(" & iRs(2,i) &" )</Span> </div></div>"
		Next
	End If
	Echo "<BR /><div id=""rsspage"">"
	Echo " <script language=""JavaScript"">"
	Echo " var pg = new showPages('pg'); "
	Echo " pg.pageCount ="&PageNum&"; "
	Echo " pg.printHtml(1); "
	Echo " </script></div>"
	Echo " "
End Sub

Sub Footer()
	Dim MSCode
	If IsSqlDataBase = 1 Then
		MSCode="SQL"
	Else
		MSCode="ACC"
	End If
	Echo "<div id=""rssfooter"">查看完整版本:  <a href=""../""> [--" & team.Club_Class(1) &" --] <A href=""#"">[-- top --] </a> <BR> Powered by <a target=""_blank"" 	href=""http://www.team5.cn"">" & team.Forum_setting(8) &" - <a href=""Licence.asp""><b style='color:#FF9900'> "& MSCode &"</b></a>  <BR /> Time "& Fix((Timer-Startime)*1000) &" second(s),query: "& SqlQueryNum &"</div>"
	team.HtmlEnd
End Sub

Sub ForumList()
	Dim ShowBbs,i,tmp
	Showbbs = team.BoardList()
	If Isarray(Showbbs) Then
		For i = 0 To UBound(Showbbs,2)
			If Showbbs(3,i) = 0 Then
				Echo "<ul><li>"& iUbb_Code(Showbbs(1,i)) &" "
				Call miniForumList(Showbbs(0,i))
				Echo " </li></ul>"
			End If
		Next
	End if
End Sub

Sub miniForumList(a)
	Dim ShowBbs,i,tmp
	Showbbs = team.BoardList()
	If Isarray(Showbbs) Then
		For i = 0 To UBound(Showbbs,2)
			If Int(Showbbs(3,i)) = Int(a) Then
				Echo "<ul><li> <a href=""?action=show&fid="&Showbbs(0,i)&""">"& iUbb_Code(Showbbs(1,i)) &"</a>" 
				If Showbbs(5,i) > 0 Then  Echo "<img src=""../skins/teams/new.gif"" border=""0"" align=""absmiddle""> "
				Echo "　　　<span style=""color:#666666"">( "&Showbbs(6,i)&" / "&Showbbs(7,i)&" )</span>"
 				If Showbbs(10,i)<>"" Then
					Echo "<FONT COLOR=""red"">[密码验证]</FONT>"
				End If
				Call miniForumList(Showbbs(0,i))
				Echo " </li></ul> "
			End If
		Next
	End if
End Sub
%>