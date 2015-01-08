<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim Fso,tid,tmp
tid = HRF(2,2,"tid")
tmp = Header
tmp = tmp & Main
tmp = tmp & Footer
Echo tmp

Function Header()
	Dim tmp
	tmp = "<link href=""../Skins/teams/bbs.css"" rel=""stylesheet"" type=""text/css"">"
	tmp = tmp &  "<script language=""javascript"" src=""../Js/common.js""></script>"
	tmp = tmp &  "<div id=""divline""><div id=""rsstop""><div class=""a1""  style=""padding: 5px;"">"& team.Club_Class(1) &" - <a href=index.asp Style=""color:ffffff"">[简约版本]</a></div></div>"
	Header = Tmp
End Function

Function footer()
	Dim MSCode,tmp
	If IsSqlDataBase = 1 Then
		MSCode="SQL"
	Else
		MSCode="ACC"
	End If
	tmp = "<div id=""rssfooter"">查看完整版本: [-- <a href=""../""> " & team.Club_Class(1) &" --] <A href=""#"">[-- top --] </a> <BR> Powered by <a target=""_blank"" 	href=""http://www.team5.cn"">" & team.Forum_setting(8) &" - <a href=""Licence.asp""><b style='color:#FF9900'> "& MSCode &"</b></a>  <BR /> Time "& Fix((Timer-Startime)*1000) &" second(s),query: "& SqlQueryNum &" "
	Echo " </div>"
	footer = tmp
End Function


Function Main()
	Dim Rs,fid,iRs,Bbsname,SQL,Page,ReList
	Dim Maxpage,PageNum,IsPage,i,u,tmp,iGUid
	Page = Request.QueryString("page")
	fid = HRF(2,2,"fid")
	Set Rs = team.Execute("Select ID,Pass,Bbsname,toltopic,lookperm,followid From ["&IsForum&"bbsconfig] where hide=0 and id="& fid)
	If ( Rs.Eof and RS.Bof) Then
		Error "此板块不存在或您没有查看此板块的权限"
	Else
		Bbsname = RS(2)
	End If
	If Int(Rs(5)) = 0 Then 
		Response.Redirect IIF(CID(team.Forum_setting(65))=1,"../Default-"&fid&".html","../Default.asp?rootid="&fid)
	End if
	If Not (RS(4) = ",") Then
		If Instr(RS(4),",") > 0 Then 
			Response.Redirect IIF(CID(team.Forum_setting(65))=1,"../Thread-"& fid &".html","../Thread.asp?tid="& fid)
		End If
	End If
	If Rs(1) <> "" Then
		Response.Redirect IIF(CID(team.Forum_setting(65))=1,"../Thread-"& fid &".html","../Thread.asp?tid="& fid)
	End if
	RS.Close:Set RS=Nothing
	Set Rs = team.execute("SELECT Topic,ReList,Goodtopic,toptopic,Username,Posttime,Content,Replies From ["&IsForum&"Forum] Where deltopic=0 and id="& tid)
	If Not (Rs.Eof And Rs.Bof) Then
		tmp = "<title> "& Rs(0) &" - Power By team board</title>"
		tmp = tmp &  "<div id=""rsstop""><div class=""a4""  style=""padding: 5px;"">标题： "& RS(0) &" - " 
		If Rs(2)=1 Then tmp = tmp &  "[精]" 
		If Rs(3)=1 Then tmp = tmp &  "[置顶]" 
		If Rs(3)=2 Then tmp = tmp &  "[总置顶]" 
		tmp = tmp & "</div></div>" 
		If Page < 2 Then
			tmp = tmp &  " <div id=""rssshow""><div id=""showarchiverleft""><FONT COLOR=""Red"">[楼主]</FONT> / 用户名："&RS(4)&" </div><div id=""showarchiverright"">  发布时间："&RS(5)&" / 查看 "& Rs(7) &"</div></div><div id=""archivers"">"
			Set iGUid = team.execute("Select UserGroupID from ["& isforum &"User] Where UserName='"& Rs(4) &"' ")
			If Not iGUid.Eof Then
				If iGUid(0) = 6 Or iGUid(0) = 7 Then
					tmp = tmp &  "<font color=""red"">==该用户已被锁定==</font>"
				Else
					tmp = tmp &  ReadPowers(iUbb_Code(Replace(RS(6),"'","")))
				End If
			Else
				tmp = tmp &  ReadPowers(iUbb_Code(Replace(RS(6),"'","")))
			End If
			iGUid.Close:Set iGUid = Nothing
			tmp = tmp & "</div>" 
		End if
		ReList = Rs(1)
	Else
		Exit Function 
	End If
	RS.Close:Set RS=Nothing
	Set Rs=Server.CreateObject ("adodb.RecordSet")
	IsPage = team.execute("Select Count(*) From ["&IsForum & ReList&"] Where topicid="&tid)(0)
	SQL="SELECT Username,Content,Posttime,Lock From ["&IsForum & ReList&"] Where topicid="&tid&" Order By ID Asc"
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum=SqlQueryNum+1
		Maxpage = 50		'每页分页数
		PageNum = Abs(int(-Abs(IsPage/Maxpage)))	'页数
		Page = CheckNum(Page,1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		iRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
	If Page<2 Then
		U=1
	Else
		U=Page*Maxpage+1-Maxpage
	End If
	If Isarray(iRs) Then
		For i=0 To Ubound(iRs,2)
			U = U+1
			tmp = tmp &  " <div id=""rssshow""><div id=""showarchiverleft"">[第"&U&"楼] / 用户名："&iRs(0,i)&" </div><div id=""showarchiverright"">发布时间："&iRs(2,i)&"</div></div><div id=""archivers"">"
			Set iGUid = team.execute("Select UserGroupID from ["& isforum &"User] Where UserName='"& iRs(0,i) &"' ")
			If Not iGUid.Eof Then
				If iGUid(0) = 6 Or iGUid(0) = 7 Then
					tmp = tmp &  "<font color=""red"">==该用户已被锁定==</font>"
				Else
					tmp = tmp &  IIF(irs(3,i)=1,"==帖子已经锁定==",ReadPowers(iUbb_Code(Replace(iRs(1,i),"'",""))))
				End If
			Else
				tmp = tmp &  IIF(irs(3,i)=1,"==帖子已经锁定==",ReadPowers(iUbb_Code(Replace(iRs(1,i),"'",""))))
			End If
			iGUid.Close:Set iGUid = Nothing
			tmp = tmp & "</div>" 
		Next
	End If
	tmp = tmp &  "<div id=""rsspage"">"
	tmp = tmp &  team.PageList(PageNum,IsPage,6)
	tmp = tmp &"</div><div id=""rsspage"">[<A HREF=""../Thread.asp?tid="& TID &"""><B>查看原帖</B></A>]</div>" 
	Main = tmp
End Function

team.HtmlEnd
%>