<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim Fid,x1,x2
team.Headers("文件搜索!")
X1="<a href=Search.asp>文件搜索</A>"
X2=""
echo team.MenuTitle
Call TestUser()
Select Case Request("action")
	Case "seachfile"
		Call seachfile
	Case Else
		Call Main
End Select
Team.footer

Sub seachfile
	Call Main()
	Dim searchkey,Rs,StrSQL,SClass,AllCount,PageNum
	Dim IsWhere,nTop,Page,Shows,Maxpage,i
	Page = HRF(2,2,"Page")
	Sclass = CID(Trim(Request("searchclass")))
	searchkey = HtmlEncode(Trim(Request("searchkey")))
	team.SearcKeys = searchkey
	team.SearchClass = Sclass
	if team.Group_Browse(4) < 1 then 
		team.Error(" 你所在的组 "&team.levelname(0)&" 没有搜索的权限.")
	End If
	If Sclass = 0 Or Not IsNumeric(Sclass) Then
		team.Error "搜索参数错误。"
	End If

	If Sclass = 1 Then
		If Not IstrueName(searchkey) Then
			team.Error "错误的搜索参数"
		End If
		If searchkey &"" = "" Then
			team.Error "搜索内容不能为空"
		End If
	ElseIf Sclass = 2 Then
		If searchkey &"" = "" Then
			team.Error "搜索内容不能为空"
		End If
	ElseIf Sclass = 3 Then
	Else
		team.Error "错误的搜索参数"
	End If 
	If Page=0 Then
		Cache.Reloadtime = 1
		Cache.Name="Usearch"
		If Cache.ObjIsEmpty() Then
			Cache.value = 1
		Else
			Cache.value = Cache.value +1
		End if
		If Cache.value > team.Forum_setting(52) Then
			team.Error "服务器同时搜索量超过系统设置，请稍后再运行搜索程序。"
		End If
		Call CheckpostTime
		Session(CacheName &"searchtime") = Now()
	End If
	If Sclass = 1 Then
		IsWhere = " UserName like '%" & searchkey & "%' " 
	ElseIf Sclass = 2 Then
		If IsSqlDataBase = 1 Then
			IsWhere = " Topic like '%" & searchkey & "%' "
		Else
			IsWhere = " InStr(1,LCase(Topic),LCase('"&searchkey&"'),0)<>0 "
		End if
	ElseIf Sclass = 3 and team.Group_Browse(4) = 2 Then
		If IsSqlDataBase = 1 Then
			IsWhere = " Content like '%" & searchkey & "%' "
		Else
			IsWhere = " InStr(1,LCase(Content),LCase('"&searchkey&"'),0)<>0 "
		End If
	ElseIf Sclass = 4 Then
		IsWhere = " Goodtopic = 1 "
	End If
	If Sclass = 3 Then
		nTop = " Top 10 "
	Else
		nTop = ""
	End If
	AllCount = team.Execute("Select Count(ID) from ["&IsForum&"forum] where deltopic=0 and Locktopic=0 and "&IsWhere&" ")(0)
	If AllCount = "" Or Not IsNumEric(AllCount) Then 
		AllCount = 0
	End if
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	UpdateUserpostExc()
	StrSQL = "Select "&nTop&" ID,forumid,Topic,Username,Views,Icon,Replies,lasttime,Goodtopic,Createpoll,Creatdiary,Creatactivity,Rewardprice,Readperm,Rewardpricetype from ["&IsForum&"forum] where deltopic=0 and Locktopic=0 and "&IsWhere&" Order By Lasttime DESC"
	Rs.Open StrSQL,Conn,1,1,&H0001
	Response.Write "<table cellspacing=1 cellpadding=10 width=98% align=center border=0><tr class=a3><td colspan=2 align=center>本次搜索共找到 <Font color=red>"&AllCount&"</Font> 条相关帖子记录</td></tr>"
	If Rs.Eof And Rs.Bof Then
		Echo " <tr class=a4><td colspan=2 align=center> 对不起，本站没有找到您要查询的内容，您想到百度去搜索关于 <B>["&searchkey&"]</B> 的信息么？ <a href=""http://www.baidu.com/baidu?tn=team5&word="&searchkey&""" target=""_blank"">【立即搜索网络】</a> </td></tr></table>"
	Else
		Maxpage = 20
		PageNum = Abs(int(-Abs(AllCount/Maxpage)))	'页数
		Page = CheckNum(Page,1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	Dim Un,tmp,ExtCredits,bbsname,Chcheid,j
	ExtCredits = Split(team.Club_Class(21),"|")
	Chcheid = team.BoardList
	For i=0 To Ubound(shows,2)
		If Request("Page")<2 Then
			Un=i+1
		Else
			Un=Int(Request("Page"))*Maxpage-Maxpage+i+1
		End If
		For j=0 to Ubound(Chcheid,2)
			If Cid(Shows(1,i)) = Cid(Chcheid(0,j)) Then
				BBsName = Chcheid(1,j)
			End if
		Next
		Echo "<tr class=a4><td height=50>"&Un&"."&iif(Shows(8,i)=1,"<img src="""&team.styleurl&"/f_good.gif"" border=""0"" align=""absmiddle"" alt=""精华"" >","")&" "
		Echo "<a Href=Thread.asp?tid="&Shows(0,i)&" target=""_blank""> "&GetColor(Shows(2,i),searchkey)&" "&iif(Cid(Shows(13,i))>0,"- [<b>阅读权限</b> "&Shows(13,i)&"]","")&" "&iif(Cid(Shows(10,i))>0,"- [<b>用户文集</b>]","")&" "&iif(Cid(Shows(11,i))>0,"- [<b>活动召集</b>]","")&" "&IIf(Cid(Shows(14,i))=0,iif(Cid(Shows(14,i))>0,"- [<b>悬赏 </b> "&IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  "  "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" "&Shows(14,i)&" "," 本积分未启用 ")&"]",""),"[已解决]")&" "&iif(DateDiff("d",Shows(7,i),date())=0,"  <img src="""&team.styleurl&"/new.gif"" border=""0"" align=""absmiddle"">","")&" </a> "
		Echo " <BR /> <font color=""green"">作者："&GetColor(Shows(3,i),searchkey)&" 浏览："&Shows(4,i)&"  回复："&Shows(6,i)&"  →   "
		Echo " <a Href=Thread.asp?tid="&shows(0,i)&" target=""_blank"">"&BBsName&"</a> "&Shows(7,i)&" </Font> "
		Echo " <hr style=""border-top:1px #B3B3B3 dashed;border-bottom:0px;height:0px;width:98%;""></hr></td> </tr>"
	Next
	Echo "</table><BR>"
	Echo "<div id=""pagediv"">"& team.PageList(PageNum,AllCount,6) &"</div><BR><div id=""rsspage"">"&Iif(team.Forum_setting(42)=1,team.BoardJump,"")&"</div>"
	Echo tmp
End Sub

Sub UpdateUserpostExc()
	'用户积分部分
	Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u
	Dim UserPostID,My_ExtSort
	If Not team.UserLoginED Then  Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	For U=0 to Ubound(ExtCredits)
		ExtSort=Split(ExtCredits(U),",")
		MustSort=Split(MustOpen(U),",")
		If ExtSort(3)=1 Then
			If U = 0 Then
				UExt = UExt &"Extcredits0=Extcredits0-"&MustSort(6)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"-"&MustSort(6)&""
			End If
			If (team.User_SysTem(14+U)-MustSort(6))-MustSort(8)<0 Then
				team.Error "您的"&ExtSort(0)&" ["& team.User_SysTem(14+U) - MustSort(4) &"] 低于积分策略下限值 ["& MustSort(8)&"] ，所以无法进行此操作。"
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where ID = "& team.TK_UserID)
End Sub

Sub Main
	Echo "<script type=""text/javascript""> "
	Echo "function validate(theform) { "
	Echo " var searchkey = theform.searchkey.value; "
	Echo " if ( searchkey ==  '') { "
    Echo "		alert('请完成内容栏。'); "
    Echo "		return false; "
	Echo "	}	"
	Echo " this.document.myform.submit.disabled = true; "
	Echo "}"
	Echo "</script>"
	Echo "<form name=""myform"" method=""post"" action=""?action=seachfile"" onSubmit=""return validate(this)"">"
	Echo "<table cellSpacing=""0"" cellPadding=""10"" width=""98%"" border=""0"" align=""center"">"
	Echo "<tr class=""a3""><td width=""30%""></td>"
	Echo "	<td class=""bold""><IMG SRC="""&team.styleurl&"/bin.GIF"" BORDER=""0"" ALT=""搜索"" align=""absmiddle"">  TEAM智能搜索系统 , 请输入要搜索的关键字 </td>"
	Echo "</tr>"
	Echo "<tr class=""a4"">"
	Echo "	<td></td><td>"
	Echo "	<input value="""" name=""searchkey"" size=""50"" onmouseover=""this.focus()"" maxlength=""50"" type=""text"" title=""智能搜索开始 -- >"" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.select();"" class=""colorblur""> "
	Echo "	<input type=""submit"" name=""submit"" value=""开始搜索"">"
	Echo "	<BR><BR> "
	Echo "	<input type=""radio"" name=""searchclass"" value=""1"" class=""radio""> 按用户名搜索 "
	Echo "	<input type=""radio"" name=""searchclass"" value=""2"" class=""radio"" checked> 按标题搜索 "
	If team.Group_Browse(4) = 2 Then  
		Echo "	<input type=""radio"" name=""searchclass"" value=""3"" class=""radio""> 按内容搜索 "
	End if
	Echo "	</td>"
	Echo "</tr>"
	Echo "</table>"
	Echo "</form> "
End Sub

Sub CheckpostTime()
	If CID(team.Forum_setting(51))<=0 Then
		Exit Sub
	Else
		If IsDate(Session(CacheName &"searchtime")) Then
			If DateDiff("s",Session(CacheName &"searchtime"),Now())< CID(team.Forum_setting(51)) Then
				team.Error "为防止有人恶意消耗系统资源，论坛限制单个用户两次搜索间隔必须大于"&team.Forum_setting(51)&"秒，您还需要等待 "& CLng(team.Forum_setting(51))-DateDiff("s",Session(CacheName &"searchtime"),Now()) &" 秒。"
			End If
		End If
	End if
End Sub
%>