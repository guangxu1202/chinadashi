<!-- #include file="CONN.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim Fid,x1,x2,ShowClass
Fid = HRF(2,2,"fid")
Set ShowClass = New ShowMyThreads
ShowClass.ShowBoard()

Class ShowMyThreads
	Public Boards,Board_Setting,Postlist,ii,i,Forumid,page
	Private CountUs,Torder,Search,TimeLimit,topicmode,TWhere,IsPage,irs,UserList,tmp
	Private Sub Class_Initialize()
		Dim Rs
		Cache.Name = "ForumsBoards_"&Fid
		Cache.Reloadtime = Cid(team.Forum_setting(44))
		If Not Cache.ObjIsEmpty() Then
			Boards = Cache.Value
		Else
			Set Rs=team.Execute("Select ID,Followid,bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Board_Key,Board_URL,Board_Code,toltopic,tolrestore,lookperm,postperm,downperm,upperm From ["&IsForum&"Bbsconfig] Where ID = "& Fid)
			If Rs.Eof And Rs.bof Then 
				Team.Error "你查询的版面ID错误。"
				Exit Sub
			Else
				Cache.Value = Rs.GetRows(-1)
			End If
			RS.Close:Set RS=Nothing
			Boards = Cache.Value
		End If
		If isarray(Boards) Then
			Board_Setting = Split(Boards(3,0),"$$$")
		End If
		If Boards(1,0) = 0 Then
			Response.Redirect "Default.asp?rootid="&Fid
		End If
		If CID(Board_Setting(8)) = 0 Then
			If Not (Boards(13,0) = ",") Then
				If Instr(Boards(13,0),",") > 0 Then 
					If Not GetUserPower Then team.Error "您没有查看本版的权限"
				End If
			End If
		End if
		team.ChooseName = Board_Setting(0)
		team.Headers(Boards(2,0))
		team.OnlinActions(Fid&",查看帖子列表,"&Boards(2,0))
		If Boards(5,0)<>"" And Not (team.IsMaster Or team.SuperMaster) Then
			If CID(Request.Cookies("Class")("LoginKey"& fid)) = 0 Then
				Response.Redirect "PassKey.asp?fid="&fid&""
			End if
		End If
	End Sub

	Private Function GetUserPower()
		GetUserPower = False
		Dim B_Lookperm,m
		B_Lookperm = Split(Boards(13,0),",")
		If Isarray(B_Lookperm) Then
			For m = 0 to Ubound(B_Lookperm)-1
				If Cid(B_Lookperm(m)) = Int(team.UserGroupID) Then GetUserPower = True
			Next 
		End  If
	End Function

	Public Sub ShowBoard()
		Dim Tmpid,j,Chcheids,CheckMaster,Moder,u
		Dim ForumNews,Reimage,Reimage2
		Dim SQL,Rs,ExtCredits,topicmode
		Chcheids = team.BoardList
		ForumNews = team.Affiche
		CheckMaster = team.GroupManages
		tmp = Team.PostHtml(0)
		x1= IIF(CID(team.Forum_setting(65))=1,"<a href=""Forums-"&FID&".html"">"& Sign_Code(Boards(2,0),1) &" </A>","<a href=""Forums.asp?fid="&FID&""">"& Sign_Code(Boards(2,0),1) &" </A>")
		For j=0 to Ubound(Chcheids,2)
			If Cid(Boards(1,0))=Cid(Chcheids(0,j)) and Chcheids(3,j)>0 Then
				x2= IIF(CID(team.Forum_setting(65))=1,"<a href=""Forums-"&Chcheids(0,j)&".html"">"& Chcheids(1,j) &" </A>","<a href=""Forums.asp?fid="&Chcheids(0,j)&""">"& Chcheids(1,j) &" </A>")
			End if
		Next
		ExtCredits = Split(team.Club_Class(21),"|")
		tmp=Replace(tmp,"{$wensurl}",team.MenuTitle)
		tmp=Replace(tmp,"{$minicoard}",ForumList(fID))
		tmp=Replace(tmp,"{$notshow}",Iif(Boards(8,0)<>"","","display:none"))
		tmp=Replace(tmp,"{$minilogo}",Ubb_Code(Boards(8,0))&"")
		If Isarray(CheckMaster) Then
			If team.Forum_setting(26)=0 Then Moder = Moder & "<select size=""1""><option> -->>版主列表</option>"
			For u=0 to Ubound(CheckMaster,2)	
				If CheckMaster(2,u) = Cid(Boards(0,0)) Then
					If team.Forum_setting(26)=1 Then
						If Moder = "" Then
							Moder = Moder & IIF(CID(team.Forum_setting(65))=1," <a href=Profile-"&CheckMaster(1,u)&".html>"&CheckMaster(1,u)&"</a> "," <a href=Profile.asp?username="&CheckMaster(1,u)&">"&CheckMaster(1,u)&"</a> ")
						Else
							Moder = Moder & IIF(CID(team.Forum_setting(65))=1,", <a href=Profile-"&CheckMaster(1,u)&".html>"&CheckMaster(1,u)&"</a> ",", <a href=Profile.asp?username="&CheckMaster(1,u)&">"&CheckMaster(1,u)&"</a> ")
						End If
					Else
						Moder = Moder &"<option> "&CheckMaster(1,u)&" </option>"
					End if
				End If
			Next
			If team.Forum_setting(26)=0 Then Moder = Moder&  "</select>"
		End if
		tmp=Replace(tmp,"{$moderated}",iif(Moder &""="","本版暂无版主",Moder))
		Dim MyAnnus,MyNames
		If IsArray(ForumNews) Then 
			MyAnnus = "<A href=""Affiche.asp#"&ForumNews(0,0)&""" target=""_blank"">"&ForumNews(1,0)&"</a>"
			MyNames = ForumNews(3,0)
		Else
			MyAnnus = ForumNews
			MyNames = ""
		End if
		tmp=Replace(tmp,"{$news}",MyAnnus)
		tmp=Replace(tmp,"{$newname}",MyNames)
		Dim ascdesc,RsP
		If HRF(2,1,"ascdesc") = "asc" Then
			AscDesc = "asc"
		Else
			AscDesc = "desc"
		End If
		Select Case HRF(2,1,"orderby")
			Case "lastpost"
				Torder="Lasttime"
			Case "dateline"
				Torder="Posttime"
			Case "replies"
				Torder="Replies"
			Case "views"
				Torder="Views"
			Case Else
				Torder="Lasttime"
		End Select
		TWhere= ""
		If HRF(2,2,"isgod") = 1 Then
			TWhere = " goodtopic=1 "
			CountUs=1
		End if
		If HRF(2,2,"filter")>=86400 Then
			If IsSqlDataBase=1 Then
				TWhere= " Datediff(Mi, Lasttime, " & SqlNowString & ") < "&HRF(2,2,"filter")/60&" "
			Else
				TWhere= " Datediff('s',Lasttime, " & SqlNowString & " ) < "&HRF(2,2,"filter")&" "
			End If
			CountUs=1
		End If
		If Request("topicmode")<>""  Then
			TWhere= " PostClass="&HRF(2,2,"topicmode")&" "
			CountUs=1
		End if
		If CountUs=1 Then				'记录总数
			Set RsP = Team.Execute("Select Count(ID) From ["&Isforum&"Forum] Where Auditing=0 and deltopic=0 and "&TWhere&" and Forumid="&Int(Fid))
			If Not(RsP.Eof or Rsp.Bof) Then
				IsPage=RsP(0)
			End If
			TWhere = "and " & TWhere
			TWhere = ReplaceStr(TWhere," and and "," and ")
		Else
			IsPage = Boards(11,0)
		End If
		Dim Maxpage,PageNum
		SQL="Select ID,Topic,Username,Views,Icon,Replies,Color,PostClass,Toptopic,Locktopic,CloseTopic,Goodtopic,LastText,Lasttime,Createpoll,Creatdiary,Creatactivity,Rewardprice,Readperm,Rewardpricetype,IsNoName From ["&IsForum&"forum] Where deltopic=0 and Auditing=0 and (Toptopic=2 or Forumid="&Int(Fid)&") "&TWhere&" Order By Toptopic "&AscDesc&","&Torder&" "&AscDesc&""
		Set Rs = Server.CreateObject ("Adodb.RecordSet")
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,Conn,1,1,&H0001
		If Not (Rs.Eof and Rs.Bof) Then 
			SqlQueryNum=SqlQueryNum+1
			Maxpage = Cid(team.Forum_setting(19))		'每页分页数
			PageNum = Abs(int(-Abs(IsPage/Maxpage)))	'页数
			Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'当前页
			Rs.AbsolutePosition=(Page-1)*Maxpage+1
			iRs=Rs.GetRows(Maxpage)
		End if
		RS.Close:Set Rs=Nothing
		ii=0
		If Not Isarray(iRs) Then
			tmp=Replace(tmp,"{$special}","")
		Else
			For i=0 To Ubound(iRs,2)
				tmp = tmp & Team.PostHtml(1)
				tmp=Replace(tmp,"{$ID}",IIF(CID(team.Forum_setting(65))=1,"thread-"& iRs(0,i) &".html","thread.asp?tid=" & iRs(0,i)))
				tmp=Replace(tmp,"{$ismasters}",iif(team.ManageUser,"<input type=""checkbox"" name=""ismanage"" value="&iRs(0,i)&" class=""radio"">",""))
				tmp=Replace(tmp,"{$topic}",Cutstr(Htmlencode(iRs(1,i)),int(team.Forum_setting(88))) & iif(CID(iRs(20,i))=1 And team.ManageUser,"  <img src="""&team.styleurl&"/t1.gif"" alt=""匿名帖子"" align=""absmiddle"">",""))
				tmp=Replace(tmp,"{$username}",IIF(CID(iRs(20,i))=1 And Not team.ManageUser,"匿名用户",iRs(2,i)))
				tmp=Replace(tmp,"{$Views}",iRs(3,i))
				Reimage2 = ""
				If Trim(iRs(2,i)) = Trim(TK_UserName) Then Reimage2 = "<img src="""&team.styleurl&"/my.gif"">"
				If Cid(iRs(4,i))>0 Then Reimage2 = "<img src=""images/brow/icon"&iRs(4,i)&".gif"">"
				tmp=Replace(tmp,"{$reimage2}",Reimage2)
				tmp=Replace(tmp,"{$reimage3}",Cid(iRs(5,i)))
				tmp=Replace(tmp,"{$mycolor}",SetColors(iRs(6,i)))
				tmp=Replace(tmp,"{$specialshow}",IIf(CID(Board_Setting(15))=1 and CID(Board_Setting(17))=1,"","display:None"))
				Dim Special,utmp,etmp,dtmp,tips
				tmp=Replace(tmp,"{$forpower}","{$tips_1}{$tips_2}{$tips_3}")
				tmp=Replace(tmp,"{$tips_1}",iif(Cid(iRs(18,i))>0,"- [<b>阅读权限</b> "&iRs(18,i)&"]",""))
				tmp=Replace(tmp,"{$tips_2}",iif(Cid(iRs(16,i))>0,"- [<b>活动召集</b>]",""))
				tmp=Replace(tmp,"{$tips_3}",IIf(Cid(iRs(19,i))=0,iif(Cid(iRs(17,i))>0,"- [<b>悬赏 </b> "&IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  "  "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" "&iRs(17,i)&" "," 本积分未启用 ")&"]",""),"[已解决]"))
				Special = ""
				If Instr(Board_Setting(19),Chr(13)&Chr(10))>0 Then
					utmp = Split(Board_Setting(19),Chr(13)&Chr(10))
					For U=0 To Ubound(utmp)
						Special = Special &" <td class=""a4""> <a href=""Forums.asp?fid="&Fid&"&topicmode="&U&""">"& utmp(u) &"</a> </td> "
					Next
				Else
					Special = "<td class=""a4""> <a href=""Forums.asp?fid="&Fid&"&topicmode=0"">"& Board_Setting(19) &"</a> </td> "
				End if
				tmp=Replace(tmp,"{$special}",IIf(CID(Board_Setting(15))=1,"<td class=""a1"">主题分类</td>"& Special &"",""))
				etmp = ""
				If Cid(iRs(7,i))<>"999" and CID(Board_Setting(18)) = 1 Then
					If Instr(Board_Setting(19),Chr(13)&Chr(10))>0 Then
						utmp = Split(Board_Setting(19),Chr(13)&Chr(10))
						If CID(iRs(7,i)) <= UBound(utmp) Then etmp = utmp(iRs(7,i))
						dtmp = iRs(7,i)
					Else
						etmp = Board_Setting(19)
						dtmp = 0
					End if
					If Etmp<>"" Then etmp = " <a href=""Forums.asp?fid="&Fid&"&topicmode="& dtmp &""">["& etmp &"]</a> "
				End if
				tmp=Replace(tmp,"{$posttopic}",iif(CID(Board_Setting(18))=1,etmp,""))
				Reimage = ""
				if iRs(5,i) = 0 then Reimage = "f_norm.gif"
				if iRs(5,i) > 0 then Reimage="f_new.gif"
				if iRs(5,i) > Cid(team.Forum_setting(22)) or iRs(3,i)>150 then reimage="f_hot.gif"
				if iRs(14,i)<>empty then Reimage="f_poll.gif"
				if iRs(10,i) = 1 then Reimage="f_locked.gif"
				if iRs(9,i) = 1 then Reimage="lock.gif"
				if iRs(8,i) = 1 then Reimage="ztop.gif"
				if iRs(8,i) = 2 then Reimage="top.gif"
				tmp=Replace(tmp,"{$reimage}","<img src="""&team.styleurl&"/"&Reimage&""" border=""0"" align=""absmiddle"">")
				tmp=Replace(tmp,"{$ismanage}",iif(iRs(11,i)=1,"<img src="""&team.styleurl&"/f_good.gif"" border=""0"" align=""absmiddle"" alt=""精华"" >",""))
				tmp=Replace(tmp,"{$lasttime}",iRs(13,i))
				tmp=Replace(tmp,"{$lastname}",IIF(Split(iRs(12,i),"$@$")(0)=" － ",iRs(2,i),Split(iRs(12,i),"$@$")(0)))
				tmp=Replace(tmp,"{$lasttop}",Split(iRs(12,i),"$@$")(1))
				tmp=Replace(tmp,"{$newimg}","{$newimg1}{$newimg2}")
				tmp=Replace(tmp,"{$newimg1}",iif(DateDiff("d",iRs(13,i),date())=0,"  <img src="""&team.styleurl&"/new.gif"" border=""0"" align=""absmiddle"">",""))
				Dim tpage,mPage,h
				tpage = ""
				mPage = Abs(Int(-Abs(Cid(iRs(5,i))/CID(team.Forum_setting(20)))))
				if mPage > 6 Then
					For h = 2 To 5
						tpage  = tpage & IIF(CID(team.Forum_setting(65))=1," <b><a href=Thread-"&iRs(0,i)&"-"&H&".html>"&h&"</a></b> "," <b><a href=Thread.asp?tid="&iRs(0,i)&"&Page="&H&">"&h&"</a></b> ")
					Next
					tpage  = tpage & "..."
					for h = mPage-1 to mPage
						tpage  = tpage & IIF(CID(team.Forum_setting(65))=1," <b><a href=Thread-"&iRs(0,i)&"-"&H&".html>"&h&"</a></b> "," <b><a href=Thread.asp?tid="&iRs(0,i)&"&Page="&H&">"&h&"</a></b> ")
					next
				Else
					For h=2 To mPage
						tpage  = tpage & IIF(CID(team.Forum_setting(65))=1," <b><a href=Thread-"&iRs(0,i)&"-"&H&".html>"&h&"</a></b> "," <b><a href=Thread.asp?tid="&iRs(0,i)&"&Page="&H&">"&h&"</a></b> ")
					Next	
				End If
				tmp=Replace(tmp,"{$newimg2}",IIF(Cid(iRs(5,i)) > CID(team.Forum_setting(20))," [<img src="""&team.styleurl&"/multipage.gif"" align=""absmiddle"">" &tpage &"] ",""))
				tmp=Replace(tmp,"{$uswindows}",iif(team.Forum_setting(43)=1,"target=""_blank""",""))
				If Cid(iRs(8,i))>0 Then ii=ii+1
				tmp=Replace(tmp,"{$topstitle}",iif(Page<2,iif(i-ii=0,"<tr class=""a3""><td>&nbsp;</td><td colspan=""5""><span class=""bold"">论坛主题</span></td></tr>",""),""))
			Next
		End If
		tmp = tmp & Team.PostHtml (2)
		tmp = Replace(tmp,"{$getnesboard}",Iif(team.Forum_setting(42)=1,team.BoardJump,""))
		tmp = Replace(tmp,"{$actionmanage}",Iif(Not team.ManageUser,"",Team.PostHtml (3)))
		'If team.ManageUser Then tmp = tmp & Team.PostHtml (7)
		Dim Isshowset
		If team.Forum_setting(39)=1 or team.Forum_setting(39)=3 Then
			if Request("showlines")="yes" or Request("showlines")="" Then
				Isshowset = "<a href=""Forums.asp?fid="&fid&"&showlines=no#online""><img src="""&team.Styleurl&"/collapsed_no.gif"" align=""right"" border=""0"" alt=""点击关闭在线状况"" /></a>"
			Else
				Isshowset = "<a href=""Forums.asp?fid="&fid&"&showlines=yes#online""><img src="""&team.Styleurl&"/collapsed_yes.gif"" align=""right"" border=""0"" alt=""点击查看在线状况"" /></a>"
			End if
		Else
			if Request("showlines")="yes" Then
				Isshowset = "<a href=""Forums.asp?fid="&fid&"&showlines=no#online""><img src="""&team.Styleurl&"/collapsed_no.gif"" align=""right"" border=""0"" alt=""点击关闭在线状况"" /></a>"
			Else
				Isshowset = "<a href=""Forums.asp?fid="&fid&"&showlines=yes#online""><img src="""&team.Styleurl&"/collapsed_yes.gif"" align=""right"" border=""0"" alt=""点击查看在线状况"" /></a>"
			End if
		End if
		tmp = Replace(tmp,"{$showimg}",Isshowset)
		tmp = Replace(tmp,"{$listonlieuser}",iif(Request("showlines")="yes" or team.Forum_setting(39)=2 or team.Forum_setting(39)=3,team.Showlines(fid),""))
		tmp=Replace(tmp,"{$onlinemany}",Team.Onlinemany)
		Dim OName,RName
		Cache.Name = "Forumidonline"& fid
		OName = Cache.Value
		Cache.Name = "Regforumidonline"& fid
		RName = Cache.Value
		tmp=Replace(tmp,"{$forumidonline}",OName)
		tmp=Replace(tmp,"{$regforumidonline}",RName)
		tmp=Replace(tmp,"{$nouseronline}",OName - RName)
		tmp = Replace(tmp,"{$onlineshow}",Iif(Request("showlines")="yes" or team.Forum_setting(39)=2 or team.Forum_setting(39)=3,"","display:None"))
		Dim Url,itrue
		If Request("filter")="" And Request("orderby")="" And Request("ascdesc")="" Then
			itrue = True
		End if
		'PageList 每页分页数,总记录数,当前页,总页数,当前Url
		tmp = Replace(tmp,"{$pagelists}",team.PageList(PageNum,IsPage,6)) 
		tmp = Replace(tmp,"{$TotalPage}",PageNum)
		tmp = Replace(tmp,"{$allPage}",IsPage)
		tmp = Replace(tmp,"{$forumid}",Fid)
		tmp = Replace(tmp,"{$looknows}",SetNowLooks)
		tmp = Replace(tmp,"{$maxsml}",Cid(team.Forum_setting(87)))
		Echo tmp
		Call team.footer
	End Sub

	Private Function SetNowLooks()
		Dim t,MyBoard,u,w,s
		MyBoard = Request.Cookies("Class")("Board")
		If InStr(MyBoard & "$$", Boards(2,0) & "$$") <= 0 Then
			Response.Cookies("Class")("Board")  = MyBoard & Fid & "@@" & Boards(2,0) & "$$"
		End If
		If team.Forum_setting(24) =0 Then
			Exit Function
		Else
			t = "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}""><option value="""" selected>最近浏览的论坛</option>"
			s = Split(MyBoard,"$$")
			For u = 0 To UBound(s) - 1
				If u >= CID(team.Forum_setting(24)) Then Exit For
				W = Split(s(u),"@@")
				t = t & IIF(CID(team.Forum_setting(65))=1,"<option value=""Forums-"&W(0)&".html"">"&W(1)&"</option>","<option value=""Forums.asp?fid="&W(0)&""">"&W(1)&"</option>")
			Next
			t = t & "</Select>"
			SetNowLooks = t
		End if
	End Function

	Private Function ForumList(B)
		Dim ShowBbs,i,Rs,Moderuser,tmp
		Showbbs = team.BoardList()
		If Not IsArray(Showbbs) Then
			Exit Function
		End if
		For i=0 to Ubound(Showbbs,2)
			If ShowBBs(3,i) = B Then
				tmp = "<div Class=""a4""><table cellspacing=""1"" cellpadding=""3"" width=""98%"" align=""center"" class=""a2""><tr class=""a6"" align=""center""><td width=""5%"">&nbsp;</td><td width=""45%"">论坛</td><td width=""5%"">主题</td><td width=""5%"">回帖数</td><td width=""5%"">今日</td><td width=""25%"">最后发表</td><td width=""15%"">版主</td></tr>"
				tmp = tmp & team.ForumList_tips(b)
				tmp = tmp & "</table></div><BR>"
			End if
		Next
		ForumList = tmp
	End Function

End Class
team.htmlend
%>