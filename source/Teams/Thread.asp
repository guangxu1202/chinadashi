<!-- #include file="CONN.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim tID,Page,MyThread,x1,x2,fID
tID = HRF(2,2,"tid")
Page = HRF(2,2,"Page")
Set MyThread = New PostClass
MyThread.ShowThread
team.footer

Class PostClass
	Private Board_Setting,Boards,ExtCredits,UClass,UserOnlineinfos,tmp,PageNum,GetNoName,Maxpage
	Private Sub Class_Initialize()
		Dim Rs,Title
		Set Rs=team.Execute("Select forumid,Topic From ["&IsForum&"forum] Where  ID = "& tID)
		If Rs.Eof And Rs.Bof Then
			Team.Error "你查询的版面ID错误。"
		Else
			fID = Rs(0)
			Title = HTMLEncode(Rs(1))
		End If
		Rs.Close:Set Rs=Nothing
		Cache.Name = "ThreadBoards_"&fid
		Cache.Reloadtime = Cid(team.Forum_setting(44))
		If Cache.ObjIsEmpty() Then
			Set Rs=team.Execute("Select bbsname,Board_Setting,Board_Code,lookperm,Pass From ["&IsForum&"Bbsconfig] Where ID = "& fID)
			If Rs.Eof And Rs.Bof Then 
				Team.Error "你查询的版面ID错误。"
				Exit Sub
				Response.End
			Else
				Boards = Rs.GetRows(-1)
				Cache.Value = Boards
			End If
			RS.Close:Set RS=Nothing
		End If
		Boards = Cache.Value
		If isarray(Boards) Then
			Board_Setting = Split(Boards(1,0),"$$$")
		Else
			team.error " 版块参数错误 。"
		End If
		ExtCredits = Split(team.Club_Class(21),"|")
		UserOnlineinfos = team.UserOnlineinfos
		team.ChooseName = Board_Setting(0)
		team.Headers(Title)
		If CID(team.Forum_setting(110)) = 2 Then
			TestUser()
		End if
		If Not (Boards(3,0) = ",") Then
			If Instr(Boards(3,0),",") > 0 Then 
				If Not GetUserPower Then team.Error "您没有查看本版的权限"
			End If
		End If
		team.OnlinActions(fID&",查看帖子,"&tiTle)
		If Boards(4,0)<>"" And Not (team.IsMaster Or team.SuperMaster) Then
			If CID(Request.Cookies("Class")("LoginKey"& fid)) = 0 Then
				Response.Redirect "PassKey.asp?fid="&fid&""
			End if
		End If
		team.IsUbb = CID(Board_Setting(6))
	End Sub

	Private Function GetUserPower()
		GetUserPower = False
		Dim B_Lookperm,m
		B_Lookperm = Split(Boards(3,0),",")
		If Isarray(B_Lookperm) Then
			For m = 0 to Ubound(B_Lookperm)-1
				If Cid(B_Lookperm(m)) = Int(team.UserGroupID) Then GetUserPower = True
			Next 
		End  If
	End Function

	Public Sub ShowThread()
		Dim RS,Gs,i,emp
		GetNoName = 0
		SQL=" T.ID,T.Topic,T.Username,T.Views,T.Icon,T.Replies,T.Color,T.PostClass,T.Toptopic,T.Locktopic,T.CloseTopic,T.Goodtopic,T.Content,T.Posttime,T.Createpoll,T.Creatdiary,T.Creatactivity,T.Rewardprice,T.Readperm,T.ReList,T.Rewardpricetype,T.Tags,U.Levelname,U.Posttopic,U.Postrevert,U.Goodtopic,U.Regtime,U.Landtime,U.Birthday,U.UserSex,U.Sign,U.UserInfo,U.Honor,U.Userface,U.ID,U.Degree,U.Postblog,U.UserCity,U.UserUp,U.Extcredits0,U.Extcredits1,U.Extcredits2,U.Extcredits3,U.Extcredits4,U.Extcredits5,U.Extcredits6,U.Extcredits7,U.UserGroupID,U.Medals,T.Createreward,T.IsNoName,T.Postip"
		Set Rs = team.Execute("Select TOP 1 "&SQL&" From ["&IsForum&"Forum] T Inner Join ["&IsForum&"User] U On U.UserName=T.UserName Where T.Auditing=0 and T.deltopic=0 and T.ID="& tID )
		If Rs.EOF And Rs.BOF Then
			Set Rs = Nothing
			Set Rs = team.Execute("Select TOP 1  ID,Topic,Username,Views,Icon,Replies,Color,PostClass,Toptopic,Locktopic,CloseTopic,Goodtopic,Content,Posttime,Createpoll,Creatdiary,Creatactivity,Rewardprice,Readperm,ReList,Rewardpricetype,Tags,IsNoName,Postip From ["&IsForum&"Forum] Where Auditing=0 and deltopic=0 and ID="& tID )
			If Not (Rs.EOF And Rs.BOF) Then
				UClass = Rs.GetRows(1)
				GetNoName = 1
			End if
		Else
			UClass = Rs.GetRows(1)
		End If
		Rs.Close:Set Rs=Nothing
		If Not IsArray(UClass) Then
			team.Error "此帖子已经被删除或锁定。"
		End If
		X1= HTMLEncode(UClass(1,0))
		X2= Sign_Code(Boards(0,0),1)
		If Page<2 Then
			tmp = Team.PostHtml (4) & Team.PostHtml (5)
		Else
			tmp = Team.PostHtml (4)
		End If
		tmp = Replace(tmp,"{$wensurl}",team.MenuTitle)
		tmp = Replace(tmp,"{$reid}","Frist")
		tmp = Replace(tmp,"{$showadv}",IIF(Cid(UClass(11,0))=1,"<div class=""refine""></div>","")) '预留广告位
		if Not team.ManageUser and Cid(UClass(9,0)) = 1 Then 
			team.error "此主贴已经被管理员锁定!"
		End If
		If Not TestUserRead Then 
			Team.Error "你没有阅读此帖的权限。"
		End If
		If UClass(21,0)<>"" Then
			Dim TagsUrl,SQL,Tags
			TagsUrl = ""
			If Instr(UClass(21,0),"|")>=0 Then
				Tags = Split(UClass(21,0),"|")
				for i = 0 to Ubound(Tags)
					TagsUrl = TagsUrl & "<A href=""Search.asp?action=seachfile&searchclass=2&searchkey="&Tags(i)&""" target=""_blank""> "& Tags(i) &" </a>"
				Next
			Else
				TagsUrl = TagsUrl & "<A href=""Search.asp?action=seachfile&searchclass=2&searchkey="&UClass(21,0)&""" target=""_blank""> "& UClass(21,0) &" </a>"
			End if
		End If
		tmp = Replace(tmp,"{$istags}",IIf(UClass(21,0)<>"","<BR><BR>关键字："& TagsUrl &"",""))
		If Cid(UClass(14,0))=1 Then
			Call PollAction()	'投票主题
		End If
		If Cid(UClass(16,0))=1 Then
			Call ActionLive()	'活动
		End If
		tmp = Replace(tmp,"{$postactionsinfo}","<script type=""text/javascript"" language=""javascript"" src=""Images/Plus/lightbox.js""></script><link rel=""stylesheet"" href=""Images/Plus/lightbox.css"" type=""text/css"" media=""screen""/>") '贴间广告
		tmp = Replace(tmp,"{$views}",UClass(3,0))
		'tmp = Replace(tmp,"{$issetcode}",iif(team.Forum_setting(48)=2,"1","0"))
		tmp = Replace(tmp,"{$forumid}",fid)
		tmp = Replace(tmp,"{$isrept}","TOPS")
		tmp = Replace(tmp,"{$isreid}",UClass(0,0))
		tmp = Replace(tmp,"{$setopic}",SerVer.UrlEncode(UClass(1,0)))
		tmp = Replace(tmp,"{$lasttime}",UClass(13,0))
		tmp = Replace(tmp,"{$notshow}","")
		tmp = Replace(tmp,"{$ismanage}",IIf(team.ManageUser,"<input type=""checkbox"" name=""fismanage"" value="&UClass(0,0)&" class=""radio"">",""))
		tmp = Replace(tmp,"{$topic}","<span style="""& SetColors(UClass(6,0)) &""">"& HTMLEncode(UClass(1,0)) & "</span>")
		'没有回复处理
		'T.ID=0,T.Topic=1,T.Username=2,T.Views=3,T.Icon=4,T.Replies=5,T.Color=6,T.PostClass=7,T.Toptopic=8,T.Locktopic=9,T.CloseTopic=10,T.Goodtopic=11,T.Content=12,T.Lasttime=13,T.Createpoll=14,T.Creatdiary=15,T.Creatactivity=16,T.Rewardprice=17,T.Readperm=18,T.ReList=19,T.Rewardpricetype=20,T.Tags=21,U.Levelname=22,U.Posttopic=23,U.Postrevert=24,U.Goodtopic=25,U.Regtime=26,U.Landtime=27,U.Birthday=28,U.serSex=29,U.Sign=30,U.UserInfo=31,U.Honor=32,U.Userface=33,U.ID=34,U.Degree=35,U.Postblog=36,U.UserCity=37,U.UserUp=38,U.Extcredits0=39,U.Extcredits1=40,U.Extcredits2=41,U.Extcredits3=42,U.Extcredits4=43,U.Extcredits5=44,U.Extcredits6=45,U.Extcredits7=46,U.UserGroupID=47,U.Medals=48
		Dim LIP
		If GetNoName = 1 Then
			LIP = UClass(23,0)
		Else
			LIP = UClass(51,0)
		End If
		If Not team.Group_Browse(10) = 1 Then 
			If InStr(Lip,".") Then
				LIP = Split(LIP,".")(0) & "." & Split(LIP,".")(1) & ".*.*"
			Else
				LIP = "*.*.*.*"
			End If 
		End If 
		If Page<2 Then
			tmp = Replace(tmp,"{$maxi}","楼主&nbsp;&nbsp; [<span onclick='copyurls()' onmouseover=""javascript:this.style.cursor='hand'"" style=""color:red"">点击复制本网址</span>] ")
			tmp = Replace(tmp,"{$rid}","0")
			tmp = Replace(tmp,"{$nameid}",UClass(0,0))
			tmp = Replace(tmp,"{$topic}",UClass(1,0))
			tmp = Replace(tmp,"{$mod}","a4")
			tmp = Replace(tmp,"{$reward}",iif(Cid(UClass(17,0))=0,"","<img Src="""&team.Styleurl&"/ico.gif"" Border=""0"" Align=""AbsMiddle"" alt=""开始悬赏咯。"">"))
			tmp = Replace(tmp,"{$reaction}","")
			Dim MyContent
			MyContent = ""
			If CID(Board_Setting(7)) = 1 Then
				MyContent = ReadCode(UClass(12,0),Team.Club_Class(1))
			Else
				MyContent = UClass(12,0)
			End If
			MyContent = Ubb_Code(UserBad(MyContent,UClass(2,0)))
			If team.ManageUser and Cid(UClass(9,0)) = 1 Then
				MyContent = MyContent & "<br /><font color=""red"">==此帖已被锁定==</font>"
			End If
			MyContent = ReadPowers(MyContent)
			If GetNoName = 0 Then
				If Cid(UClass(47,0)) = 6 Or Cid(UClass(47,0)) = 7 Then MyContent = "<font color=""red"">==该用户已被锁定==</font>"
			End If
			tmp = Replace(tmp,"{$content}",team.AdvShows(8,fid) & IIF(CID(team.Forum_setting(110)) = 1 And Not team.UserLoginED,"<fieldset class=textquote><legend><strong><FONT COLOR=""red""><B>请登陆论坛查看所有内容</B></FONT></strong></legend>"& cutstr(Replacehtml(MyContent),200) &"</fieldset>",MyContent)  & "<br><dl id=""loadtopicstoplist""><img src=""images/loading.gif""><script>loadtopicstop('"& UClass(2,0) &"','"& tid &"')</script></dl>" ) '简单处理
			tmp = Replace(tmp,"{$smallimg}",iif(Cid(UClass(4,0))>0,"<img src=""images/brow/icon"&UClass(4,0)&".gif"" border=""0"" align=""absmiddle"">",""))
			tmp = Replace(tmp,"{$fortopuser}","{$totop_1}{$totop_2}")
			Select Case UClass(8,0)
				Case 1
					tmp = Replace(tmp,"{$totop_1}","<B>本帖已被置顶</B>")
				Case 2
					tmp = Replace(tmp,"{$totop_1}","<B>本帖已被区置顶</B>")
				Case 3
					tmp = Replace(tmp,"{$totop_1}","<B>本帖已被总置顶</B>")
				Case Else
					tmp = Replace(tmp,"{$totop_1}","")
			End Select
			tmp = Replace(tmp,"{$totop_2}",IIf(Cid(UClass(20,0))=0,iif(Cid(UClass(17,0))>0,"[未解决] <img src="""&team.Styleurl&"/point.gif"" align=""absmiddle""> 本帖悬赏"& IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1," "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" "&UClass(17,0)&" "," 本积分未启用 ")&" ",""),iif(Cid(UClass(17,0))>0,"[已解决]本帖悬赏"& IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  " "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" "&UClass(17,0)&" "," 本积分未启用 ")&" ","")))
			If GetNoName = 1 Then
				tmp = Replace(tmp,"{$username}",IIF(CID(UClass(22,0))=1 And Not team.ManageUser,"匿名用户",UClass(2,0)))
				tmp = Replace(tmp,"{$birthday}","")
				tmp = Replace(tmp,"{$usex}","")
				tmp = Replace(tmp,"{$sign}","")
				tmp = Replace(tmp,"{$userqq}","")
				tmp = Replace(tmp,"{$honor}","")
				tmp = Replace(tmp,"{$userimg}","")
				tmp = Replace(tmp,"{$uid}"," <br>")
				tmp = Replace(tmp,"{$levelname}","")
				tmp = Replace(tmp,"{$regtime}"," ")
				tmp = Replace(tmp,"{$postcount}","<br>")
				tmp = Replace(tmp,"{$online}","<img src="&team.Styleurl&"/offline.gif border='0' align='absmiddle' alt='此用户未登陆!'>")
				tmp = Replace(tmp,"{$masterimg}","")
				tmp = Replace(tmp,"{$userext}","")
				tmp = Replace(tmp,"{$mycity}","")	
				tmp = Replace(tmp,"{$userMedals}","")
			Else
				tmp = Replace(tmp,"{$username}",IIF(CID(UClass(50,0))=1 And Not team.ManageUser,"匿名用户",UClass(2,0)))
				tmp=Replace(tmp,"{$birthday}",Astro(UClass(28,0)))
				tmp=Replace(tmp,"{$usex}",GetUseSex(UClass(29,0)))
				tmp=Replace(tmp,"{$sign}",iif(UClass(30,0)&""="","","<img src="""&team.Styleurl&"/line.gif"" border=""0""><br><div style=""overflow: hidden; max-height: 6em; maxHeight: 77px;"">"& Sign_Code(UClass(30,0),CID(Split(UClass(22,0),"||")(4)))&"</div>"))
				tmp = Replace(tmp,"{$userqq}",iif(Split(UClass(31,0),"|")(0)&""="","","<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="&Split(UClass(31,0),"|")(0)&"&Site=team5.cn&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"&Split(UClass(31,0),"|")(0)&":5 alt=""点击这里给我发消息"" onerror=""javascript:this.src='images/qqerr.gif'""></a>"))
				tmp=Replace(tmp,"{$honor}",IIf(UClass(32,0)<>"",UClass(32,0)&"<br>",""))
				tmp=Replace(tmp,"{$userimg}",iif(UClass(33,0)&""="","","<img src="""&Fixjs(UClass(33,0))&""" border=""0"" onload='javascript:if(this.width>"&CID(team.Forum_setting(108))&")this.width="&CID(team.Forum_setting(108))&";if(this.height>"&CID(team.Forum_setting(109))&")this.height="&CID(team.Forum_setting(109))&";'onerror='javascript:this.src=""images/face/error.gif""'><br>"))
				tmp=Replace(tmp,"{$uid}"," UID "& UClass(34,0) &" <br>")
				tmp=Replace(tmp,"{$stylename}",Split(UClass(22,0),"||")(1))
				tmp=Replace(tmp,"{$levelname}",IIF(CID(UClass(50,0))=1 And Not team.ManageUser,"...",Split(UClass(22,0),"||")(0)))
				tmp=Replace(tmp,"{$regtime}"," 注册 "& FormatDateTime(UClass(26,0),1) &" <br>")
				tmp=Replace(tmp,"{$postcount}","帖子 "& UClass(23,0)+Cid(UClass(24,0)) &"<br>")
				tmp = Replace(tmp,"{$online}", Iif(InStr(UserOnlineinfos,"$$"&UClass(2,0)&"$$")>0, "<img src="&team.Styleurl&"/online.gif border='0' align='absmiddle' alt='此用户在线!&#xA;共计在线时长"&UClass(35,0)&"分钟'>","<img src="&team.Styleurl&"/offline.gif border='0' align='absmiddle' alt='此用户离线!&#xA;共计在线时长"&UClass(35,0)&"分钟'>"))
				tmp = Replace(tmp,"{$masterimg}",UserStar(Split(UClass(22,0),"||")(3))&"<br>"& IIF(Split(UClass(22,0),"||")(2)&""="","","<img src="""&Split(UClass(22,0),"||")(2)&""" border=""0""><br>") &"")
				If CID(UClass(25,0))>0 Then
					emp = emp & "精华&nbsp;" & UClass(25,0) & "&nbsp;<br />"
				End if
				for i = 0 to ubound(ExtCredits)
					If Split(ExtCredits(i),",")(4) =1 Then
						emp = emp & ""& Split(ExtCredits(i),",")(0) & "&nbsp;"& UClass(39 + i,0) &"&nbsp;"& Split(ExtCredits(i),",")(1) &" <br />"
					End if
				Next
				tmp = Replace(tmp,"{$userext}",emp)
				tmp = Replace(tmp,"{$mycity}",iif(UClass(37,0)<>""," 来自 "&UClass(37,0)&" <br>",""))
				Dim UserMedals
				If UClass(48,0)&""<>"" Then
					UserMedals = "" : Emp = ""
					If Instr(UClass(48,0),"$$$")>0 Then
						UserMedals = Split(UClass(48,0),"$$$")
						For i = 0 to Ubound(UserMedals)-1
							Emp = Emp & "<img src=""images/plus/"&Split(UserMedals(i),"&&&")(0)&""" align=""absmiddle"" alt="""&Split(UserMedals(i),"&&&")(1)&"""> "
						Next
						tmp = Replace(tmp,"{$userMedals}",Emp)
					End if
				Else
					tmp = Replace(tmp,"{$userMedals}","")
				End if
			End If
		End If
		If Cid(UClass(5,0))>0 Then
			tmp = tmp & ReMyTopic
		End If
		tmp =  tmp & Team.PostHtml (6)
		If CID(team.Forum_setting(29)) = 1 and UClass(10,0) = 0 Then
			tmp = tmp & Team.PostHtml (7)
		End if	
		'PageList 每页分页数,总记录数,当前页,总页数,当前Url
		tmp = Replace(tmp,"{$pagelister}",team.PageList(PageNum,UClass(5,0),6)) 
		tmp = Replace(tmp,"{$istags}","")
		tmp = Replace(tmp,"{$closemanages}",IIF(team.ManageUser,"","None"))
		tmp = Replace(tmp,"{$tid}",tID)
		tmp = Replace(tmp,"{$surl}",IIF(Request("seesmile")="yes","Thread.asp?tid="& tID & "",team.ActUrl &"&seesmile=yes"))
		tmp = Replace(tmp,"{$surlalt}",IIF(Request("seesmile")="yes","显示关闭","显示更多"))
		tmp = Replace(tmp,"{$maxsml}",Cid(team.Forum_setting(87)))
		tmp = Replace(tmp,"{$fid}",fID)
		tmp = Replace(tmp,"{$rid}",tID)
		If Not Request.Cookies("posttime") = Empty Then
			Dim itts
			itts = "[距离您上次发帖已经<span id=""stime"">3</span> 秒]<script type=""text/javascript"">countup("& DateDiff("s",Request.Cookies("posttime"),Now()) &");</script>"
		End if
		tmp = Replace(tmp,"{$forutime}","")
		tmp = Replace(tmp,"{$TotalPage}",PageNum)
		tmp = Replace(tmp,"{$allPage}",UClass(5,0))
		tmp = Replace(tmp,"{$eazys}",iif(Cid(team.Forum_setting(48))>0,iif(Cid(Session("postnum"))> Cid(team.Forum_setting(48)) ,"","display:none"),"display:none"))
		tmp = Replace(tmp,"{$abouttags}",AboutTipoc)	'相关主题
		tmp = Replace(tmp,"{$postmax}",Cid(team.Forum_setting(67)))
		tmp = Replace(tmp,"{$postmin}",cid(team.Forum_setting(64)))
		Echo tmp
		team.execute("Update ["&IsForum&"Forum] Set Views=Views+1  Where ID="& tID )
	End Sub

	Private Function AboutTipoc()
		Dim TagWhere,Gs,Tagtmp,Tags,i
		TagWhere = ""
		If UClass(21,0)&"" = "" Then
			Exit Function
		End if
		If Instr(UClass(21,0),"|")>=0 Then
			Tags = Split(UClass(21,0),"|")
			for i = 0 to Ubound(Tags)
				If i = 0 Then
					If IsSqlDataBase = 1 Then
						TagWhere = " Topic like '%"&HtmlEncode(Tags(i))&"%' "
					Else
						TagWhere = " InStr(1,LCase(Topic),LCase('"&HtmlEncode(Tags(i))&"'),0)<>0 "
					End if
				Else
					If IsSqlDataBase = 1 Then
						TagWhere = TagWhere & " or Topic like '%"&HtmlEncode(Tags(i))&"%' "
					Else
						TagWhere = TagWhere & " or InStr(1,LCase(Topic),LCase('"&HtmlEncode(Tags(i))&"'),0)<>0 "
					End if
				End if
			Next
		Else
			If IsSqlDataBase = 1 Then
				TagWhere = " Topic like '%"&HtmlEncode(UClass(21,0))&"%' "
			Else
				TagWhere = " InStr(1,LCase(Topic),LCase('"&HtmlEncode(UClass(21,0))&"'),0)<>0 "
			End if
		End If
		Set Gs = team.execute("Select Top 5 ID,Topic,UserName,Views,Replies,Lasttime From ["&Isforum&"Forum] Where deltopic=0 and "&TagWhere&" and Not (ID="&tID&") order By Lasttime Desc")
		If Gs.Eof And Gs.Bof Then
			Exit Function
		Else
			Tagtmp = "<table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""6"" align=""center"" Class=""a2""><tr class=""tab1""><td> 相关主题 </td><td> 作者 </td><td> 查看/回复 </td><td> 最后更新 </td></tr>"
			Do While not Gs.Eof
				Tagtmp = Tagtmp & IIF(CID(team.Forum_setting(65))=1,"<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td><a href=""thread-"&Gs(0)&".html"" target=""_blank"">"&GetColor(Gs(1),UClass(21,0))&"</a></td><td align=""center""> "&Gs(2)&" </td><td align=""center""> "&Gs(3)&" / "&Gs(4)&"</td> <td align=""center""> "&Gs(5)&" </td></tr> ","<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td><a href=""thread.asp?tid="&Gs(0)&""" target=""_blank"">"&GetColor(Gs(1),UClass(21,0))&"</a></td><td align=""center""> "&Gs(2)&" </td><td align=""center""> "&Gs(3)&" / "&Gs(4)&"</td> <td align=""center""> "&Gs(5)&" </td></tr> ")
				Gs.MoveNext
			Loop
			Tagtmp = Tagtmp & " </table> "
		End If
		Gs.Close:Set Gs=Nothing
		AboutTipoc = Tagtmp
	End Function

	Private Function TestUserRead()
		TestUserRead = True
		If Int(UClass(18,0))>0 Then TestUserRead = False
		If team.ManageUser Then TestUserRead = True
		If team.UserLoginED Then
			If Trim(TK_UserName)=Trim(UClass(2,0)) Then TestUserRead = True
			If Cid(Team.Group_Browse(1)) > Cid(UClass(18,0)) Then TestUserRead = True
		End if
	End Function

	Private Sub ActionLive()
		Dim Vs,Rs
		Set Vs = team.execute("Select PlayName,PlayClass,PlayCity,PlayFrom,Playto,Playplace,PlayCost,PlayGender,PlayNum,PlayStop,PlayUserNum From ["&IsForum&"Activity] Where RootID="& tID ) 
		If Vs.Eof And Vs.Bof Then
			Exit Sub
		Else
			tmp = Replace(tmp,"{$postactionsinfo}",Team.PostHtml (10))
			tmp = Replace(tmp,"{$paytopic}",Vs(0))
			tmp = Replace(tmp,"{$playclass}",Vs(1))
			tmp = Replace(tmp,"{$playtime}",iif(Vs(4)<>"",VS(3) &" 至 " & Vs(4) & " 商定",Vs(3)))
			tmp = Replace(tmp,"{$playcity}",Vs(2)&" " & Vs(5))
			tmp = Replace(tmp,"{$playmoney}",Vs(6)&"")
			tmp = Replace(tmp,"{$playsex}",iif(Vs(7)=0,"不限",iif(Vs(7)=1,"男性","女性")))
			tmp = Replace(tmp,"{$playnum}",Vs(8))
			tmp = Replace(tmp,"{$playaction}",Vs(10))
			tmp = Replace(tmp,"{$playclosetime}",Vs(9))
		End If
		If Vs(10) > 0 Then
			Set Rs = team.execute("Select PlayUser,PlayClass,Playtext From ["&IsForum&"ActivityUser] Where RootID="& tID &" and PlayUser='"&tk_UserName&"'")
			If Rs.Eof Then
				tmp = Replace(tmp,"{$msgs}","")
				tmp = Replace(tmp,"{$myinfos}","")
				tmp = Replace(tmp,"{$disabled}","")
			Else
				tmp = Replace(tmp,"{$msgs}","Display:None")
				tmp = Replace(tmp,"{$myinfos}",IIF(Rs(1) = 0,"<tr><td class=""altbg1"" Colspan=""2"">您的加入申请已发出，请等待发起者的审批</td></tr>","<tr><td class=""altbg1"" Colspan=""2"">"& RS(2) &" </td></tr>"))
				tmp = Replace(tmp,"{$disabled}","disabled")
			End If
		Else
			tmp = Replace(tmp,"{$msgs}","")
			tmp = Replace(tmp,"{$myinfos}","")
			tmp = Replace(tmp,"{$disabled}","")
		End if
		Vs.Close:Set Vs=Nothing
	End Sub


	Private Sub PollAction
		Dim Vs,Checktmp,Vote,Numvote,i,umvote,vmp,ump,vip
		Set Vs = team.execute("Select PollClose,Pollday,PollMax,Polltime,Pollmult,Polltopic,PollResult,PollUser From ["&IsForum&"Fvote] Where RootID="& tID) 
		If Not Vs.Eof Then
			tmp = Replace(tmp,"{$postactionsinfo}",Team.PostHtml (11))
			Vote=Split(Vs(5),"|")
			Numvote=Split(Vs(6),"|")
 			vmp = ""
			umvote = 0
			for i = 0 to ubound(Numvote)
				umvote = umvote + Numvote(i)
			next
			for i = 0 to ubound(vote)
				If umvote = 0 Then umvote = 1
				Checktmp = IIf(Vs(4)=1,"<input type=""checkbox"" name=""pollanswers"&i&""" value=""1"" onclick='checkbox(this)' class=""radio"">","<input type=""radio"" name=""pollanswers"" value="""&i&""" onclick='checkbox(this)' class=""radio"">")
				If Vs(7)<>"" Then
					If Instr(Vs(7),"$#$")>0 Then
						If Instr(Vs(7),TK_UserName&"$#$")>0 Then Checktmp = ""
					Else
						If "|"&Trim(Vs(7))&"$$" = "|"&Trim(TK_UserName)&"$$" Then Checktmp = ""
					End if
				End if
				vmp = vmp & "<tr><td class=""altbg1"" width=""22%"">"&Checktmp&" "&vote(i)&"</td><td class=""altbg2""><div class=""percent""><div style=""width:"&(Numvote(i)/umvote)*500&" ""></div></div><div class=""percenttxt""> &nbsp; &nbsp; "&Formatnumber((Numvote(i)/umvote)*100)&"<u>(%) &nbsp;["&Numvote(i)&"票]</u></div></td></tr>"
			next
			tmp = Replace(tmp,"{$votetitle}",UClass(1,0))
			tmp = Replace(tmp,"{$alllpost}",iif(Vs(4)=1,"(多选，最多"&Vs(2)&"项)","(单选)"))
			If Request("showvoters")="yes" Then
				ump = "<tr><td class=""altbg1"" colspan=""2""><B>参与投票的会员:</B><BR><BR>"
				If Vs(7)<>"" Then
					If Instr(VS(7),"$#$")>0 Then
						vip = Split(VS(7),"$#$")
						for i = 0 to ubound(vip)
							ump = ump & IIF(CID(team.Forum_setting(65))=1," <img src="""&team.styleurl&"/gm5.gif"" border=""0"" align=""absmiddle""> <a href=""Profile-"&vip(i)&".html"" target=""_blank""> "& vip(i) &" </a>&nbsp;&nbsp;"," <img src="""&team.styleurl&"/gm5.gif"" border=""0"" align=""absmiddle""> <a href=""Profile.asp?username="&vip(i)&""" target=""_blank""> "& vip(i) &" </a>&nbsp;&nbsp;")
						next
					Else
						ump = ump & IIF(CID(team.Forum_setting(65))=1," <img src="""&team.styleurl&"/gm5.gif"" border=""0"" align=""absmiddle""> <a href=""Profile-"&VS(7)&".html"" target=""_blank""> "& VS(7) &" </a>&nbsp;&nbsp;"," <img src="""&team.styleurl&"/gm5.gif"" border=""0"" align=""absmiddle""> <a href=""Profile.asp?username="&VS(7)&""" target=""_blank""> "& VS(7) &" </a>&nbsp;&nbsp;")
					End if
				End if
				ump = ump & "</td></tr>"
			End if
			tmp = Replace(tmp,"{$voteshow}",vmp)
			tmp = Replace(tmp,"{$ponum}",iif(Vs(4)=1,Cid(Vs(2)),0))
			tmp = Replace(tmp,"{$yesorno}",iif(Request("showvoters")="yes","no","yes"))
			tmp = Replace(tmp,"{$polluser}",iif(Request("showvoters")="yes",ump,""))
			tmp = Replace(tmp,"{$display}",iif(Instr(Vs(7),"|"&TK_UserName&"|")>0 or Cid(Vs(0))=1 or (Cid(Vs(1))>0 And DateDiff("d",CDate(Vs(3)),Date())>Cid(Vs(1))),"disabled=disabled",""))
		End if
		Vs.Close:Set Vs=Nothing
	End Sub

	Private Function ReMyTopic
		Dim Maxi,ReGetNoName,Rtitle
		Dim Trs,SQL,Rs2,UNext,i
		Maxpage = Int(team.Forum_setting(20))
		ReGetNoName = 0
		If Page<2 Then
			Maxi=1
		Else
			Maxi=Page*Maxpage+1-Maxpage
		End If
		If CID(Board_Setting(9))=1 Then
			ReGetNoName = 1
			SQL = "Select ID,topicid,Username,Content,Posttime,Lock,Reward,ReTopic,IsNoName,Auditing,postip From ["& Isforum & UClass(19,0) &"] Where topicid="&tID&" Order By Reward Asc,ID ASC"
		Else
			SQL=" Select T.ID,T.topicid,T.Username,T.Content,T.Posttime,T.Lock,T.Reward,T.ReTopic,U.Levelname,U.Posttopic,U.Postrevert,U.Goodtopic,U.Regtime,U.Landtime,U.Birthday,U.UserSex,U.Sign,U.UserInfo,U.Honor,U.Userface,U.ID,U.Degree,U.Postblog,U.UserCity,U.UserUp,U.Extcredits0,U.Extcredits1,U.Extcredits2,U.Extcredits3,U.Extcredits4,U.Extcredits5,U.Extcredits6,U.Extcredits7,U.UserGroupID,U.Medals,T.IsNoName,T.Auditing,T.PostIp From ["&Isforum & UClass(19,0)&"] T Inner Join ["&IsForum&"User] U On U.UserName=T.UserName Where T.topicid="& tID &" Order By T.ID ASC "
		End if
		Set Rs2 = Server.CreateObject ("Adodb.RecordSet")
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs2.Open SQL,Conn,1,1,&H0001
		PageNum = Abs(int(-Abs(UClass(5,0)/Maxpage)))	'页数
		Page = CheckNum(Page,1,1,1,PageNum)	'当前页	
		If Rs2.Eof and Rs2.Bof Then
			Set Rs2 = Nothing
			Set Rs2 = Server.CreateObject ("Adodb.RecordSet")
			SQL = "Select ID,topicid,Username,Content,Posttime,Lock,Reward,ReTopic,IsNoName,Auditing,PostIp From ["& Isforum & UClass(19,0) &"] Where topicid="&tID&" Order By Reward Asc,ID ASC"
			Rs2.Open Sql,Conn,1,1,&H0001
			If Rs2.Eof and Rs2.Bof Then
				Exit Function
			End If
			ReGetNoName = 1
		End If
		Rs2.AbsolutePosition=(Page-1)*Maxpage+1
		If Rs2.Eof Then	'加入出现错误，采用最费时的算法
			PageNum = Rs2.pagecount
			Page = CheckNum(Page,1,1,1,PageNum)
			Rs2.AbsolutePosition=(Page-1)*Maxpage+1
		End if
		Trs = Rs2.GetRows(Maxpage)
		SqlQueryNum = SqlQueryNum+1
		Rs2.Close:Set Rs2=Nothing
		If Not IsArray(Trs) Then
			Exit Function
		End If
		'T.ID=0,T.topicid=1,T.Username=2,T.Content=3,T.Posttime=4,T.Lock=5,T.Reward=6,T.ReTopic=7,U.Levelname=8,U.Posttopic=9,U.Postrevert=10,U.Goodtopic=11,U.Regtime=12,U.Landtime=13,U.Birthday=14,U.UserSe=15,U.Sign=16,U.UserInfo=17,U.Honor=18,U.Userface=19,U.ID=20,U.Degree=21,U.Postblog=22,U.UserCity=23,U.UserUp=24,U.Extcredits0=25,U.Extcredits1=26,U.Extcredits2=27,U.Extcredits3=28,U.Extcredits4=29,U.Extcredits5=30,U.Extcredits6=31,U.Extcredits7=32,U.UserGroupID=33,U.Medals=34
		Dim tmp1
		tmp1 = ""
		For i=0 To Ubound(Trs,2)
			Maxi = Maxi+1
			If CID(Board_Setting(9))=1 Then
				tmp1 = tmp1 & Team.PostHtml (12)
			Else
				tmp1 = tmp1 & Team.PostHtml (5)
			End If
			Dim LIP
			If ReGetNoName = 1 Then
				LIP = Trs(10,i)
			Else
				LIP = Trs(37,i)
			End If
			If Not team.Group_Browse(10) = 1 Then 
				If InStr(Lip,".") Then
					LIP = Split(LIP,".")(0) & "." & Split(LIP,".")(1) & ".*.*"
				Else
					LIP = "*.*.*.*"
				End If 
			End If
			tmp1 = Replace(tmp1,"{$rid}",Trs(0,i))
			tmp1 = Replace(tmp1,"{$reid}",Trs(0,i))
			tmp1 = Replace(tmp1,"{$nameid}",Trs(0,i))
			tmp1 = Replace(tmp1,"{$reward}",IIF(Cid(UClass(20,0))=1 and Cid(Trs(6,i))=1,"<Img Src="""&team.StyleUrl&"/flag.Gif"" Border=""0"" Align=""AbsMiddle""> <b>最佳答案</b>",""))
			tmp1 = Replace(tmp1,"{$maxi}","第"& Maxi & "楼")
			tmp1 = Replace(tmp1,"{$mod}",iif(Maxi Mod 2=1,"a4","a3"))
			tmp1 = Replace(tmp1,"{$smallimg}","")
			Dim MyLocks
			MyLocks = ""
			If CID(Board_Setting(7))=1 Then
				MyLocks = ReadCode(Trs(3,i),Team.Club_Class(1))
			Else
				MyLocks = Trs(3,i)
			End if
			If CID(Trs(5,i)) = 1 Then
				If team.ManageUser Then
					Mylocks = UBB_Code(UserBad(MyLocks,Trs(2,i))) & "<br /><font color=""red"">==此帖已被锁定==</font>"
				Else
					Mylocks = "<br /><font color=""red"">==此帖已被锁定==</font>"
				End If
			Else
				Mylocks = UBB_Code(UserBad(MyLocks,Trs(2,i)))
			End If
			Mylocks = ReadPowers(Mylocks)
			If ReGetNoName = 0 Then
				If Cid(Trs(33,i)) = 6 Or Cid(Trs(33,i)) = 7 Then Mylocks = "<font color=""red"">==该用户已被锁定==</font>"
			End If
			Rtitle = iif(Trs(7,i)&"" = "","回复："& UClass(1,0) , Trs(7,i))
			If ReGetNoName = 1 Then
				If CID(Trs(9,i)) = 1 Then
					Mylocks = "<h1>==此回复未审核==</h1>"
					Rtitle = " -- "
				End If
			Else
				If CID(Trs(36,i)) = 1 Then
					Mylocks = "<h1>==此回复未审核==</h1>"
					Rtitle = " -- "
				End If
			End If
			tmp1 = Replace(tmp1,"{$topic}",Rtitle)
			tmp1 = Replace(tmp1,"{$content}",IIF(CID(team.Forum_setting(110)) = 1 And Not team.UserLoginED,"<fieldset class=textquote><legend><strong><FONT COLOR=""red""><B>请登陆论坛查看所有内容</B></FONT></strong></legend>"& cutstr(Replacehtml(Mylocks),200) &"</fieldset>",Mylocks))
			If ReGetNoName = 1 Then
				tmp1 = Replace(tmp1,"{$username}",Trs(2,i))
			Else
				tmp1 = Replace(tmp1,"{$username}",IIF(CID(Trs(35,i))=1 And Not team.ManageUser,"匿名用户",Trs(2,i)))
			End if
			tmp1 = Replace(tmp1,"{$lasttime}",Trs(4,i))
			tmp1 = Replace(tmp1,"{$isrept}",Trs(0,i))
			tmp1 = Replace(tmp1,"{$ismanage}",IIf(team.ManageUser,"<input type=""checkbox"" name=""ismanage"" value="&Trs(0,i)&" class=""radio"">",""))
			tmp1 = Replace(tmp1,"{$fortopuser}",team.AdvShows(4,0))
			If ReGetNoName = 1 Then
				tmp1 = Replace(tmp1,"{$username}",IIF(CID(Trs(8,i))=1 And Not team.ManageUser,"匿名用户",Trs(2,i)))
				tmp1 = Replace(tmp1,"{$birthday}","")
				tmp1 = Replace(tmp1,"{$usex}","")
				tmp1 = Replace(tmp1,"{$sign}","")
				tmp1 = Replace(tmp1,"{$userqq}","")
				tmp1 = Replace(tmp1,"{$honor}","")
				tmp1 = Replace(tmp1,"{$userimg}","")
				tmp1 = Replace(tmp1,"{$uid}"," <br>")
				tmp1 = Replace(tmp1,"{$levelname}","")
				tmp1 = Replace(tmp1,"{$regtime}","  <br>")
				tmp1 = Replace(tmp1,"{$postcount}"," <br>")
				tmp1 = Replace(tmp1,"{$online}","<img src="&team.Styleurl&"/offline.gif border='0' 	align='absmiddle' alt='此用户未登陆!'>")
				tmp1 = Replace(tmp1,"{$masterimg}","")
				tmp1 = Replace(tmp1,"{$userext}","")
				tmp1 = Replace(tmp1,"{$mycity}","")
				tmp1 = Replace(tmp1,"{$userMedals}","")
				tmp1 = Replace(tmp1,"{$reaction}","")
			Else
				tmp1 = Replace(tmp1,"{$username}",IIF(CID(Trs(35,i))=1 And Not team.ManageUser,"匿名用户",Trs(2,i)))
				If GetNoName=0 Then
					tmp1 = Replace(tmp1,"{$reaction}",IIf(CID(UClass(20,0))=0 And CID(UClass(49,0))=1,IIf(Trim(UClass(2,0))<>Trim(Trs(2,i)) and Trim(UClass(2,0))=Trim(TK_UserName),"<a href=""Command.asp?action=bestanswer&tid="&tID&"&rid="&Trs(0,i)&""" onclick=""checkclick('您确认要把该回复选为“最佳答案”吗？')"" title=""将该回复选为“最佳答案”""><img src="""&team.StyleUrl&"/right.gif"" border=""0"" />最佳答案</a>",""),""))
				Else
					tmp1 = Replace(tmp1,"{$reaction}","")
				End If
				tmp1=Replace(tmp1,"{$birthday}",Astro(Trs(14,i)))
				tmp1=Replace(tmp1,"{$usex}",GetUseSex(Trs(15,i)))
				tmp1=Replace(tmp1,"{$sign}",iif(Trs(16,i)&""="","","<img src="""&team.Styleurl&"/line.gif"" 	border=""0""><br><div style=""overflow: hidden; max-height: 6em; maxHeight: 77px;"">"&  Sign_Code(Trs(16,i),CID(Split(Trs(8,i),"||")(4))) &"</div>"))
				tmp1 = Replace(tmp1,"{$userqq}",iif(Split(Trs(17,i),"|")(0)&""="","","<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin="&Split(Trs(17,i),"|")(0)&"&Site=team5.cn&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:"&Split(Trs(17,i),"|")(0)&":5 alt=""点击这里给我发消息"" onerror=""javascript:this.src='images/qqerr.gif'""></a>"))
				tmp1=Replace(tmp1,"{$honor}",IIf(Trs(18,i)<>"",Trs(18,i)&"<br>",""))
				tmp1=Replace(tmp1,"{$userimg}",iif(Trs(19,i)&""="","","<img src="""&Fixjs(Trs(19,i))&""" border=""0"" onload='javascript:if(this.width>"&CID(team.Forum_setting(108))&")this.width="&CID(team.Forum_setting(108))&";if(this.height>"&CID(team.Forum_setting(109))&")this.height="&CID(team.Forum_setting(109))&";'onerror='javascript:this.src=""images/face/error.gif""'><br>"))
				tmp1 = Replace(tmp1,"{$uid}"," UID "& Trs(20,i) &" <br>")
				tmp1 = Replace(tmp1,"{$levelname}",IIF(CID(Trs(35,i))=1 And Not team.ManageUser,"...",Split(Trs(8,i),"||")(0)))
				tmp1 = Replace(tmp1,"{$stylename}",Split(Trs(8,i),"||")(1))
				tmp1 = Replace(tmp1,"{$regtime}"," 注册 "& FormatDateTime(Trs(12,i),1) &" <br>")
				tmp1 = Replace(tmp1,"{$postcount}","帖子 "& Trs(9,i)+Cid(Trs(10,i)) &"<br>" )
				tmp1 = Replace(tmp1,"{$online}",Iif(InStr(UserOnlineinfos,"$$"&Trs(2,i)&"$$")>0, "<img src="&team.Styleurl&"/online.gif border='0' align='absmiddle' alt='此用户在线!&#xA;共计在线时长"&Trs(21,i)&"分钟'>","<img src="&team.Styleurl&"/offline.gif border='0' align='absmiddle' alt='此用户离线!&#xA;共计在线时长"&Trs(21,i)&"分钟'>"))
				tmp1 = Replace(tmp1,"{$masterimg}",UserStar(Split(Trs(8,i),"||")(3))&"<br>" & IIF(Split(Trs(8,i),"||")(2)&""="","","<img src="""&Split(Trs(8,i),"||")(2)&""" border=""0""><br>") &"")
				Dim Emp,U,UserMedals
				emp = ""
				If CID(Trs(11,i))>0 Then
					emp = emp & "精华&nbsp;" & Trs(11,i) & "&nbsp;<br />"
				End if
				for u = 0 to ubound(ExtCredits)
					If Split(ExtCredits(u),",")(4) =1 Then
						emp = emp & ""& Split(ExtCredits(u),",")(0) & "&nbsp;"& Trs(25+u,i) &"&nbsp;"& Split(ExtCredits(u),",")(1) &" <br />"
					End if
				Next
				tmp1 = Replace(tmp1,"{$userext}",emp)
				tmp1 = Replace(tmp1,"{$mycity}",iif(Trs(23,i)<>""," 来自 "&Trs(23,i)&" <br>",""))
				u = 0
				If Trs(34,i)&""<>"" Then
					UserMedals = "" :Emp=""
					If Instr(Trs(34,i),"$$$")>0 Then
						UserMedals = Split(Trs(34,i),"$$$")
						For u = 0 to Ubound(UserMedals)-1
							Emp = Emp & "<img src=""images/plus/"&Split(UserMedals(u),"&&&")(0)&""" align=""absmiddle"" alt="""&Split(UserMedals(u),"&&&")(1)&"""> "
						Next
						tmp1 = Replace(tmp1,"{$userMedals}",Emp)
					End if
				Else
					tmp1 = Replace(tmp1,"{$userMedals}","")
				End if
			End If
		Next
		ReMyTopic = tmp1
	End Function
	Private Sub Class_Terminate()
		Err.Clear
		If IsObject(Conn) Then Conn.Close:Set Conn=Nothing
		If IsObject(Cache) Then Cache.Close:Set Cache=Nothing
		If IsObject(MyThread) Then team.Close:Set MyThread=Nothing
		Response.End
	End Sub
End Class
%>