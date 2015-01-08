<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim Fid,tID,Titles,x1,x2,forumid
Public Boards,Board_Setting
fID = Cid(Request("fid")) 
tID = Cid(Request("tid"))
Select Case Request("newpage")
	Case "post"
		Call posts
	Case "reply"
		Call replays
	Case "edit"
		Call edits
	Case Else
		team.error "参数错误"
End Select

Sub posts	
	Dim tmp,ismaste
	Dim ExtCredits
	Titles = "发表新帖"
	ConfigSet()
	X1="发表新帖"
	X2 = Sign_Code(Boards(2,0),1)
	If CID(team.Group_Browse(13)) = 0 Then
		team.Error " 您所在的组没有发帖的权限。"
	End if
	Call NewUserpostTime()
	ExtCredits = Split(team.Club_Class(21),"|")
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(Team.PostHtml (8)),"postaction"))
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"postinfo"))
	tmp = Replace(tmp,"{$weburl}",team.MenuTitle)
	tmp = Replace(tmp,"{$username}",TK_UserName)
	tmp = Replace(tmp,"{$posttime}",Now())
	tmp = Replace(tmp,"{$readperm}","0")
	tmp = Replace(tmp,"{$topics}","")
	tmp = Replace(tmp,"{$ischecked}","")
	tmp = Replace(tmp,"{$resubjet}","")
	tmp = Replace(tmp,"{$postmax}",Cid(team.Forum_setting(67)))
	tmp = Replace(tmp,"{$postmin}",cid(team.Forum_setting(64)))
	tmp = Replace(tmp,"{$topicmax}",cid(team.Forum_setting(89)))
	tmp = Replace(tmp,"{$display}",iif(CID(team.Forum_setting(48))>0,iif(Cid(Session("postnum"))> CID(team.Forum_setting(48)),"","display:none"),"display:none"))
	tmp = Replace(tmp,"{$pollmaxto}",cid(team.Forum_setting(68)))
	tmp = Replace(tmp,"{$isrot}",1)
	tmp = Replace(tmp,"{$messages}","")
	tmp = Replace(tmp,"{$postcolor}","")
	tmp = Replace(tmp,"{$mycolor}",IIf(team.ManageUser,"","None"))
	tmp = Replace(tmp,"{$dispoll}","display:none")
	tmp = Replace(tmp,"{$disactivity}","display:none")
	tmp = Replace(tmp,"{$disreward}","display:none")
	tmp = Replace(tmp,"{$enddatetime}","0")
	tmp = Replace(tmp,"{$maxchoices}","10")
	tmp = Replace(tmp,"{$pollaction}","")
	tmp = Replace(tmp,"{$editpoll}","")
	tmp = Replace(tmp,"{$activityname}","")
	tmp = Replace(tmp,"{$activitycity}","")
	tmp = Replace(tmp,"{$activityplace}","")
	tmp = Replace(tmp,"{$activityclass}","")
	tmp = Replace(tmp,"{$starttimefrom}","")
	tmp = Replace(tmp,"{$starttimeto}","")
	tmp = Replace(tmp,"{$cost}","0")
	tmp = Replace(tmp,"{$activitynumber}","")
	tmp = Replace(tmp,"{$activityexpiration}","")
	tmp = Replace(tmp,"{$rewardprice}","1")
	tmp = Replace(tmp,"{$pollcheck}","")
	tmp = Replace(tmp,"{$closepoll}","")
	tmp = Replace(tmp,"{$ischenks}","display:none")
	tmp = Replace(tmp,"{$isfsopen}",iif(CID(team.Forum_setting(119))=1,"","display:none"))
	tmp = Replace(tmp,"{$fsname}",team.Forum_setting(117))
	tmp = Replace(tmp,"{$distag}","")
	tmp = Replace(tmp,"{$tags}","")
	tmp = Replace(tmp,"{$fid}",Fid)
	tmp = Replace(tmp,"{$actions}","saves")
	tmp = Replace(tmp,"{$revenue}",team.Forum_setting(11))
	tmp = Replace(tmp,"{$wrname}",IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  " ( "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" ) "," (本积分未启用) "))
	tmp = Replace(tmp,"{$setmode}",Cid(team.Forum_setting(98)))
	tmp = Replace(tmp,"{$maxsml}",Cid(team.Forum_setting(87)))
	tmp = Replace(tmp,"{$iscc}",IIF(Len(team.Forum_setting(114))>=5,"<!-- cc视频插件代码/by team board --><object width='72' height='30'><param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='http://union.bokecc.com/flash/plugin.swf?userID="&team.Forum_setting(114)&"&type=team' /><embed src='http://union.bokecc.com/flash/plugin.swf?userID="&team.Forum_setting(114)&"&type=team' type='application/x-shockwave-flash' width='72' height='30' allowScriptAccess='always'></embed></object><!-- cc视频插件代码/by team board -->",""))
	tmp = Replace(tmp,"{$surl}",team.ActUrl)
	Dim Special,utmp,u
	Special = ""
	If Int(Board_Setting(15))=1 and Int(Board_Setting(17))=1 Then
		If Instr(Board_Setting(19),Chr(13)&Chr(10))>0 Then
			utmp = Split(Board_Setting(19),Chr(13)&Chr(10))
			For U=0 To Ubound(utmp)
				Special = Special &" <option value="""&U&""">"& utmp(u) &"</option>" 
			Next
		Else
			Special = "<option value=""0"">"& Board_Setting(19) &"</option>"
		End if	
		tmp = Replace(tmp,"{$posttopic}","")
	Else
		tmp = Replace(tmp,"{$posttopic}","none")
	End If
	tmp = Replace(tmp,"{$topiclist}",Special)
	tmp = Replace(tmp,"{$postaction}","发表新话题")
	If team.IsMaster  Then 
		ismaste = "<INPUT name=""istop"" type=""checkbox"" value=""1"" class=""checkbox"" /> 置顶主题<br/><INPUT name=""isgood"" type=""checkbox"" value=""1"" class=""checkbox"" />  加为精华<br><INPUT name=""islocks"" type=""checkbox"" value=""1"" class=""checkbox"" />  锁定主题<br>"
	End If
	If team.UserLoginED=True Then
		ismaste = ismaste & "<input class=""checkbox"" type=""checkbox"" name=""todiary"" value=""1""> 加入文集<br><input name=""createpoll"" type=""checkbox"" id=""createpoll"" onclick=""expandoptions('divPoll');"" value=""1"" class=""checkbox"" /> 发布投票</label><br/><input name=""creatactivity"" type=""checkbox"" id=""creatactivity"" onclick=""expandoptions('divactivity');"" value=""1"" class=""checkbox"" /> 发起活动</label><br/> "
		If Cid(team.Forum_setting(99)) > 0 Then
			ismaste = ismaste & "<input name=""createreward"" type=""checkbox"" id=""createreward"" onclick=""expandoptions('divreward');"" value=""1"" class=""checkbox"" /> 悬赏问题</label><br/>"
		End If
		If CID(team.Group_Browse(17))=1 Then
			ismaste = ismaste & "<INPUT name=""isnotname"" type=""checkbox"" value=""1"" class=""checkbox"" /> 匿名发贴<br/>"
		End if
		ismaste = ismaste & "<input name=""getmsgforme"" type=""checkbox"" id=""getmsgforme"" value=""1"" class=""checkbox"" /> 订阅回复通知</label>"
	End If
	If CID(Board_Setting(2))>=1 Then
		ismaste = ismaste & "<BR> <FONT COLOR=""red"">注意:此版块的帖子需要审核!</FONT> "
	End if
	Dim PostRanNum,UpTypes
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
	If Len(team.Group_Browse(29))>2 Or InStr(team.Group_Browse(29),",")>0 Then
		UpTypes = team.Group_Browse(29)
	Else
		UpTypes = ReplaceStr(team.Forum_setting(73),"|",",")
	End if
	tmp = Replace(tmp,"{$tid}",tID)
	tmp = Replace(tmp,"{$managesif}",ismaste)
	tmp = Replace(tmp,"{$maxupfile}",team.Forum_setting(71))
	tmp = Replace(tmp,"{$filetype}",UpTypes)
	tmp = Replace(tmp,"{$postrannum}",PostRanNum)
	tmp = Replace(tmp,"{$oneups}",CID(team.Group_Browse(26)))
	Echo tmp
End Sub

Sub edits	
	Dim tmp,ismaste,Rs,Rs1,SQL
	Dim ExtCredits,Us
	Titles = "编辑帖子"
	ConfigSet()
	Call NewUserpostTime()
	ExtCredits = Split(team.Club_Class(21),"|")
	If IsNumEric(Request("retopicid")) Then
		tmp = iHtmlEncode(TempCode(HtmlEncode(Team.PostHtml (8)),"postaction"))
		tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"postinfo"))
	Else
		tmp = iHtmlEncode(BlackTmp(HtmlEncode(Team.PostHtml (8)),"postaction"))
		tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"postinfo"))
	End if
	Set Rs = team.execute( "Select Topic,UserName,Color,PostClass,Content,Createpoll,Creatdiary,Creatactivity,Createreward,Rewardprice,Readperm,Rewardpricetype,Tags,Icon,ReList,Locktopic From ["&IsForum&"Forum] Where ID="& tID )
	If Rs.Eof Then
		team.error " 您查询的帖子不存在。"
	Else
		If Int(RS(15)) = 1 Then
			team.Error "此主题帖已经被锁定，无法进行编辑操作。"
		End if
		X1= IIF(CID(team.Forum_setting(65))=1,"编辑帖子 &raquo; [ <a href=""thread-"&tID&".html""> "&Rs(0)&" </a>]","编辑帖子 &raquo; [ <a href=""thread.asp?tid="&tID&"""> "&Rs(0)&" </a>]")
		X2 = Sign_Code(Boards(2,0),1)
		tmp = Replace(tmp,"{$weburl}",team.MenuTitle)
		Dim Special,utmp,u
		Special = ""
		If Int(Board_Setting(15))=1 and Int(Board_Setting(17))=1 Then
			If Instr(Board_Setting(19),Chr(13)&Chr(10))>0 Then
				utmp = Split(Board_Setting(19),Chr(13)&Chr(10))
				For U=0 To Ubound(utmp)
					Special = Special &" <option value="""&U&""" "
					If Int(Rs(3)) = u Then Special = Special &" SELECTED "
					Special = Special &">"& utmp(u) &"</option>" 
				Next
			Else
				Special = "<option value=""0"" "
				If Int(Rs(3)) = 0 Then Special = Special &" SELECTED "
				Special = Special &" >"& Board_Setting(19) &"</option>"
			End if	
			tmp = Replace(tmp,"{$posttopic}","")
		Else
			tmp = Replace(tmp,"{$posttopic}","none")
		End If
		tmp = Replace(tmp,"{$isfsopen}",iif(CID(team.Forum_setting(119))=1,"","display:none"))
		tmp = Replace(tmp,"{$fsname}",team.Forum_setting(117))
		tmp = Replace(tmp,"{$topiclist}",Special)
		tmp = Replace(tmp,"{$postmax}",Cid(team.Forum_setting(67)))
		tmp = Replace(tmp,"{$postmin}",cid(team.Forum_setting(64)))
		tmp = Replace(tmp,"{$topicmax}",cid(team.Forum_setting(89)))
		tmp = Replace(tmp,"{$display}","display:none")
		tmp = Replace(tmp,"{$pollmaxto}",cid(team.Forum_setting(68)))
		tmp = Replace(tmp,"{$revenue}",team.Forum_setting(11))
		tmp = Replace(tmp,"{$posttime}",Now())
		tmp = Replace(tmp,"{$actions}","edsaves&amp;tid="&tid&"&retopicid="&Request("retopicid")&"")
		tmp = Replace(tmp,"{$fid}",Fid)
		tmp = Replace(tmp,"{$distag}","")
		tmp = Replace(tmp,"{$wrname}",IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  " ( "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" ) "," (本积分未启用) "))
		tmp = Replace(tmp,"{$setmode}",Cid(team.Forum_setting(98)))
		tmp = Replace(tmp,"{$maxsml}",Cid(team.Forum_setting(87)))
		tmp = Replace(tmp,"{$iscc}",IIF(Len(team.Forum_setting(114))>=5,"<!-- cc视频插件代码/by team board --><object width='72' height='30'><param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='http://union.bokecc.com/flash/plugin.swf?userID="&team.Forum_setting(114)&"&type=team' /><embed src='http://union.bokecc.com/flash/plugin.swf?userID="&team.Forum_setting(114)&"&type=team' type='application/x-shockwave-flash' width='72' height='30' allowScriptAccess='always'></embed></object><!-- cc视频插件代码/by team board -->",""))
		tmp = Replace(tmp,"{$postaction}","编辑主题")
		tmp = Replace(tmp,"{$surl}",team.ActUrl)
		If IsNumEric(Request("retopicid")) Then
			Set Rs1=team.execute("Select UserName,ReTopic,Content,Lock From ["&IsForum & Rs(14) &"] Where ID="& Cid(Request("retopicid")) )
			If Rs1.Eof Then
				team.error " 您查询的回帖子存在。"
			Else
				If Int(RS1(3)) = 1 Then
					team.Error "此帖已经被锁定，无法进行编辑操作。"
				End if
				If Trim(Rs1(0)) <> Trim(tk_UserName) Then
					If team.Group_Manage(0) = 0 Then team.Error "您没有编辑他人帖子的权限。"
					Set us = team.execute("Select UserGroupID From ["&IsForum&"User] Where UserName='"&Rs1(0)&"'")
					If Not Us.Eof Then
						Select Case Int(Us(0))
							Case 1
								If Not team.IsMaster Then team.Error "您没有编辑管理员的帖子的权限。"
							Case 2
								If Not ( team.IsMaster Or team.SuperMaster) Then team.Error "您没有编辑超级版主帖子的权限。"
							Case 3
								If  Not (team.IsMaster Or team.SuperMaster Or team.BoardMaster) Then team.Error "您没有编辑版主帖子的权限。"
						End Select
					End If
					Us.Close:Set Us=Nothing
				End If
				tmp = Replace(tmp,"{$isrot}",0)
				tmp = Replace(tmp,"{$resubjet}","<tr><td class=""altbg1"" class=""bold""> 标题 </td><td class=""altbg2"">  <input type=""text"" name=""subject"" id=""subject"" size=""45"" value="""&Rs1(1)&""" tabindex=""103"" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';"" class=""colorblur""> (选填)</td></tr>")
				tmp = Replace(tmp,"{$username}",Rs1(0))
				tmp = Replace(tmp,"{$messages}",EditConts(Rs1(2)))
				tmp = Replace(tmp,"{$maxupfile}",team.Forum_setting(71))
				tmp = Replace(tmp,"{$managesif}",IIF(CID(team.Group_Browse(17))=1,"<INPUT name=""isnotname"" type=""checkbox"" value=""1"" class=""checkbox"" /> 匿名发贴<br/>",""))
			End If
		Else
			If Trim(Rs(1)) <> Trim(tk_UserName) Then
				If team.Group_Manage(0) = 0 Then team.Error "您没有编辑他人帖子的权限。"
				Set us = team.execute("Select UserGroupID From ["&IsForum&"User] Where UserName='"&Rs(0)&"'")
				If Not Us.Eof Then
					Select Case Int(Us(0))
						Case 1
							If Not team.IsMaster Then team.Error "您没有编辑管理员的帖子的权限。"
						Case 2
							If Not ( team.IsMaster Or team.SuperMaster) Then team.Error "您没有编辑超级版主帖子的权限。"
						Case 3
							If  Not (team.IsMaster Or team.SuperMaster Or team.BoardMaster) Then team.Error "您没有编辑版主帖子的权限。"
					End Select
				End If
				Us.Close:Set Us=Nothing
			End If
			tmp = Replace(tmp,"{$isrot}",1)
			tmp = Replace(tmp,"{$resubjet}","")
			tmp = Replace(tmp,"{$username}",Rs(1))
			tmp = Replace(tmp,"{$tags}",Rs(12)&"")
			tmp = Replace(tmp,"{$topics}",Rs(0)&"")
			tmp = Replace(tmp,"{$postcolor}",IIF(Cid(Rs(2))>0,"<OPTION value="""&Rs(2)&""" selected>改变颜色?</OPTION>",""))
			tmp = Replace(tmp,"{$ischecked}",Cid(Rs(13)))
			tmp = Replace(tmp,"{$messages}",EditConts(Rs(4)))
			tmp = Replace(tmp,"{$dispoll}",Iif(Rs(5)=1,"","display:none"))
			tmp = Replace(tmp,"{$disactivity}",Iif(Rs(7)=1,"","display:none"))
			tmp = Replace(tmp,"{$disreward}",Iif(Rs(8)=1,"","display:none"))
			tmp = Replace(tmp,"{$readperm}",Rs(10))
			If team.UserLoginED=True Then
				ismaste = ismaste & "<input class=""checkbox"" type=""checkbox"" name=""todiary"" value=""1""> 加入文集<br>"
			End If
			If CID(Rs(5)) = 1 Then
				ismaste = ismaste & "<input name=""createpoll"" type=""checkbox"" id=""createpoll"" value=""1"" class=""checkbox"" CHECKED/> 编辑投票 </label> <BR>"
			End If
			If CID(Rs(7)) = 1 Then
				ismaste = ismaste & "<input name=""creatactivity"" type=""checkbox"" id=""creatactivity"" value=""1"" class=""checkbox"" CHECKED/> 发起活动</label> <BR>"
			End If
			If CID(team.Group_Browse(17))=1 Then
				ismaste = ismaste & "<INPUT name=""isnotname"" type=""checkbox"" value=""1"" class=""checkbox"" /> 匿名发贴<br/>"
			End if
			ismaste = ismaste & "<input name=""getmsgforme"" type=""checkbox"" id=""getmsgforme"" value=""1"" class=""checkbox"" /> 订阅回复通知</label>"
			tmp = Replace(tmp,"{$managesif}",ismaste)
			If CID(Rs(5))=1 Then
				Set Rs1 = team.execute("Select PollClose,Pollday,PollMax,Polltime,Pollmult,Polltopic From   ["&Isforum&"FVote] Where Rootid="& tID)
				If Not Rs1.Eof Then
					Dim Vomp,Vimp,i
					tmp = Replace(tmp,"{$enddatetime}",Rs1(1))
					tmp = Replace(tmp,"{$maxchoices}",Rs1(2))
					tmp = Replace(tmp,"{$pollaction}","Display:None")
					tmp = Replace(tmp,"{$pollcheck}",iif(Rs1(4)=1,"checked",""))
					tmp = Replace(tmp,"{$ischenks}",iif(Rs1(4)=1,"","display:none"))
					tmp = Replace(tmp,"{$closepoll}",iif(Rs1(0)=1 or (Cid(Rs1(1))>0 And DateDiff("d",CDate(Rs1(3)),Date())>Cid(Rs1(1))),"<input class=""checkbox"" type=""checkbox"" name=""closevote"" value=""1"" checked onclick=""return   false""> 关闭投票<br>","<input class=""checkbox"" type=""checkbox"" name=""closevote"" value=""1""> 关闭投票<br>"))
					If Instr(Rs1(5),"|")>0 Then
						Vomp = Split(Rs1(5),"|")
						for i = 0 to Ubound(Vomp)
							Vimp = Vimp &" <input type=""text"" size=""70"" name=""pollitemid"" value="""&Vomp(i)&""" class=""colorblur"" onfocus=""this.className='colorfocus';"" onblur=""this.className='colorblur';"" readonly/> "
						next
					End if
					tmp = Replace(tmp,"{$editpoll}",iif(Rs1(0)=1 or (Cid(Rs1(1))>0 And DateDiff("d",CDate(Rs1(3)),Date())>Cid(Rs1(1))),"投票已关闭",Vimp))
				End if
				Rs1.Close:Set Rs1=Nothing
			End if
			If CID(Rs(7))=1 Then
				Set Rs1 = team.execute("Select PlayName,PlayCity,Playplace,PlayClass,PlayFrom,Playto,PlayCost,PlayGender,PlayNum,PlayStop,PlayUserNum From   ["&Isforum&"Activity] Where Rootid="& tID)
				If Not Rs1.Eof Then
					tmp = Replace(tmp,"{$activityname}",Rs1(0)&"")
					tmp = Replace(tmp,"{$activitycity}",Rs1(1)&"")
					tmp = Replace(tmp,"{$activityplace}",Rs1(2)&"")
					tmp = Replace(tmp,"{$activityclass}",Rs1(3)&"")
					tmp = Replace(tmp,"{$starttimefrom}",Rs1(4))
					tmp = Replace(tmp,"{$starttimeto}",Rs1(5)&"")
					tmp = Replace(tmp,"{$cost}",Rs1(6))
					tmp = Replace(tmp,"{$activitynumber}",Rs1(8))
					tmp = Replace(tmp,"{$activityexpiration}",Rs1(9))
				End if
				Rs1.Close:Set Rs1=Nothing
			End if
			If CID(Rs(8))=1 Then
				tmp = Replace(tmp,"{$rewardprice}",Rs(9))
			End if
		End If
		tmp = Replace(tmp,"{$mycolor}",IIf(team.ManageUser,"","None"))
		tmp = Replace(tmp,"{$postcolor}","")
		tmp = Replace(tmp,"{$maxupfile}",team.Forum_setting(71))
		Dim PostRanNum,UpTypes
		Randomize
		PostRanNum = Int(900*rnd)+1000
		Session("UploadCode") = Cstr(PostRanNum)
		If Len(team.Group_Browse(29))>2 Or InStr(team.Group_Browse(29),",")>0 Then
			UpTypes = team.Group_Browse(29)
		Else
			UpTypes = ReplaceStr(team.Forum_setting(73),"|",",")
		End if
		tmp = Replace(tmp,"{$filetype}",UpTypes)
		tmp = Replace(tmp,"{$postrannum}",PostRanNum)
		tmp = Replace(tmp,"{$tid}",tID)
		tmp = Replace(tmp,"{$oneups}",CID(team.Group_Browse(26)))
		Echo tmp
	End if
End Sub

Sub replays
	Dim tmp,ismaste
	Dim ExtCredits
	Titles = "回复帖子"
	ConfigSet()
	X1="回复帖子"
	X2 = Sign_Code(Boards(2,0),1)
	If CID(team.Group_Browse(14)) = 0 Then
		team.Error " 您所在的组没有回复贴子的权限。"
	End If
	If CID(Board_Setting(5)) = 1 Then
		team.Error " 本版块设置了回帖限制，您无法对此版块的帖子发表任何评论回复。"
	End If 	
	Call NewUserpostTime()
	ExtCredits = Split(team.Club_Class(21),"|")
	tmp = iHtmlEncode(TempCode(HtmlEncode(Team.PostHtml (8)),"postaction"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"postinfo"))
	tmp = Replace(tmp,"{$weburl}",team.MenuTitle)
	tmp = Replace(tmp,"{$username}",TK_UserName)
	tmp = Replace(tmp,"{$isfsopen}",iif(CID(team.Forum_setting(119))=1,"","display:none"))
	tmp = Replace(tmp,"{$fsname}",team.Forum_setting(117))
	tmp = Replace(tmp,"{$posttime}",Now())
	tmp = Replace(tmp,"{$readperm}","0")
	tmp = Replace(tmp,"{$topics}","")
	tmp = Replace(tmp,"{$ischecked}","")
	tmp = Replace(tmp,"{$mycolor}","None")
	tmp = Replace(tmp,"{$resubjet}","<tr><td class=""altbg1"" class=""bold""> 标题 </td><td class=""altbg2"">  <input type=""text"" name=""subject"" id=""subject"" size=""45"" value="""" tabindex=""103"" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';"" class=""colorblur""> (选填)</td></tr>")
	tmp = Replace(tmp,"{$postmax}",Cid(team.Forum_setting(67)))
	tmp = Replace(tmp,"{$postmin}",cid(team.Forum_setting(64)))
	tmp = Replace(tmp,"{$topicmax}",cid(team.Forum_setting(89)))
	tmp = Replace(tmp,"{$display}",iif(CID(team.Forum_setting(48))>0,iif(Cid(Session("postnum"))> CID(team.Forum_setting(48)),"","display:none"),"display:none"))
	tmp = Replace(tmp,"{$pollmaxto}",cid(team.Forum_setting(68)))
	tmp = Replace(tmp,"{$isrot}",0)
	If Request("quote") = 1 Then
		Dim Rs
		Set Rs = team.execute("select Content,ReList,UserName,Posttime From ["&IsForum&"Forum] Where Deltopic = 0 and CloseTopic = 0 and ID="&tID)
		If Rs.Eof Then
			team.Error "您回复的帖子ID错误。"
		Else
			If Request("isrept") = "TOPS" Or Not IsNumeric(Request("isrept")) then	
				tmp = Replace(tmp,"{$messages}","[quote]<b>以下是引用<i>"& Rs(2) &"</i>在"& Rs(3) &"的发言：</b><br>"& EditConts(Rs(0)) &"[/quote]")
			Else
				Set Rs = team.execute("select content,UserName,Posttime from ["&IsForum & RS(1) &"] Where ID="& HRF(2,2,"isrept") )
				If Rs.Eof Then
					team.Error "您引用的帖子ID错误。"
				Else
					tmp = Replace(tmp,"{$messages}","[quote]<b>以下是引用<i>"& Rs(1) &"</i>在"& Rs(2) &"的发言：</b><br>"& EditConts(Rs(0)) &"[/quote]")
				End if
			End if
		End If
		Rs.Close:Set Rs=Nothing
	Else
		tmp = Replace(tmp,"{$messages}","")
	End if
	tmp = Replace(tmp,"{$postcolor}","")
	tmp = Replace(tmp,"{$dispoll}","display:none")
	tmp = Replace(tmp,"{$disactivity}","display:none")
	tmp = Replace(tmp,"{$disreward}","display:none")
	tmp = Replace(tmp,"{$enddatetime}","0")
	tmp = Replace(tmp,"{$maxchoices}","10")
	tmp = Replace(tmp,"{$pollaction}","")
	tmp = Replace(tmp,"{$editpoll}","")
	tmp = Replace(tmp,"{$activityname}","")
	tmp = Replace(tmp,"{$activitycity}","")
	tmp = Replace(tmp,"{$activityplace}","")
	tmp = Replace(tmp,"{$activityclass}","")
	tmp = Replace(tmp,"{$starttimefrom}","")
	tmp = Replace(tmp,"{$starttimeto}","")
	tmp = Replace(tmp,"{$cost}","0")
	tmp = Replace(tmp,"{$activitynumber}","")
	tmp = Replace(tmp,"{$activityexpiration}","")
	tmp = Replace(tmp,"{$rewardprice}","1")
	tmp = Replace(tmp,"{$pollcheck}","")
	tmp = Replace(tmp,"{$closepoll}","")
	tmp = Replace(tmp,"{$ischenks}","display:none")
	tmp = Replace(tmp,"{$distag}","")
	tmp = Replace(tmp,"{$tags}","")
	tmp = Replace(tmp,"{$fid}",Fid)
	tmp = Replace(tmp,"{$actions}","resaves&amp;tid="&tid&"")
	tmp = Replace(tmp,"{$revenue}",team.Forum_setting(11))
	tmp = Replace(tmp,"{$wrname}",IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  " ( "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" ) "," (本积分未启用) "))
	tmp = Replace(tmp,"{$setmode}",Cid(team.Forum_setting(98)))
	tmp = Replace(tmp,"{$maxsml}",Cid(team.Forum_setting(87)))
	tmp = Replace(tmp,"{$iscc}",IIF(Len(team.Forum_setting(114))>=5,"<!-- cc视频插件代码/by team board --><object width='72' height='30'><param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='http://union.bokecc.com/flash/plugin.swf?userID="&team.Forum_setting(114)&"&type=team' /><embed src='http://union.bokecc.com/flash/plugin.swf?userID="&team.Forum_setting(114)&"&type=team' type='application/x-shockwave-flash' width='72' height='30' allowScriptAccess='always'></embed></object><!-- cc视频插件代码/by team board -->",""))
	tmp = Replace(tmp,"{$surl}",team.ActUrl)
	tmp = Replace(tmp,"{$postaction}","回复帖子")
	If CID(team.Group_Browse(17))=1 Then
		ismaste = ismaste &"<INPUT name=""isnotname"" type=""checkbox"" value=""1"" class=""checkbox"" /> 匿名回复<br/>"
	End if
	If CID(Board_Setting(2))>=1 Then
		ismaste = ismaste & "<BR> <FONT COLOR=""red"">注意:此版块的帖子需要审核!</FONT> "
	End if
	tmp = Replace(tmp,"{$managesif}",ismaste)
	tmp = Replace(tmp,"{$maxupfile}",team.Forum_setting(71))
	Dim PostRanNum,UpTypes
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
	If Len(team.Group_Browse(29))>2 Or InStr(team.Group_Browse(29),",")>0 Then
		UpTypes = team.Group_Browse(29)
	Else
		UpTypes = ReplaceStr(team.Forum_setting(73),"|",",")
	End if
	tmp = Replace(tmp,"{$filetype}",UpTypes)
	tmp = Replace(tmp,"{$postrannum}",PostRanNum)
	tmp = Replace(tmp,"{$tid}",tID)
	tmp = Replace(tmp,"{$oneups}",CID(team.Group_Browse(26)))
	Echo tmp
End Sub

Private Function GetUserPostPower()
	GetUserPostPower = False
	Dim B_Lookperm,m
	B_Lookperm = Split(Boards(13,0),",")
	If Isarray(B_Lookperm) Then
		For m = 0 to Ubound(B_Lookperm)-1
			If Cid(B_Lookperm(m)) = Int(team.UserGroupID) Then GetUserPostPower = True
		Next 
	End  If
End Function

Sub ConfigSet()
	Dim Rs
	Cache.Name = "Boards_"&Fid
	Cache.Reloadtime = Cid(team.Forum_setting(44))
	If Not Cache.ObjIsEmpty() Then
		Boards = Cache.Value
	Else
		Set Rs=team.Execute("Select ID,Followid,bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Board_Key,Board_URL,toltopic,tolrestore,Board_Code,postperm From ["&IsForum&"Bbsconfig] Where  ID = "& Int(Fid))
		If Rs.Eof Then 
			Team.Error "你查询的版面ID错误。"
			Exit Sub
		Else
			Boards = Rs.GetRows(-1)
			Cache.Value = Boards
		End If
		RS.Close:Set RS=Nothing
	End If
	If isarray(Boards) Then
		Board_Setting = Split(Boards(3,0),"$$$")
	End if
	team.Headers(Boards(2,0) & " - " & Titles)
	team.ChkPost()
	If CID(Board_Setting(9))= 0 Then
		If Not team.UserLoginED Then team.Error " 您所在的组没有此动作的权限。"
	End If
	If Request("newpage") = "post" Then
		If Not (Boards(13,0) = ",") Then
			If Instr(Boards(13,0),",") > 0 Then 
				If Not GetUserPower Then team.Error "您所在的组没有在本版发表新帖的权限"
			End If
		End If
	End if
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

'过虑多余的编辑代码
Function EditConts(strContent)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="<p align=right><font color=#000066>(.*?)<\/font><\/p>"
	strContent=re.Replace(strContent,"")
	set re=Nothing
	EditConts = Server.HtmlEncode(strContent)
End Function

Sub NewUserpostTime()
	If CID(Board_Setting(9))=1 Then Exit Sub
	If Cid(team.Forum_setting(14))>0 And team.UserLoginED And Not team.ManageUser Then
		If Not IsDate(team.User_SysTem(9)) Then team.User_SysTem(9) = Now()
		If DateDiff("s",CDate(team.User_SysTem(9)),Now()) < Cid(team.Forum_setting(14))*60 Then 
			team.error "新注册用户必须停留 <font color=red> "&team.Forum_setting(14)&" </font> 分钟以上才可发表帖子。"
		End if
	End If
End Sub
team.footer
%>