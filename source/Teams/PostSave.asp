<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim fID,tID,PostSave,CutKey
fID = Cid(Request("fid"))
tID = Cid(Request("tid"))
CutKey = 50		'����ҳ��ʾ���ַ�����Ҫ��ȡ���ֶγ���
Set PostSave = New MypostSave
Set PostSave = Nothing
Class MypostSave
	Public Boards,Board_Setting
	Private Sub Class_Initialize()
		Select Case Request("action")
			Case "saves"	'��������
				Call newsave
			Case "resaves"	'�ظ�����
				Call resaves
			Case "edsaves"	'�༭����
				Call edsaves
			Case Else
				Team.Error("��������!")
		End Select
	End Sub
	Private	Sub newsave
		Dim Subject,Message,SQL,Rs,uID,Code,readperm,IsColor,Auditing
		Dim CacheRs,Rewardprice,NextID,Istop,isgood,info,SetMyTime,Settotime
		Call ConfigSet
		Call CheckpostTime
		Call NewUserpostTime()
		Istop = Cid(Request.Form("istop"))
		Isgood = Cid(Request.Form("isgood"))
		Subject = HtmlEncode(Trim(Request.Form("subject")))
		Message = team.Checkstr(Trim(Request.Form("message")))
		If CID(team.Group_Browse(13)) = 0 Then
			team.Error " �����ڵ���û�з�����Ȩ�ޡ�"
		End if
		If strLength(Subject)<=cid(team.Forum_setting(64)) Then
			team.error " ���������������� "&team.Forum_setting(64)&" ���ַ��� "
		ElseIf strLength(Subject)>cid(team.Forum_setting(89)) Then
			team.error " �����������ܴ��� "&team.Forum_setting(89)&" ���ַ��� "
		ElseIf strLength(Message)<=cid(team.Forum_setting(64)) Then
			team.error " ���������������� "&team.Forum_setting(64)&" ���ַ��� "
		ElseIf strLength(Message)>cid(team.Forum_setting(67)) Then
			team.error " �����������ܶ��� "&team.Forum_setting(67)&" ���ַ��� "
		End If
		If CID(Board_Setting(16)) = 1 Then
			If CID(Request.Form("posttopic")) = 999 Then team.Error2 " ��ѡ��ר�����"
		End if
		If CID(team.Group_Browse(18)) = 0 Then
			readperm = 0
		Else
			readperm = Cid(Request.Form("readperm"))
		End If
		Code = Trim(Request.Form("code"))
		If Cid(team.Forum_setting(48)) >0 Then
			If Cid(Session("postnum"))> Cid(team.Forum_setting(48)) then
				if Not Team.CodeIsTrue(code) Then
					team.error "��֤�������ˢ�º�������д��"
				End If
			End If
		End if
		If Cid(Request.Form("createpoll")) = 1 Then
			if Request.Form("pollitemid") = "" Then 
				team.Error " ͶƱ���Ϊ�ա�"
			End if
		End If
		If Cid(Request.Form("creatactivity")) = 1 Then
			If Request.Form("activityname") = "" Then
				team.Error " ����Ʋ���Ϊ�� "
			End if
			If Request.Form("activityclass") = "" Then
				team.Error " ������Ϊ�� "
			End if		
			If Request.Form("activityplace") = "" Then
				team.Error " ��ص㲻��Ϊ�� "
			End If
			If Request.Form("activitytime") = "0" Then 
				If Request.Form("starttimefrom1") = "" Then
					team.Error " �ʱ�䲻��Ϊ�� "
				End If
				SetMyTime = Request.Form("starttimefrom1")
				Settotime = ""
			Else
				If Request.Form("starttimefrom2") = "" Then
					team.Error " ���ʼʱ�䲻��Ϊ�� "
				End If
				SetMyTime = Request.Form("starttimefrom2")
				Settotime = Request.Form("starttimeto")
			End If
		End if
		If Cid(Request.Form("createreward")) = 1 Then
			Rewardprice = Cid(Request.Form("rewardprice"))
		Else
			Rewardprice = 0
		End If
		If CID(team.Group_Browse(19)) = 0 Then
			IsColor = 0
		Else
			IsColor = Cid(Request.Form("color"))
		End If
		Auditing = 0
		If CID(Board_Setting(2))>=1 Then
			If CID(team.Group_Browse(16))<2 Then
				Auditing = 1
			End If 
		End if
		SQL="insert into ["&Isforum&"forum] (Forumid,Topic,Username,Views,Icon,Replies,Color,PostClass,Content,Toptopic,Locktopic,Goodtopic,Posttime,Lasttime,LastText,Postip,Createpoll,Creatdiary,Creatactivity,Createreward,Rewardprice,Readperm,ReList,Rewardpricetype,Tags,CloseTopic,Deltopic,RunMsg,IsNoName,Auditing) values ("&fid&",'"&Subject&"','"&TK_UserName&"',0,"&Cid(Request.Form("icon"))&",0,"& IsColor &","&Cid(Request.Form("posttopic"))&",'"&Message&"',"&Istop&","&Cid(Request.Form("islocks"))&","&Isgood&","&SqlNowString&","&SqlNowString&",'"&TK_UserName&"$@$"&Cutstr(Subject,150)&"','"&Remoteaddr&"',"&Cid(Request.Form("createpoll"))&","&Cid(Request.Form("todiary"))&","&Cid(Request.Form("creatactivity"))&","&Cid(Request.Form("createreward"))&","&Rewardprice&","& readperm &",'"&team.Club_Class(11)&"',0,'"&HtmlEncode(Request.Form("tags"))&"',0,0,"&Cid(Request.Form("getmsgforme"))&","&Cid(Request.Form("isnotname"))&","&Auditing&") "
		team.execute (SQL)
		Set Rs = team.execute("Select Max(ID) from [forum] Where Forumid="&Int(Fid))
		If Not Rs.Eof Then
			uID = Rs(0)
		End if
		Rs.Close:Set Rs=Nothing
		If Cid(Request.Form("createpoll")) = 1 Then
			Dim Pollitemid,PollResult,PollResultmax,i
			Pollitemid = Replace(Replace(HtmlEncode(Request.Form("pollitemid"))," ",""),",","|")
			PollResult = Split(Pollitemid,"|")
			for i = 0 to Ubound(PollResult)
				if i = 0 Then
					PollResultmax = "0"
				else
					PollResultmax = PollResultmax &"|0"
				end if
			next
			team.execute ("insert into ["&Isforum&"FVote] (Rootid,PollClose,Pollday,PollMax,Polltime,Pollmult,Polltopic,PollResult) values ("&uID&",0,"&Cid(Request.Form("enddatetime"))&","&Cid(Request.Form("maxchoices"))&","&SqlNowString&","&Cid(Request.Form("multiplepoll"))&",'"&Pollitemid&"','"&PollResultmax&"') ")
			info = " <li> ͶƱ������� ��"
		End if
		If Cid(Request.Form("creatactivity")) = 1 Then
			team.execute ("insert into ["&Isforum&"Activity] (Rootid,PlayName,PlayClass,PlayCity,PlayFrom,Playto,Playplace,PlayCost,PlayGender,PlayNum,PlayStop,PlayUserNum) values ("&uID&",'"&HtmlEncode(Request.Form("activityname"))&"','"&HtmlEncode(Request.Form("activityclass"))&"','"&HtmlEncode(Request.Form("activitycity"))&"','"&Replace(SetMyTime,"'","''")&"','"&Replace(Settotime,"'","''")&"','"&HtmlEncode(Request.Form("activityplace"))&"',"&Cid(Request.Form("cost"))&","&Cid(Request.Form("gender"))&","&Cid(Request.Form("activitynumber"))&",'"&Replace(Request.Form("activityexpiration"),"'","''")&"',0) ")
			info = " <li> �������� ��"
		End if
		If Request.form("upfileid")<>"" Then
			Dim NewID
			NewID=Split(Replace(Replace(Request.Form("upfileid")," ",""),"'",""),",")
			For i=0 To Ubound(newid)-1
				team.execute("Update ["&IsForum&"Upfile] Set ID= "&Int(uID)&" Where FileID="& Int(newid(i)))
			Next
		End If
		Call UpdateUserpostExc()
		NextID = team.execute("Select Followid From [bbsconfig] Where ID="& fID )(0)
		team.execute("Update ["&IsForum&"Bbsconfig] Set Board_Last='<A href="& IIF(CID(team.Forum_setting(65))=1,"Thread-"&uID&".html","Thread.asp?tid="&uID&"") &" target=""_blank"">"&Cutstr(Subject,CutKey)&"</a>$@$"&TK_UserName&"$@$"&Now()&"',Toltopic=Toltopic+1,today=today+1 Where ID = "& Int(fID) )
		team.execute("update ["&IsForum&"ClubConFig] Set today=today+1,PostNum=PostNum+1")
		team.LockCache "PostNum" , Application(CacheName&"_PostNum")+1
		team.LockCache "TodayNum" , Application(CacheName&"_TodayNum")+1
		Set CacheRs = Team.Execute("Select Followid,ID From ["&IsForum&"bbsconfig] Where ID="& Int(NextID))
		If Not CacheRs.Eof Then
			If Cid(CacheRs(0)) > 0 Then
				NextID = CacheRs(0)
				team.execute("Update ["&IsForum&"Bbsconfig] Set Board_Last='<A href="& IIF(CID(team.Forum_setting(65))=1,"Thread-"&uID&".html","Thread.asp?tid="&uID&"") &" target=""_blank"">"&Cutstr(Subject,CutKey)&"</a>$@$"&TK_UserName&"$@$"&Now()&"',today=today+1 Where ID = "& Cid(CacheRs(1)))
			End If
		End If
		CacheRs.Close:Set CacheRs=Nothing
		Cache.DelCache("BoardLists")
		If Cid(team.Forum_setting(48)) >0 Then
			Session("postnum") = Session("postnum")+1
		End if
		Response.Cookies("posttime") = Now()
		team.error1 info & IIF(CID(team.Forum_setting(65))=1,"�����ⷢ��ɹ�<li><a href=Thread-"&uID&".html>��������</a><li><a href=Forums-"&fid&".html>������̳</a><meta http-equiv=refresh content=3;url=Thread-"&uID&".html> ","<li>�����ⷢ��ɹ�<li><a href=Thread.asp?tid="&uID&">��������</a><li><a href=Forums.asp?fid="&fid&">������̳</a><meta http-equiv=refresh content=3;url=Thread.asp?tid="&uID&"> ")
	End Sub


	Private Sub edsaves
		Dim Subject,Message,SQL,Rs,uID,Code,Istop,isgood,info,Rewardprice
		Dim ReList,PostNames,SetMyTime,Settotime,PostTime,readperm,IsColor
		Call ConfigSet
		'Istop = Cid(Request.Form("istop"))
		'Isgood = Cid(Request.Form("isgood"))
		Subject = HtmlEncode(Trim(Request.Form("subject")))
		Message = team.Checkstr(Trim(Request.Form("message")))
		Code = Trim(Request.Form("code"))
		If strLength(Message)<=cid(team.Forum_setting(64)) Then
			team.error " ���������������� "&team.Forum_setting(64)&" ���ַ��� "
		ElseIf strLength(Message)>cid(team.Forum_setting(67)) Then
			team.error " �����������ܶ��� "&team.Forum_setting(67)&" ���ַ��� "
		End If
		Set rs = team.execute("select ReList,UserName,PostTime,Locktopic From ["&IsForum&"Forum] Where ID="&tID)
		If Rs.Eof Then
			team.Error "���༭������ID����"
		Else
			If Int(RS(3)) = 1 Then
				team.Error "���������Ѿ����������޷����б༭������"
			End if
			ReList = Rs(0)
			PostNames = Rs(1)
			PostTime = Rs(2)
		End If
		Rs.Close:Set Rs=Nothing
		If IsNumEric(Request("retopicid")) Then
			Set Rs = team.execute("Select UserName,PostTime,Lock From ["&IsForum & ReList &"] Where ID = "& Cid(Request("retopicid")))
			If Rs.Eof Then
				team.Error "���༭������ID����"
			Else
				If Int(RS(2)) = 1 Then
					team.Error "�����Ѿ����������޷����б༭������"
				End if
				If Not UpPostName(Rs(0)) Then
					team.Error "��û�б༭�������ӵ�Ȩ�ޡ�"
				Else
					If Int(team.Forum_setting(94)) > 0 And DateDiff("s",RS(1),Now()) > team.Forum_setting(94) And Not team.ManageUser Then
						team.Error " �������Ѿ������˱༭ʱ�ޣ����޷��༭�ˡ�"
					Else
						team.execute ("Update ["&IsForum & ReList &"] Set ReTopic='"&Subject&"',Content='"&UpEditTags(Message,RS(1))&"',IsNoName="& Cid(Request.Form("isnotname")) &" Where ID = "& Cid(Request("retopicid")) )
						info = info & " <li> �����༭��� ��"
					End if
				End if
			End if
		Else
			If Not UpPostName(PostNames) Then
				team.Error "��û�б༭�������ӵ�Ȩ�ޡ�"
			End if
			If strLength(Subject)<cid(team.Forum_setting(64)) Then
				team.error " ���������������� "&team.Forum_setting(64)&" ���ַ��� "
			ElseIf strLength(Subject)>cid(team.Forum_setting(89)) Then
				team.error " �����������ܴ��� "&team.Forum_setting(89)&" ���ַ��� "
			End if
			If Cid(Request.Form("createpoll")) = 1 Then
				if Request.Form("pollitemid") = "" Then 
					team.Error " ͶƱ���Ϊ�ա�"
				End if
			End If
			If Cid(Request.Form("creatactivity")) = 1 Then
				If Request.Form("activityname") = "" Then
					team.Error " ����Ʋ���Ϊ�� "
				End if
				If Request.Form("activityclass") = "" Then
					team.Error " ������Ϊ�� "
				End if		
				If Request.Form("activityplace") = "" Then
					team.Error " ��ص㲻��Ϊ�� "
				End If
				If Request.Form("activitytime") = 0 Then 
					If Request.Form("starttimefrom1") = "" Then
						team.Error " �ʱ�䲻��Ϊ�� " & Request.Form("starttimefrom1")
					End If
					SetMyTime = Request.Form("starttimefrom1")
					Settotime = ""
				Else
					If Request.Form("starttimefrom2") = "" Then
						team.Error " ���ʼʱ�䲻��Ϊ�� "
					End If
					SetMyTime = Request.Form("starttimefrom2")
					Settotime = Request.Form("starttimeto")
				End If
			End if
			If Cid(Request.Form("createreward")) = 1 Then
				Rewardprice = Cid(Request.Form("rewardprice"))
			Else
				Rewardprice = 0
			End If
			If CID(team.Group_Browse(18)) = 0 Then
				readperm = 0
			Else
				readperm = Cid(Request.Form("readperm"))
			End If
			If CID(team.Group_Browse(19)) = 0 Then
				IsColor = 0
			Else
				IsColor = Cid(Request.Form("color"))
			End if		
			If strLength(Subject)<=cid(team.Forum_setting(64)) Then
				team.error " ���������������� "&team.Forum_setting(64)&" ���ַ��� "
			ElseIf strLength(Subject)>cid(team.Forum_setting(89)) Then
				team.error " �����������ܴ��� "&team.Forum_setting(89)&" ���ַ��� "
			End If
			team.execute ("Update ["&Isforum&"forum] Set Topic='"&Subject&"',Icon="&Cid(Request.Form("icon"))&",Color="& IsColor &",PostClass="&Cid(Request.Form("posttopic"))&",Content='"&UpEditTags(Message,PostTime)&"',Lasttime="&SqlNowString&",LastText='"&TK_UserName&"$@$"&Cutstr(Subject,150)&"',Createpoll="&Cid(Request.Form("createpoll"))&",Creatdiary="&Cid(Request.Form("todiary"))&",Creatactivity="&Cid(Request.Form("creatactivity"))&",Createreward="&Cid(Request.Form("createreward"))&",Rewardprice="&Rewardprice&",Readperm="& readperm &",ReList='"&team.Club_Class(11)&"',Tags='"&HtmlEncode(Request.Form("tags"))&"',RunMsg="&Cid(Request.Form("getmsgforme"))&",IsNoName="& Cid(Request.Form("isnotname")) &" Where ID="&Int(tid))
			If Cid(Request.Form("createpoll")) = 1 Then
				'Dim Pollitemid,PollResult,PollResultmax,i
				'Pollitemid = Replace(Replace(HtmlEncode(Request.Form("pollitemid"))," ",""),",","|")
				'PollResult = Split(Pollitemid,"|")
				'for i = 0 to Ubound(PollResult)
					'if i = 0 Then
						'PollResultmax = "0"
					'else
						'PollResultmax = PollResultmax &"|0"
					'end if
				'next
				'team.execute ("Update ["&Isforum&"FVote] Set Pollday="&Cid(Request.Form("enddatetime"))&",PollMax="&Cid(Request.Form("maxchoices"))&",Pollmult="&Cid(Request.Form("multiplepoll"))&",Polltopic='"&Pollitemid&"',PollResult='"&PollResultmax&"' Where RootID="&Int(tid))
				team.execute ("Update ["&Isforum&"FVote] Set Pollday="&Cid(Request.Form("enddatetime"))&",PollMax="&Cid(Request.Form("maxchoices"))&",Pollmult="&Cid(Request.Form("multiplepoll"))&",PollClose="&Cid(Request.Form("closevote"))&" Where RootID="&Int(tid))
				info = " <li> ͶƱ������� ��"
			End if
			If Cid(Request.Form("creatactivity")) = 1 Then
				team.execute ("Update ["&Isforum&"Activity] Set PlayName='"&HtmlEncode(Request.Form("activityname"))&"',PlayClass='"&HtmlEncode(Request.Form("activityclass"))&"',PlayCity='"&HtmlEncode(Request.Form("activitycity"))&"',PlayFrom='"&SetMyTime&"',Playto='"&Replace(Settotime,"'","''")&"',Playplace='"&HtmlEncode(Request.Form("activityplace"))&"',PlayCost="&Cid(Request.Form("cost"))&",PlayGender="&Cid(Request.Form("gender"))&",PlayNum="&Cid(Request.Form("activitynumber"))&",PlayStop='"&Replace(Request.Form("activityexpiration"),"'","''")&"' Where RootID="&Int(tid))
				info = " <li> �������� ��"
			End if
			Dim CacheRs,NextID
			NextID = team.execute("Select Followid From ["&IsForum&"bbsconfig] Where ID="& fID )(0)
			team.execute("Update ["&IsForum&"Bbsconfig] Set Board_Last='<A href="& IIF(CID(team.Forum_setting(65))=1,"Thread-"&tID&".html","Thread.asp?tid="&tID&"") &" target=""_blank"">"&Cutstr(Subject,CutKey)&"</a>$@$"&TK_UserName&"$@$"&Now()&"' Where ID = "& fID )
			Set CacheRs = Team.Execute("Select Followid,ID From ["&IsForum&"bbsconfig] Where ID="& NextID)
			If Not CacheRs.Eof Then
				If Cid(CacheRs(0)) > 0 Then
					NextID = CacheRs(0)
					team.execute("Update ["&IsForum&"Bbsconfig] Set Board_Last='<A href="& IIF(CID(team.Forum_setting(65))=1,"Thread-"&tID&".html","Thread.asp?tid="&tID&"") &" target=""_blank"">"&Cutstr(Subject,CutKey)&"</a>$@$"&TK_UserName&"$@$"&Now()&"' Where ID = "& Cid(CacheRs(1)))
				End If
			End If
			CacheRs.Close:Set CacheRs=Nothing
			Cache.DelCache("BoardLists")
			info = info & " <li> ����༭��� ��"
		End if
		If Request.form("upfileid")<>"" Then
			Dim NewID,i
			NewID=Split(Replace(Replace(Request.Form("upfileid")," ",""),"'",""),",")
			For i=0 To Ubound(newid)-1
				team.execute("Update ["&IsForum&"Upfile] Set ID= "&tID&" Where FileID="& newid(i))
			Next
		End if
		team.error1 info & IIF(CID(team.Forum_setting(65))=1,"<li><a href=Thread-"&tID&".html>��������</a><li><a href=Forums-"&fid&".html>������̳</a><meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&"> ","<li><a href=Thread.asp?tid="&tID&">��������</a><meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&"> ")
	End Sub

	Private Sub resaves
		Dim Message,SQL,Rs,Code,ReTopic,Auditing
		Dim ReForumid,fastReTopic,RsCountlist,Pageinfo
		Dim RePage,i
		Call CheckpostTime
		If CID(team.Group_Browse(14)) = 0 Then
			team.Error " �����ڵ���û�лظ����ӵ�Ȩ�ޡ�"
		End If
		Message = team.Checkstr(Trim(Request.Form("message")))
		ReTopic = HTMLEncode(Trim(Request.Form("subject")))
		If strLength(Message)<=cid(team.Forum_setting(64)) Then
			team.error " ���������������� "&team.Forum_setting(64)&" ���ַ��� "
		ElseIf strLength(Message)>cid(team.Forum_setting(67)) Then
			team.error " �����������ܶ��� "&team.Forum_setting(67)&" ���ַ��� "
		End If
		Code = Trim(Request.Form("code"))
		If Cid(team.Forum_setting(48)) >0 Then
			If Cid(Session("postnum"))> Cid(team.Forum_setting(48))  then
				if Not Team.CodeIsTrue(code) Then
					team.error "��֤�������ˢ�º�������д��"
				End If
			End If
		End if
		Dim rCreatactivity,rCreatereward,rRewardpricetype,rNames,rRunMsg
		Set Rs = team.execute("Select forumid,topic,Replies,Creatactivity,Createreward,Rewardpricetype,UserName,RunMsg,ID,CloseTopic From ["&Isforum&"forum] Where ID="& tID)
		If Not Rs.Eof Then
			fID = Rs(0)
			fastReTopic = Rs(1)
			RsCountlist = Rs(2)
			rCreatactivity = Rs(3)
			rCreatereward = Rs(4)
			rRewardpricetype = RS(5)
			rNames = Rs(6)
			rRunMsg = RS(7)
			If CID(Rs(9)) = 1 Then
				team.Error " �������Ѿ��ر�"
			End If
		Else
			team.Error " ����ID���� "
		End if
		Rs.Close:Set Rs=Nothing
		Call ConfigSet
		Call NewUserpostTime()
		If CID(Board_Setting(5)) = 1 Then
			team.Error " ����������˻������ƣ����޷��Դ˰������ӷ����κ����ۻظ���"
		End If 
		If CID(rCreatactivity)=1 Then
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&rNames&"','"&TK_UserName&"','��������̳ϵͳ�Զ����͵�֪ͨ����Ϣ��<BR> ������Ļ��֯ [url=Thread.asp?tid="&tID&"] �� "&fastReTopic&"  ��[/url]���û���������鿴��ϸ���',"&SqlNowString&",'���Ϣ����')")
			team.execute("Update ["&IsForum&"User] Set Newmessage=Newmessage+1 Where UserName='"&rNames&"'")
		End if
		If CID(rCreatereward)=1 and CID(rRewardpricetype)=0 Then
			If Trim(rNames) <> Trim(TK_UserName) Then
				team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&rNames&"','"&TK_UserName&"','��������̳ϵͳ�Զ����͵�֪ͨ����Ϣ��<BR>�����͵����� [url=Thread.asp?tid="&tID&"]�� "&fastReTopic&"  ��[/url]���û���������鿴��ϸ���',"&SqlNowString&",'������Ϣ����')")
				team.execute("Update ["&IsForum&"User] Set Newmessage=Newmessage+1 Where UserName='"&rNames&"'")
			End if
		End If
		If CID(rRunMsg)=1 Then
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&rNames&"','"&TK_UserName&"','��������̳ϵͳ�Զ����͵�֪ͨ����Ϣ��<BR> ����������� [url=Thread.asp?tid="&tID&"] �� "&fastReTopic&"  ��[/url]���û��ظ�����鿴��ϸ���',"&SqlNowString&",'�ظ�����֪ͨ')")
			team.execute("Update ["&IsForum&"User] Set Newmessage=Newmessage+1 Where UserName='"&rNames&"'")
		End If
		Auditing = 0
		If CID(Board_Setting(2))=2 Then
			If CID(team.Group_Browse(16))=0 Or CID(team.Group_Browse(16))=2 Then
				Auditing = 1
			End if
		End if
		team.execute ("insert into ["&Isforum & team.Club_Class(11)&"] (Topicid,UserName,ReTopic,Content,Posttime,Postip,IsNoName,Auditing) values ("&tid&",'"&TK_UserName&"','"&ReTopic&"','"&Message&"',"&SqlNowString&",'"&Remoteaddr&"',"& Cid(Request.Form("isnotname")) &","&Auditing&") ")
		team.execute(" Update ["&Isforum&"forum] Set Lasttime="&SqlNowString&",LastText='"&TK_UserName&"$@$"&Cutstr(HTMLEnCode(Message),150)&"',Replies=Replies+1 Where ID="&tid)
		Call UpdateUserpostExc()
		RePage = Abs(int(-Abs(RsCountlist/team.Forum_setting(20))))
		If int(RePage)>1 Then
			Pageinfo = IIF(CID(team.Forum_setting(65))=1,"-"& RePage,"&Page="& RePage)
		End if
		Dim CacheRs,NextID
		NextID = team.execute("Select Followid From [bbsconfig] Where ID="& fID )(0)
		team.execute("Update ["&IsForum&"Bbsconfig] Set Board_Last='�ظ���<A href="& IIF(CID(team.Forum_setting(65))=1,"Thread-"&tID&".html","Thread.asp?tid="&tID&"") &" target=""_blank"">"&Cutstr(fastReTopic,CutKey)&"</a> $@$"&TK_UserName&"$@$"&Now()&"',Tolrestore=Tolrestore+1,today=today+1 Where ID = "& fID )
		team.execute("update ["&IsForum&"ClubConFig] Set today=today+1,RepostNum=RepostNum+1")
		team.LockCache "RepostNum",Application(CacheName&"_RepostNum")+1
		team.LockCache "TodayNum", Application(CacheName&"_TodayNum") + 1		
		Set CacheRs = Team.Execute("Select Followid,ID From ["&IsForum&"bbsconfig] Where ID="& NextID)
		If Not CacheRs.Eof Then
			If Cid(CacheRs(0)) > 0 Then
				NextID = CacheRs(0)
				team.execute("Update ["&IsForum&"Bbsconfig] Set Board_Last='�ظ���<A href="& IIF(CID(team.Forum_setting(65))=1,"Thread-"&tID&".html","Thread.asp?tid="&tID&"") &" target=""_blank"">RE:"&Cutstr(fastReTopic,CutKey)&"</a>$@$"&TK_UserName&"$@$"&Now()&"',today=today+1 Where ID = "& Cid(CacheRs(1)))
			End If
		End If	
		CacheRs.Close:Set CacheRs=Nothing
		Cache.DelCache("BoardLists")
		If Cid(team.Forum_setting(48)) >0 Then
			Session("postnum") = Session("postnum")+1
		End If 
		If Request.form("upfileid")<>"" Then
			Dim NewID
			NewID=Split(Replace(Replace(Request.Form("upfileid")," ",""),"'",""),",")
			For i=0 To Ubound(newid)-1
				team.execute("Update ["&IsForum&"Upfile] Set ID= "&tID&" Where FileID="& newid(i))
			Next
		End If
		Response.Cookies("posttime") = Now()
		team.error1 IIF(CID(team.Forum_setting(65))=1,"�������ӳɹ�<li><a href=Thread-"& tID & Pageinfo &".html>��������</a><li><a href=Forums-"&fID&".html>������̳</a><meta http-equiv=refresh content=3;url=Thread-"& tID & Pageinfo &".html> ","<li>�������ӳɹ�<li><a href=Thread.asp?tid="& tID & Pageinfo &">��������</a><li><a href=Forums.asp?fid="&fID&">������̳</a><meta http-equiv=refresh content=3;url=Thread.asp?tid="& tID & Pageinfo & "> ")
	End Sub

	Private Sub UpdateUserpostExc()
		'�û����ֲ���
		Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u
		Dim UserPostID,MY_ExtCredits,My_ExtSort
		If Not team.UserLoginED Then  Exit Sub
		ExtCredits = Split(team.Club_Class(21),"|")
		MustOpen = Split(team.Club_Class(22),"|")
		MY_ExtCredits=Split(Board_Setting(14),"|")
		If Request("action") = "saves" Then
			For U=0 to Ubound(ExtCredits)
				ExtSort=Split(ExtCredits(U),",")
				MustSort=Split(MustOpen(U),",")
				My_ExtSort=Split(MY_ExtCredits(U),",")
				If ExtSort(3)=1 Then
					If U = 0 Then
						IF Board_Setting(12) = 1 Then
							UExt = UExt &",Extcredits0=Extcredits0+"&My_ExtSort(0)&""
						Else
							UExt = UExt &",Extcredits0=Extcredits0+"&MustSort(0)&""
						End If
					ElseIf U = 1 Then
						IF Board_Setting(12) = 1 Then
							UExt = UExt &",Extcredits1=Extcredits1+"&My_ExtSort(0)&""
						Else
							UExt = UExt &",Extcredits1=Extcredits1+"&MustSort(0)&""
						End If
					ElseIf U = 2 Then
						IF Board_Setting(12) = 1 Then
							UExt = UExt &",Extcredits2=Extcredits2+"&My_ExtSort(0)&""
						Else
							UExt = UExt &",Extcredits2=Extcredits2+"&MustSort(0)&""
						End If
					Else
						UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"+"&MustSort(0)&""
					End if
				End if
			Next
			team.execute("Update ["&IsForum&"User] Set Posttopic=Posttopic+1"&UExt&" Where ID = "& team.TK_UserID )
		ElseIf Request("action") = "resaves" Then
			For U=0 to Ubound(ExtCredits)
				ExtSort=Split(ExtCredits(U),",")
				MustSort=Split(MustOpen(U),",")
				My_ExtSort=Split(MY_ExtCredits(U),",")
				If ExtSort(3)=1 Then
					If U = 0 Then
						IF Board_Setting(13) = 1 Then
							UExt = UExt &",Extcredits0=Extcredits0+"&My_ExtSort(1)&""
						Else
							UExt = UExt &",Extcredits0=Extcredits0+"&MustSort(1)&""
						End If
					ElseIf U = 1 Then
						IF Board_Setting(13) = 1 Then
							UExt = UExt &",Extcredits1=Extcredits1+"&My_ExtSort(1)&""
						Else
							UExt = UExt &",Extcredits1=Extcredits1+"&MustSort(1)&""
						End If
					ElseIf U = 2 Then
						IF Board_Setting(13) = 1 Then
							UExt = UExt &",Extcredits2=Extcredits2+"&My_ExtSort(1)&""
						Else
							UExt = UExt &",Extcredits2=Extcredits2+"&MustSort(1)&""
						End If
					Else
						UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"+"&MustSort(1)&""
					End if
				End if
			Next
			team.execute("Update ["&IsForum&"User] Set Postrevert=Postrevert+1"&UExt&" Where ID = "& team.TK_UserID )
		End If
		If (team.User_SysTem(5)+Int(team.User_SysTem(6))) Mod 5 < 1 Then UpUserExc()
	End Sub

	Private Sub UpUserExc()
		Dim Rs,NewLevelName,NewClass,NewID
		Set Rs = team.execute("select GroupName,MemberRank,ID from ["&IsForum&"UserGroup] where ID="& team.UserGroupID)
		If Rs.Eof Or Rs.BOF Then
			Set Rs=Nothing : Set Rs = team.execute("select top 1 ID,GroupName,UserColor,UserImg,rank,Members,IsBrowse from ["&IsForum&"UserGroup] where not MemberRank=-1 and MemberRank <="&team.UserGroupExs&" ")
			If Not (Rs.Eof And Rs.Bof) Then
				NewLevelName = Rs(1)&"||"& Rs(2) &"||"& Rs(3) &"||"& Rs(4) & "||" & Split(Rs(6),"|")(21)
				NewClass = Rs(5)
				NewID = Rs(0)
			Else
				Set Rs = team.execute("select top 1 ID,GroupName,UserColor,UserImg,rank,Members,IsBrowse from ["&IsForum&"UserGroup] where not MemberRank=-1 and MemberRank = 0")
				If Not Rs.Eof Then
					NewLevelName = Rs(1)&"||"& Rs(2) &"||"& Rs(3) &"||"& Rs(4) & "||" & Split(Rs(6),"|")(21)
					NewClass = Rs(5)
					NewID = Rs(0)
				End If
			End If
			'������||��ɫ||ͼƬ||����||ǩ��UBB
		Else
			If Rs(1) = -1 Then
				Set Rs = team.execute("select ID,GroupName,UserColor,UserImg,rank,Members,IsBrowse from ["&IsForum&"UserGroup] where MemberRank = -1 and ID="& Rs(2))
				If Not Rs.Eof Then
					NewLevelName = Rs(1)&"||"& Rs(2) &"||"& Rs(3) &"||"& Rs(4) & "||" & Split(Rs(6),"|")(21)
					NewClass = Rs(5)
					NewID = Rs(0)	
				End if
			Else
				Set Rs = team.execute("select top 1 ID,GroupName,UserColor,UserImg,rank,Members,IsBrowse from ["&IsForum&"UserGroup] where not MemberRank=-1 and MemberRank <="&team.UserGroupExs&" ")
				If Not (Rs.Eof And Rs.Bof) Then
					NewLevelName = Rs(1)&"||"& Rs(2) &"||"& Rs(3) &"||"& Rs(4) & "||" & Split(Rs(6),"|")(21)
					NewClass = Rs(5)
					NewID = Rs(0)
				Else
					Set Rs = team.execute("select top 1 ID,GroupName,UserColor,UserImg,rank,Members,IsBrowse from ["&IsForum&"UserGroup] where not MemberRank=-1 and MemberRank = 0")
					If Not Rs.Eof Then
						NewLevelName = Rs(1)&"||"& Rs(2) &"||"& Rs(3) &"||"& Rs(4) & "||" & Split(Rs(6),"|")(21)
						NewClass = Rs(5)
						NewID = Rs(0)
					End If
				End If
			End If
		End If
		team.execute("Update ["&IsForum&"User] Set Members='"&NewClass&"',UserGroupID="&CID(NewID)&",Levelname='"&NewLevelName&"' Where ID="& Int(team.TK_UserID))
		'Session(CacheName&"_UserLogin") = ""
	End Sub

	Private Sub ConfigSet()
		Dim Rs
		Cache.Name = "SaveThreadBoards_"&Fid
		Cache.Reloadtime = Cid(team.Forum_setting(44))
		If Not Cache.ObjIsEmpty() Then
			Boards = Cache.Value
		Else
			Set Rs=team.Execute("Select ID,Followid,bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Board_Key,Board_URL,toltopic,tolrestore,Board_Code,lookperm,postperm,downperm,upperm From ["&IsForum&"Bbsconfig] Where  ID = "& Fid)
			If Rs.Eof Then 
				Team.Error "���ѯ�İ���ID����"
				Exit Sub
			Else
				Boards = Rs.GetRows(-1)
				Cache.Value = Boards
			End If
			RS.Close:Set RS=Nothing
		End If
		If isarray(Boards) Then
			Board_Setting = Split(Boards(3,0),"$$$")
		End If
		team.ChkPost()
		If Boards(1,0) = 0 Then
			team.Error "һ���鲻������"
		End if
		'If Not team.UserLoginED Then 
			'team.Error " �����ڵ���û�з�����Ȩ�ޡ�"
		'End If
		If Not IstrueName(tk_UserName) Then 
			team.Error " �����û����д�����ַ��� "
		End If
		If Request("newpage") = "post" Then
			If Not (Boards(14,0) = ",") Then
				If Instr(Boards(14,0),",") > 0 Then 
					If Not GetUserPower Then team.Error "�����ڵ���û���ڱ��淢��������Ȩ��"
				End If
			End If	
		End if
		If Boards(5,0)<>"" And Not (team.IsMaster Or team.SuperMaster) Then
			If CID(Request.Cookies("Class")("LoginKey"& fid)) = 0 Then
				team.Error "�������½����������ſ��Է�����ظ�����"
			End if
		End If		
	End Sub

	Private Function GetUserPower()
		GetUserPower = False
		Dim B_Postperm,m
		B_Postperm = Split(Boards(14,0),",")
		If Isarray(B_Postperm) Then
			For m = 0 to Ubound(B_Postperm)-1
				If Cid(B_Postperm(m)) = Int(team.UserGroupID) Then GetUserPower = True
			Next 
		End  If
	End Function

	Function UpEditTags(uName,uTime)
		Dim tmp
		tmp = uName
		If Not (team.IsMaster Or team.SuperMaster ) Then
			If team.Forum_setting(95)=1 And DateDiff("s",uTime,Now())> 0 Then
				tmp = tmp & "<p align=right><font color=#000066> " &TK_UserName& " ���༭�� "&Now()&"</font></p>"
			End If
		End if
		UpEditTags = tmp
	End Function

	Function UpPostName(uName)
		Dim Hname,i
		UpPostName = False
		If Trim(uName) = Trim(TK_UserName) Then
			UpPostName=True 
		End if
		If team.Group_Manage(0) = 1 Then
			UpPostName=True 
		End if
	End Function

	Private Sub NewUserpostTime()
		If CID(Board_Setting(9))=1 Then Exit Sub
		If Cid(team.Forum_setting(14))>0 And team.UserLoginED And Not team.ManageUser Then
			If Not IsDate(team.User_SysTem(9)) Then team.User_SysTem(9) = Now()
			If DateDiff("s",CDate(team.User_SysTem(9)),Now()) < Cid(team.Forum_setting(14))*60 Then 
				team.error "��ע���û�����ͣ�� <font color=red> "&team.Forum_setting(14)&" </font> �������ϲſɷ������ӡ�"
			End if
		End If
	End Sub
	Private Sub CheckpostTime()
		If CID(team.Forum_setting(50))<=0 Then
			Exit Sub
		Else
			If IsDate(Request.Cookies("posttime")) Then
				If DateDiff("s",Request.Cookies("posttime"),Now()) <= CLng(team.Forum_setting(50)) Then 
					team.Error "Ϊ��ֹ�����ó����ˮ����̳���Ƶ����û����η�������������"&team.Forum_setting(50)&"�룬������Ҫ�ȴ� "& CLng(team.Forum_setting(50))-DateDiff("s",Request.Cookies("posttime"),Now()) &" ��ſ��Է�����"
				End If
			End If 
		End if
	End Sub
End Class
team.htmlend
%>