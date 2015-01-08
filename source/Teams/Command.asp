<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim tID,fID,rID,x1,x2
fID = HRF(2,2,"fid")
tID = HRF(2,2,"tid")
rID = HRF(2,2,"rid")
team.ChkPost()
Call TestUser()
team.Headers(Team.Club_Class(1))
Select Case HRF(2,1,"action")
	Case "bestanswer"
		Call bestanswer
	Case "credits"
		Call credits
	Case "activityapplies"
		Call activityapplies
	Case "activityapplylist"
		Call activityapplylist
	Case "getuseraction"
		Call getuseraction
	Case "votepoll"
		Call votepoll
	Case "lookip"
		Call LookIPs
	Case "seebuy"
		Call Seebuy
	Case "buypost"
		Call buypost()
	Case Else
		team.Error "参数错误"
End Select
team.footer()

Sub buypost()
	Dim Buyid,PostName,ExtCredits,Rs,SQL,money
	Buyid = HRF(2,2,"buyid")
	Money = HRF(2,2,"money")
	PostName = HRF(2,1,"postname")
	If (Not IsNumeric(Buyid) or Buyid="") or (Not IsNumeric(Money) or Money="") Then 
		Team.Error "参数错误!"
	Else
		If Int(team.User_SysTem (14+Cid(team.Forum_setting(99)))) < Money Then
			team.error " 您的余额不够,无法购买此帖子 。"
		Else
			Set Rs=Server.CreateObject("Adodb.RecordSet")
			SQL="Select Name From ["&IsForum&"ListRec] Where PostID="& Buyid
			If Not IsObject(Conn) Then ConnectionDatabase
			Rs.Open SQL,Conn,3,2
			If Rs.BOF and Rs.EOF Then
				team.Execute("insert into "&IsForum&"ListRec (PostID,Name) values ("&Buyid&",'"&TK_UserName&",')" )
			Else
				If Instr(Rs(0),TK_UserName&",")>0 Then 
					Team.Error "你已经购买了此帖,无需重复购买!"
				Else
					RS(0) = RS(0) & TK_UserName & ","
					Rs.Update
				End If
			End If
			ExtCredits = Split(team.Club_Class(21),"|")
			team.Execute("Update ["&IsForum&"User] set Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"-"&Money&",NewMessage=NewMessage+1  Where UserName='"&TK_UserName&"' ")
			Team.Execute("Update ["&IsForum&"User] set Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"+"&Money -(Money * team.Forum_setting(11) )&",NewMessage=NewMessage+1 Where UserName='"&PostName&"'")
			'短信
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&PostName&"','"&TK_UserName&"','购买帖子成功,系统自动扣除你共计[ "& Money & Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0) &"]元以支付购买费用。',"&SqlNowString&",'交易信息通知',0)")
			'短信
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&TK_UserName&"','"&PostName&"','恭喜您，用户"&tk_UserName&"成功购买了你发布的帖子，扣除每次交易需要支付的交易税 [ "& (Money * team.Forum_setting(11)) & Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0) &" ]，此次交易您一共得到了 [ "&Money -(Money * team.Forum_setting(11))&" "&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" ] 的收入。<BR><a href=""Thread.asp?tid="& Buyid &""">查看交易帖子</a>',"&SqlNowString&",'交易信息通知',0)")
		End If
		team.error "<li>购买帖子成功。" & IIF(CID(team.Forum_setting(65))=1,"<li><a href=Thread-"&Buyid&".html>返回主题</a></li><li><a href=""Default.html"">返回论坛首页</a></li><meta http-equiv=refresh content=3;url=Thread.asp?tid="&Buyid&"> ","<li><a href=Thread.asp?tid="&Buyid&">返回主题</a></li><li><a href=""Default.asp"">返回论坛首页</a></li><meta http-equiv=refresh content=3;url=Thread.asp?tid="&Buyid&"> ")
	End if
End Sub

Sub seebuy()
	Dim Buyid,buyname,o,Uname,rs
	Buyid = HRF(2,2,"buyid")
	Echo "<table border=""0"" width=""80%"" align=""center"" cellspacing=""1"" cellpadding=""3"" class=""a2"">"
	Echo "<tr class=""a1""><td Colspan=""2"">查看已经购买的用户列表</td></tr></table><br>"

	If (Not IsNumeric(Buyid) or Buyid="") Then 
		Team.Error("参数错误!")
	Else
		Echo "<table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""80%"" align=center class=a2><tr class=tab1 align=center><td>已购买此帖用户列表</td><td>已经购买</td></tr>"
		Set Rs=Team.Execute("Select Name From ["&Isforum&"ListRec] Where PostID="& int(Buyid) )
		If Not Rs.Eof Then
			Uname = Split(RS(0),",")
			For o=0 To Ubound(Uname)-1
				Echo "<tr class=tab4><td> "&Uname(o)&" </td><td> √ </td></tr>"
			Next
		Else
			Echo "<tr class=tab4><td colspan=2>暂无购买人员记录</td></tr>"
		End If
		Echo "</table><BR /><center><input onclick=""history.back(-1)"" type=""submit"" value="" &lt;&lt; 返 回 上 一 页 "" name=""Submit""> "
	End if
End sub


Sub LookIPs
	if team.Group_Manage(10) = 1  then 
		Dim Rs,ReList,SQL
		If rID = 0 Then
			SQL = "Select UserName,Posttime,postip,ReList From ["&IsForum&"forum] Where ID="& tID
		Else
			ReList = team.execute("Select ReList From ["&IsForum&"forum] Where ID="& tID)(0)
			SQL = "Select UserName,Posttime,Postip From ["&IsForum & ReList &"] Where ID="& rID
		End If
		Set Rs = team.execute(SQL)
		If Not (Rs.Eof And Rs.Bof) Then
				Echo "<table border=""0"" width=""98%"" align=""center"" cellspacing=""1"" cellpadding=""3"" class=""a2"">"
				Echo "<tr class=""a3""><td Colspan=""2"">查看用户IP</td></tr>"
				Echo "<tr Class=""a4""><td>用户名</td><td>"&RS(0)&"</td></tr>"
				Echo "<tr Class=""a4""><td>时  间</td><td>"&RS(1)&"</td></tr>"
				Echo "<tr Class=""a4""><td>IP地址</td><td>"&RS(2)&" - "&team.address(RS(2))&"</td></tr>"
				Echo "</table><br><input onclick=""history.back(-1)"" type=""submit"" value="" << 返 回 上 一 页 "" name=""Submit"">"
		End if
		Rs.Close:Set Rs=Nothing
	End If
End Sub

Sub votepoll
	Dim Rs,MyPoll,PollResult,i,WPoll
	If tID = 0 Then
		team.Error "参数错误"
	Else
		Set Rs = team.execute("Select PollClose,Pollday,PollMax,Polltime,Pollmult,Polltopic,PollResult,PollUser From ["&IsForum&"Fvote] Where RootID="& tID)
		If Rs.Eof And Rs.bof Then
			team.Error "参数错误"
		Else
			If CID(team.Group_Browse(15)) = 0 Then
				team.Error " 您所在的组没有发帖的权限。"
			End if
			If Rs(0) = 1 Then
				team.Error "此投票主题已经关闭。"
			End If
			If InStr(Rs(7),"$#$")>0 Then
				Dim TestName
				TestName = Split(Rs(7),"$#$")
				For i = 0 To UBound(TestName)
					If tk_UserName = TestName(i) Then team.Error "您已经投过票了。"
				Next
			Else
				If Rs(7) = tk_UserName Then
					team.Error "您已经投过票了。"
				End If
			End if
			If Rs(4) = 0 Then
				If Replace(Replace(HRF(1,1,"pollanswers")," ",""),",","")="" Then
					Team.Error("无效投票,请选择投票选项。")
				End if
			End If
			If CID(RS(3)) >0 Then
				If DateDiff("d",RS(3),Date()) > Rs(1) Then
					Team.Error("投票已经过期。")
				End If 
			End if
			MyPoll = Split(Rs(5),"|")
			WPoll = Split(Rs(6),"|")
			For i=0 To Ubound(MyPoll)
				If Rs(4) = 0 Then
					If PollResult = "" Then
						If i = CID(HRF(1,1,"pollanswers")) Then
							PollResult = PollResult & WPoll(i) + 1
						Else
							PollResult = PollResult & WPoll(i)
						End if
					Else
						If i = CID(HRF(1,1,"pollanswers")) Then
							PollResult = PollResult & "|" & WPoll(i) + 1
						Else
							PollResult = PollResult & "|" & WPoll(i) 
						End if
					End If
				Else
					If PollResult = "" Then
						PollResult = PollResult & WPoll(i) + CID(HRF(1,1,"pollanswers"&i))
					Else
						PollResult = PollResult & "|" & WPoll(i) + CID(HRF(1,1,"pollanswers"&i))
					End If
				End if
			Next
			Dim GetName
			If Rs(7) &"" = "" Then
				GetName = tk_UserName
			Else
				GetName = Rs(7) & "$#$" &tk_UserName
			End if
			team.execute ("Update ["&IsForum&"Fvote] Set PollResult='"&PollResult&"',PollUser='"&GetName&"' Where RootID="& tID)
			team.Error1 " 投票完成，现在返回中 。<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		End if
	End if
End Sub


Sub getuseraction
	Dim ho,PName
	PName = team.execute("Select UserName From ["&IsForum&"Forum] Where ID="&tID)(0)
	If Not tk_UserName = PName Then
		team.Error " 您不是发起人,无法审核用户 "
	Else
		If Request.form("deleteid") = "" Then
			If Request.form("delsubmit") = "" Then	
				team.Error " 请选定需要审核的用户 "
			Else
				team.Error " 请选定需要删除的用户提交"
			End if
		Else
			for each ho in request.form("deleteid")
				If Request.form("delsubmit") = "" Then
					team.execute("Update ["&Isforum&"ActivityUser] Set PlayClass=1 Where ID="&Int(ho))
				Else
					team.execute("Delete From ["&Isforum&"ActivityUser] Where ID="&Int(ho))
				End if
			next
		End If
		If Request.form("delsubmit") = "" Then
			team.Error1 " 活动人员审核完成 。<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		Else
			team.Error1 " 活动人员剔除完成 。<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		End if
	End if
End Sub

Sub activityapplylist
	Dim Vs,Rs,tmp
	Set Vs = team.execute("Select PlayName,PlayClass,PlayCity,PlayFrom,Playto,Playplace,PlayCost,PlayGender,PlayNum,PlayStop,PlayUserNum From ["&IsForum&"Activity] Where RootID="& tID ) 
	If Vs.Eof And Vs.Bof Then
		team.Error "参数错误"
		Exit Sub
	Else
		tmp = Replace(Team.PostHtml (10),"{$paytopic}",Vs(0))
		tmp = Replace(tmp,"{$playclass}",Vs(1))
		tmp = Replace(tmp,"{$playtime}",iif(Vs(4)<>"",VS(3) &" 至 " & Vs(4) & " 商定",Vs(3)))
		tmp = Replace(tmp,"{$playcity}",Vs(2)&" " & Vs(5))
		tmp = Replace(tmp,"{$playmoney}",Vs(6)&"")
		tmp = Replace(tmp,"{$playsex}",iif(Vs(7)=0,"不限",iif(Vs(7)=1,"男性","女性")))
		tmp = Replace(tmp,"{$playnum}",Vs(8))
		tmp = Replace(tmp,"{$playaction}",Vs(10))
		tmp = Replace(tmp,"{$playclosetime}",Vs(9))
		tmp = Replace(tmp,"{$msgs}","Display:None")
		tmp = Replace(tmp,"{$myinfos}","")
		tmp = Replace(tmp,"{$disabled}","disabled")
		Echo tmp
		Dim PName
		PName = team.execute("Select UserName From ["&IsForum&"Forum] Where ID="&tID)(0)
		If Vs(10) > 0 And tk_UserName = PName Then
			Echo "<form method=""post"" action=""?action=getuseraction&tid="&tid&"""><table cellspacing=""1"" cellpadding=""3"" align=""center"" width=""98%"" border=""0"" class=""a2"">"
			Echo "<tr class=""tab1""><td><input type=""checkbox"" name=""chkall"" onClick=""checkall(this.form)"" class=""radio""></td><td width=""20%"">活动申请者</td><td width=""20%""> 留言 </td><td  width=""20%""> 每人花销 </td><td width=""20%""> 申请时间 </td><td width=""10%""> 状态 </td></tr>"
			Set Rs =  team.execute("Select PlayUser,Playtext,PlayClass,PlayBy,playBysomach,PlayTime,ID From ["&IsForum&"ActivityUser] Where RootID="& tID )
			Do While Not Rs.Eof
				Echo "<tr class=""tab3""><td><Input Name=""newid"" type=""hidden"" value="&RS(6)&"><input type=""checkbox"" name=""deleteid"" value="&RS(6)&" class=""radio""></td><td> "&RS(0)&" </td><td> "&RS(1)&" </td><td> "&IIF(RS(4)=0,"自付",RS(4)& " 元")&" </td><td> "&RS(5)&" </td><td> "&IIF(RS(2)=0,"尚未审核","已审核")&" </td></tr> "
				Rs.Movenext
			Loop
			Rs.Close:Set Rs=Nothing
			Echo "</table><br><center><input type=""submit"" name=""forumlinksubmit"" value=""提 交""> <input type=""submit"" name=""delsubmit"" value=""删除用户提交""></form></center>"
		End If
	End If
	Vs.Close:Set Vs=Nothing
End Sub


Sub activityapplies
	Dim Rs
	If tID = 0 Then
		team.Error "参数错误"
	Else
		Set Rs =  team.execute("Select PlayUser From ["&IsForum&"ActivityUser] Where RootID="& tID &" and PlayUser='"&tk_UserName&"'")
		If Not (Rs.Eof And Rs.Bof) Then
			team.Error "您已经申请过了，无需重复提交。<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		Else
			team.execute ("insert into ["&IsForum&"ActivityUser] (Rootid,PlayUser,Playtext,PlayClass,PlayBy,playBysomach,PlayTime) values ("&tID&",'"&tk_UserName&"','"&HRF(1,1,"playmessage")&"',0,"&CID(HRF(1,2,"payment"))&","&CID(HRF(1,2,"payvalue"))&","&SqlNowString&") ")
			team.execute ("Update ["&IsForum&"Activity] Set PlayUserNum=PlayUserNum+1 Where RootID="& tID)
			team.Error1 "您的申请已经记录，请等待审核。<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"

			Dim PName
			PName = team.execute("Select UserName From ["&IsForum&"Forum] Where ID="&tID)(0)
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&PName&"','"&TK_UserName&"','这是由论坛系统自动发送的通知短消息。<BR> 您发起的活动组织 [url=Thread.asp?tid="&tID&"] 『 "&fastReTopic&"  』[/url]有用户申请加入，[url=Command.asp?action=activityapplylist&tid="&tID&"]请查看详细情况[/url]',"&SqlNowString&",'活动消息回馈')")
			team.execute("Update ["&IsForum&"User] Set Newmessage=Newmessage+1 Where UserName='"&Rs(6)&"'")

		End if
	End if
End Sub

Sub LookIP
	if Team.opsetup(15)=1  then 
		Dim SQL,ReList
		If isnumeric(Request("retopicid")) Then
			ReList=team.execute("Select ReList from [forum] where id="&Int(Request("id")))(0)
			SQL="Select username,Posttime,Postip From ["&ReList&"] where id="&Request("retopicid")
		Else
			SQL="Select username,Posttime,Postip From [forum] where id="&request("id")
		End if
		Set Rs=Team.Execute(SQL)
		If Not Rs.Eof Then
			With Response
				.Write "<table border=0 width=90% align=center cellspacing=1 cellpadding=3 class=a2>"
				.Write "<tr class=a3><td Colspan=2>查看用户IP</td></tr>"
				.Write "<tr Class=a4><td>用户名</td><td>"&RS(0)&"</td></tr>"
				.Write "<tr Class=a4><td>时  间</td><td>"&RS(1)&"</td></tr>"
				.Write "<tr Class=a4><td>IP地址</td><td>"&RS(2)&" - "&team.address(RS(2))&"</td></tr>"
				.Write "<tr class=a3><td Colspan=2 align=center><a href=ShowPost.asp?id="&Request("id")&">BACK</a></td></tr>"
				.Write "</table>"
			End With
		End If
		Rs.Close:Set Rs=Nothing
	End If
End Sub

Sub credits
	Dim Rs,MustOpen,M,MustSort,ExtCredits,ExtSort
	If fID = 0 Then
		team.error " 参数错误。"
	Else
		Set Rs=team.execute("Select Board_Setting From ["&IsForum&"Bbsconfig] Where ID="& fID)
		If Rs.Eof Then
			team.Error " 您查询的版块错误。"
		Else
			Dim Board_Setting
			Board_Setting = Split(Rs(0),"$$$")
			x1 = "积分策略说明"
			Echo team.MenuTitle
			Echo "<table cellspacing=""1"" cellpadding=""3"" width=""98%"" align=""center"" class=""a2"">"
			Echo "<tr class=""a1"">"
			Echo "	<td colspan=""10"">积分策略说明</td>"
			Echo "</tr>"
			Echo "<tr align=""center"" class=""tab4"">"
			Echo "	<td>积分代号</td>"
			Echo "	<td>发主题(+)</td>"
			Echo "	<td>回复(+)</td>"
			Echo "	<td>加精华(+)</td>"
			Echo "	<td>上传附件(+)</td>"
			Echo "	<td>下载附件(-)</td>"
			Echo "	<td>发短消息(-)</td>"
			Echo "	<td>搜索(-)</td>"
			Echo "	<td>访问推广(+)</td>"
			Echo "	<td>积分策略下限</td>"
			Echo "</tr>"
		End If
		Dim MY_ExtCredits,My_ExtSort
		ExtCredits= Split(team.Club_Class(21),"|")
		MustOpen = Split(team.Club_Class(22),"|")
		MY_ExtCredits=Split(Board_Setting(14),"|")
		For M=0 to Ubound(MustOpen)
			ExtSort=Split(ExtCredits(M),",")
			MustSort=Split(MustOpen(M),",")	
			My_ExtSort=Split(MY_ExtCredits(M),",")
			If Split(ExtCredits(M),",")(3)=1 Then
				Echo "<tr align=""center"">"
				Echo "	<td bgcolor=""#F8F8F8"">"&ExtSort(0)&"</td>"
				Echo "	<td bgcolor=""#FFFFFF"">"&IIF(Board_Setting(12) = 1,My_ExtSort(0),MustSort(0))&"</td>"
				Echo "	<td bgcolor=""#F8F8F8"">"&IIF(Board_Setting(13) = 1,My_ExtSort(1),MustSort(1))&"</td>"
				Echo "	<td bgcolor=""#FFFFFF"">"&IIF(Board_Setting(3) = 1,My_ExtSort(2),MustSort(2))&"</td>"
				Echo "	<td bgcolor=""#F8F8F8"">"&MustSort(3)&"</td>"
				Echo "	<td bgcolor=""#FFFFFF"">"&MustSort(4)&"</td>"
				Echo "	<td bgcolor=""#F8F8F8"">"&MustSort(5)&"</td>"
				Echo "	<td bgcolor=""#FFFFFF"">"&MustSort(6)&"</td>"
				Echo "	<td bgcolor=""#F8F8F8"">"&MustSort(7)&"</td>"
				Echo "	<td bgcolor=""#FFFFFF"">"&MustSort(8)&"</td>"
				Echo "</tr> "
			End If
		Next
		Echo "</table>"
	End if
End Sub

Sub bestanswer
	Dim Rs,Rs1,ExtCredits
	If tID = 0 Then
		team.error " 参数错误。"
	Else
		ExtCredits = Split(team.Club_Class(21),"|")
		Set Rs = team.execute("Select UserName,Topic,Rewardprice,Posttime,Rewardpricetype,ReList From ["&Isforum&"forum] Where ID="& tID)
		If Not Rs.Eof Then
			If Rs(4) = 1 Then
				team.Error "本帖已经结束悬赏状态"
			Else
				team.execute("Update ["&Isforum&"forum] Set Rewardpricetype=1 Where ID="& tID)
				team.execute("Update ["&Isforum & RS(5) &"] Set Reward=1 Where ID="& rID)
				'从悬赏用户扣分
				team.execute("Update ["&Isforum&"User] Set Extcredits"&team.Forum_setting(99)&"=Extcredits"&team.Forum_setting(99)&"-"&RS(2)&" Where UserName='"& Rs(0) &"'")
				'给正确用户加分
				Set Rs1 = team.execute("Select UserName From ["&IsForum & Rs(5) &"] Where ID="& rID )
				If Not Rs.Eof Then
					team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&Rs1(0)&"','"&TK_UserName&"','恭喜，您的回复被 "&Rs(0)&" 选为悬赏最佳答案。[BR] 悬赏帖子链接:  [url=Thread.asp?tid="&tID&"] 『 "&RS(1)&"  』[/url] [BR] ，您得到了 "&RS(2)&" 点的 "&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" [BR] 如果您对本操作有异议，请与作者取得联系。',"&SqlNowString&",'您发表的回复被选为最佳答案')")
					team.execute("Update ["&IsForum&"User] Set Newmessage=Newmessage+1,Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"+"&CID(RS(2))&" Where UserName='"&Rs1(0)&"'")
				End If
			End if
		End If
		team.error1 "<li>设置最佳答案成功，您可以手动 <a href=Thread.asp?tid="&tID&">返回主题</a> ，或等待系统自动返回主题。<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&"> "
	End If
End Sub

%>