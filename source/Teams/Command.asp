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
		team.Error "��������"
End Select
team.footer()

Sub buypost()
	Dim Buyid,PostName,ExtCredits,Rs,SQL,money
	Buyid = HRF(2,2,"buyid")
	Money = HRF(2,2,"money")
	PostName = HRF(2,1,"postname")
	If (Not IsNumeric(Buyid) or Buyid="") or (Not IsNumeric(Money) or Money="") Then 
		Team.Error "��������!"
	Else
		If Int(team.User_SysTem (14+Cid(team.Forum_setting(99)))) < Money Then
			team.error " ��������,�޷���������� ��"
		Else
			Set Rs=Server.CreateObject("Adodb.RecordSet")
			SQL="Select Name From ["&IsForum&"ListRec] Where PostID="& Buyid
			If Not IsObject(Conn) Then ConnectionDatabase
			Rs.Open SQL,Conn,3,2
			If Rs.BOF and Rs.EOF Then
				team.Execute("insert into "&IsForum&"ListRec (PostID,Name) values ("&Buyid&",'"&TK_UserName&",')" )
			Else
				If Instr(Rs(0),TK_UserName&",")>0 Then 
					Team.Error "���Ѿ������˴���,�����ظ�����!"
				Else
					RS(0) = RS(0) & TK_UserName & ","
					Rs.Update
				End If
			End If
			ExtCredits = Split(team.Club_Class(21),"|")
			team.Execute("Update ["&IsForum&"User] set Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"-"&Money&",NewMessage=NewMessage+1  Where UserName='"&TK_UserName&"' ")
			Team.Execute("Update ["&IsForum&"User] set Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"+"&Money -(Money * team.Forum_setting(11) )&",NewMessage=NewMessage+1 Where UserName='"&PostName&"'")
			'����
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&PostName&"','"&TK_UserName&"','�������ӳɹ�,ϵͳ�Զ��۳��㹲��[ "& Money & Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0) &"]Ԫ��֧��������á�',"&SqlNowString&",'������Ϣ֪ͨ',0)")
			'����
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&TK_UserName&"','"&PostName&"','��ϲ�����û�"&tk_UserName&"�ɹ��������㷢�������ӣ��۳�ÿ�ν�����Ҫ֧���Ľ���˰ [ "& (Money * team.Forum_setting(11)) & Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0) &" ]���˴ν�����һ���õ��� [ "&Money -(Money * team.Forum_setting(11))&" "&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" ] �����롣<BR><a href=""Thread.asp?tid="& Buyid &""">�鿴��������</a>',"&SqlNowString&",'������Ϣ֪ͨ',0)")
		End If
		team.error "<li>�������ӳɹ���" & IIF(CID(team.Forum_setting(65))=1,"<li><a href=Thread-"&Buyid&".html>��������</a></li><li><a href=""Default.html"">������̳��ҳ</a></li><meta http-equiv=refresh content=3;url=Thread.asp?tid="&Buyid&"> ","<li><a href=Thread.asp?tid="&Buyid&">��������</a></li><li><a href=""Default.asp"">������̳��ҳ</a></li><meta http-equiv=refresh content=3;url=Thread.asp?tid="&Buyid&"> ")
	End if
End Sub

Sub seebuy()
	Dim Buyid,buyname,o,Uname,rs
	Buyid = HRF(2,2,"buyid")
	Echo "<table border=""0"" width=""80%"" align=""center"" cellspacing=""1"" cellpadding=""3"" class=""a2"">"
	Echo "<tr class=""a1""><td Colspan=""2"">�鿴�Ѿ�������û��б�</td></tr></table><br>"

	If (Not IsNumeric(Buyid) or Buyid="") Then 
		Team.Error("��������!")
	Else
		Echo "<table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""80%"" align=center class=a2><tr class=tab1 align=center><td>�ѹ�������û��б�</td><td>�Ѿ�����</td></tr>"
		Set Rs=Team.Execute("Select Name From ["&Isforum&"ListRec] Where PostID="& int(Buyid) )
		If Not Rs.Eof Then
			Uname = Split(RS(0),",")
			For o=0 To Ubound(Uname)-1
				Echo "<tr class=tab4><td> "&Uname(o)&" </td><td> �� </td></tr>"
			Next
		Else
			Echo "<tr class=tab4><td colspan=2>���޹�����Ա��¼</td></tr>"
		End If
		Echo "</table><BR /><center><input onclick=""history.back(-1)"" type=""submit"" value="" &lt;&lt; �� �� �� һ ҳ "" name=""Submit""> "
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
				Echo "<tr class=""a3""><td Colspan=""2"">�鿴�û�IP</td></tr>"
				Echo "<tr Class=""a4""><td>�û���</td><td>"&RS(0)&"</td></tr>"
				Echo "<tr Class=""a4""><td>ʱ  ��</td><td>"&RS(1)&"</td></tr>"
				Echo "<tr Class=""a4""><td>IP��ַ</td><td>"&RS(2)&" - "&team.address(RS(2))&"</td></tr>"
				Echo "</table><br><input onclick=""history.back(-1)"" type=""submit"" value="" << �� �� �� һ ҳ "" name=""Submit"">"
		End if
		Rs.Close:Set Rs=Nothing
	End If
End Sub

Sub votepoll
	Dim Rs,MyPoll,PollResult,i,WPoll
	If tID = 0 Then
		team.Error "��������"
	Else
		Set Rs = team.execute("Select PollClose,Pollday,PollMax,Polltime,Pollmult,Polltopic,PollResult,PollUser From ["&IsForum&"Fvote] Where RootID="& tID)
		If Rs.Eof And Rs.bof Then
			team.Error "��������"
		Else
			If CID(team.Group_Browse(15)) = 0 Then
				team.Error " �����ڵ���û�з�����Ȩ�ޡ�"
			End if
			If Rs(0) = 1 Then
				team.Error "��ͶƱ�����Ѿ��رա�"
			End If
			If InStr(Rs(7),"$#$")>0 Then
				Dim TestName
				TestName = Split(Rs(7),"$#$")
				For i = 0 To UBound(TestName)
					If tk_UserName = TestName(i) Then team.Error "���Ѿ�Ͷ��Ʊ�ˡ�"
				Next
			Else
				If Rs(7) = tk_UserName Then
					team.Error "���Ѿ�Ͷ��Ʊ�ˡ�"
				End If
			End if
			If Rs(4) = 0 Then
				If Replace(Replace(HRF(1,1,"pollanswers")," ",""),",","")="" Then
					Team.Error("��ЧͶƱ,��ѡ��ͶƱѡ�")
				End if
			End If
			If CID(RS(3)) >0 Then
				If DateDiff("d",RS(3),Date()) > Rs(1) Then
					Team.Error("ͶƱ�Ѿ����ڡ�")
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
			team.Error1 " ͶƱ��ɣ����ڷ����� ��<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		End if
	End if
End Sub


Sub getuseraction
	Dim ho,PName
	PName = team.execute("Select UserName From ["&IsForum&"Forum] Where ID="&tID)(0)
	If Not tk_UserName = PName Then
		team.Error " �����Ƿ�����,�޷�����û� "
	Else
		If Request.form("deleteid") = "" Then
			If Request.form("delsubmit") = "" Then	
				team.Error " ��ѡ����Ҫ��˵��û� "
			Else
				team.Error " ��ѡ����Ҫɾ�����û��ύ"
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
			team.Error1 " ���Ա������ ��<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		Else
			team.Error1 " ���Ա�޳���� ��<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		End if
	End if
End Sub

Sub activityapplylist
	Dim Vs,Rs,tmp
	Set Vs = team.execute("Select PlayName,PlayClass,PlayCity,PlayFrom,Playto,Playplace,PlayCost,PlayGender,PlayNum,PlayStop,PlayUserNum From ["&IsForum&"Activity] Where RootID="& tID ) 
	If Vs.Eof And Vs.Bof Then
		team.Error "��������"
		Exit Sub
	Else
		tmp = Replace(Team.PostHtml (10),"{$paytopic}",Vs(0))
		tmp = Replace(tmp,"{$playclass}",Vs(1))
		tmp = Replace(tmp,"{$playtime}",iif(Vs(4)<>"",VS(3) &" �� " & Vs(4) & " �̶�",Vs(3)))
		tmp = Replace(tmp,"{$playcity}",Vs(2)&" " & Vs(5))
		tmp = Replace(tmp,"{$playmoney}",Vs(6)&"")
		tmp = Replace(tmp,"{$playsex}",iif(Vs(7)=0,"����",iif(Vs(7)=1,"����","Ů��")))
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
			Echo "<tr class=""tab1""><td><input type=""checkbox"" name=""chkall"" onClick=""checkall(this.form)"" class=""radio""></td><td width=""20%"">�������</td><td width=""20%""> ���� </td><td  width=""20%""> ÿ�˻��� </td><td width=""20%""> ����ʱ�� </td><td width=""10%""> ״̬ </td></tr>"
			Set Rs =  team.execute("Select PlayUser,Playtext,PlayClass,PlayBy,playBysomach,PlayTime,ID From ["&IsForum&"ActivityUser] Where RootID="& tID )
			Do While Not Rs.Eof
				Echo "<tr class=""tab3""><td><Input Name=""newid"" type=""hidden"" value="&RS(6)&"><input type=""checkbox"" name=""deleteid"" value="&RS(6)&" class=""radio""></td><td> "&RS(0)&" </td><td> "&RS(1)&" </td><td> "&IIF(RS(4)=0,"�Ը�",RS(4)& " Ԫ")&" </td><td> "&RS(5)&" </td><td> "&IIF(RS(2)=0,"��δ���","�����")&" </td></tr> "
				Rs.Movenext
			Loop
			Rs.Close:Set Rs=Nothing
			Echo "</table><br><center><input type=""submit"" name=""forumlinksubmit"" value=""�� ��""> <input type=""submit"" name=""delsubmit"" value=""ɾ���û��ύ""></form></center>"
		End If
	End If
	Vs.Close:Set Vs=Nothing
End Sub


Sub activityapplies
	Dim Rs
	If tID = 0 Then
		team.Error "��������"
	Else
		Set Rs =  team.execute("Select PlayUser From ["&IsForum&"ActivityUser] Where RootID="& tID &" and PlayUser='"&tk_UserName&"'")
		If Not (Rs.Eof And Rs.Bof) Then
			team.Error "���Ѿ�������ˣ������ظ��ύ��<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"
		Else
			team.execute ("insert into ["&IsForum&"ActivityUser] (Rootid,PlayUser,Playtext,PlayClass,PlayBy,playBysomach,PlayTime) values ("&tID&",'"&tk_UserName&"','"&HRF(1,1,"playmessage")&"',0,"&CID(HRF(1,2,"payment"))&","&CID(HRF(1,2,"payvalue"))&","&SqlNowString&") ")
			team.execute ("Update ["&IsForum&"Activity] Set PlayUserNum=PlayUserNum+1 Where RootID="& tID)
			team.Error1 "���������Ѿ���¼����ȴ���ˡ�<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&">"

			Dim PName
			PName = team.execute("Select UserName From ["&IsForum&"Forum] Where ID="&tID)(0)
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&PName&"','"&TK_UserName&"','��������̳ϵͳ�Զ����͵�֪ͨ����Ϣ��<BR> ������Ļ��֯ [url=Thread.asp?tid="&tID&"] �� "&fastReTopic&"  ��[/url]���û�������룬[url=Command.asp?action=activityapplylist&tid="&tID&"]��鿴��ϸ���[/url]',"&SqlNowString&",'���Ϣ����')")
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
				.Write "<tr class=a3><td Colspan=2>�鿴�û�IP</td></tr>"
				.Write "<tr Class=a4><td>�û���</td><td>"&RS(0)&"</td></tr>"
				.Write "<tr Class=a4><td>ʱ  ��</td><td>"&RS(1)&"</td></tr>"
				.Write "<tr Class=a4><td>IP��ַ</td><td>"&RS(2)&" - "&team.address(RS(2))&"</td></tr>"
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
		team.error " ��������"
	Else
		Set Rs=team.execute("Select Board_Setting From ["&IsForum&"Bbsconfig] Where ID="& fID)
		If Rs.Eof Then
			team.Error " ����ѯ�İ�����"
		Else
			Dim Board_Setting
			Board_Setting = Split(Rs(0),"$$$")
			x1 = "���ֲ���˵��"
			Echo team.MenuTitle
			Echo "<table cellspacing=""1"" cellpadding=""3"" width=""98%"" align=""center"" class=""a2"">"
			Echo "<tr class=""a1"">"
			Echo "	<td colspan=""10"">���ֲ���˵��</td>"
			Echo "</tr>"
			Echo "<tr align=""center"" class=""tab4"">"
			Echo "	<td>���ִ���</td>"
			Echo "	<td>������(+)</td>"
			Echo "	<td>�ظ�(+)</td>"
			Echo "	<td>�Ӿ���(+)</td>"
			Echo "	<td>�ϴ�����(+)</td>"
			Echo "	<td>���ظ���(-)</td>"
			Echo "	<td>������Ϣ(-)</td>"
			Echo "	<td>����(-)</td>"
			Echo "	<td>�����ƹ�(+)</td>"
			Echo "	<td>���ֲ�������</td>"
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
		team.error " ��������"
	Else
		ExtCredits = Split(team.Club_Class(21),"|")
		Set Rs = team.execute("Select UserName,Topic,Rewardprice,Posttime,Rewardpricetype,ReList From ["&Isforum&"forum] Where ID="& tID)
		If Not Rs.Eof Then
			If Rs(4) = 1 Then
				team.Error "�����Ѿ���������״̬"
			Else
				team.execute("Update ["&Isforum&"forum] Set Rewardpricetype=1 Where ID="& tID)
				team.execute("Update ["&Isforum & RS(5) &"] Set Reward=1 Where ID="& rID)
				'�������û��۷�
				team.execute("Update ["&Isforum&"User] Set Extcredits"&team.Forum_setting(99)&"=Extcredits"&team.Forum_setting(99)&"-"&RS(2)&" Where UserName='"& Rs(0) &"'")
				'����ȷ�û��ӷ�
				Set Rs1 = team.execute("Select UserName From ["&IsForum & Rs(5) &"] Where ID="& rID )
				If Not Rs.Eof Then
					team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&Rs1(0)&"','"&TK_UserName&"','��ϲ�����Ļظ��� "&Rs(0)&" ѡΪ������Ѵ𰸡�[BR] ������������:  [url=Thread.asp?tid="&tID&"] �� "&RS(1)&"  ��[/url] [BR] �����õ��� "&RS(2)&" ��� "&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" [BR] ������Ա����������飬��������ȡ����ϵ��',"&SqlNowString&",'������Ļظ���ѡΪ��Ѵ�')")
					team.execute("Update ["&IsForum&"User] Set Newmessage=Newmessage+1,Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"+"&CID(RS(2))&" Where UserName='"&Rs1(0)&"'")
				End If
			End if
		End If
		team.error1 "<li>������Ѵ𰸳ɹ����������ֶ� <a href=Thread.asp?tid="&tID&">��������</a> ����ȴ�ϵͳ�Զ��������⡣<meta http-equiv=refresh content=3;url=Thread.asp?tid="&tID&"> "
	End If
End Sub

%>