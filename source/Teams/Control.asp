<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
Dim x1,x2,fID
team.Headers(Team.Club_Class(1) &" - �������")
testUser()
Select Case Request("action")
	Case "usercp"
		Call UserCp
	Case "edituserinfo"
		Call Edituserinfo
	Case "bank"
		Call UserBank
	Case "edituserbank"
		Call edituserbank
	Case "friend"
		Call UserFriend
	Case "edituserfriend"
		Call edituserfriend
	Case "delfriend"
		Call delfriend
	Case "buyuserbank"
		Call buyuserbank
	Case "subscription"
		Call Subscription
	Case "subscriptionok"
		Call Subscriptionok
	Case "subscriptionshow"
		Call Subscriptionshow
	Case Else
		Call Main()
End Select
team.Footer()


Sub Subscriptionshow
	Dim tid,rs
	tID = HRF(2,2,"tid")
	If tid = 0 Then
		team.Error "�����ڴ�����"
	Else
		Set Rs = team.execute("select Url from ["&Isforum&"Favorites] where id="& tid)
		If Rs.eof And Rs.Bof Then
			team.Error "��û���ղش�����"
		Else
			team.execute("update ["&Isforum&"Favorites] set look=look+1 where id="& tid)
			Response.Redirect Rs(0)
		End If 
	End if
End Sub

Sub Subscriptionok
	Dim ho
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"Favorites] Where ID="&Int (ho))
	Next
	team.Error1 "ѡ�����ղ��Ѿ�ɾ��.��ȴ�ϵͳ����.<meta http-equiv=refresh content=3;url="""& Request.ServerVariables("http_referer") &""">"
End sub

Sub Subscription
	Dim tmp,Rs,Ms,Ump,UserInfo,UserMedals,Emp
	x2 = ""
	x1 = " <a href=""Control.asp""> ���ĵ����� </a> "
	tmp = Replace(Team.UserHtml (1),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"favorites"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"uersinfo"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"userbank"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"friends"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"usercp"))
	Set Rs = team.execute("Select UserGroupID,Levelname,Usermail,Userhome,Userface,UserCity,UserSex,Question,Answer,Honor,Birthday,Sign,Medals,UserInfo,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members,Friend From ["&Isforum&"User] where ID="& team.TK_UserID)
	If Not Rs.Eof Then
		Ms = Rs.GetString(,1, "$$##$$","","")
	End if
	Rs.Close:Set Rs=Nothing
	Ump = Split(Ms,"$$##$$")
	UserInfo = Split(Ump(13),"|")		
	'UserInfo = QQ|ICQ|YAHOO|MSN|TAOBAO|ALIPAY
	tmp = Replace(tmp,"{$usermail}",Ump(2))
	tmp = Replace(tmp,"{$userhome}",Ump(3))
	tmp = Replace(tmp,"{$userface}",Fixjs(Ump(4)))
	tmp = Replace(tmp,"{$sign}",Sign_Code(Ump(11),CID(Split(Ump(1),"||")(4))))
	tmp = Replace(tmp,"{$userqq}",IIf(UserInfo(0)<>"","<a target=""_blank"" href=""tencent://message/?uin="&UserInfo(0)&"&Site=team5.cn&Menu=yes""><img border=""0"" SRC=""http://wpa.qq.com/pa?p=1:"&UserInfo(0)&":7"" alt=""���������ҷ���Ϣ"" onerror=""javascript:this.src='images/qqerr.gif'""></a>",""))
	tmp = Replace(tmp,"{$qq}",UserInfo(0))
	tmp = Replace(tmp,"{$icq}",UserInfo(1))
	tmp = Replace(tmp,"{$yahoo}",UserInfo(2))
	tmp = Replace(tmp,"{$msn}",UserInfo(3))
	tmp = Replace(tmp,"{$taobao}",IIF(UserInfo(4)<>"","<script type=""text/javascript"">document.write('<a target=""_blank"" href=""http://amos1.taobao.com/msg.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2""><img src=""http://amos1.taobao.com/online.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2"" alt=""�Ա�����"" border=""0"" />"&UserInfo(4)&"</a> ');</script>",""))
	tmp = Replace(tmp,"{$alipay}",UserInfo(5))
	If Ump(12)<>"" Then
		UserMedals = ""
		If Instr(Ump(12),"$$$")>0 Then
			Dim i
			UserMedals = Split(Ump(12),"$$$")
			For i = 0 to Ubound(UserMedals)-1
				Emp = Emp & "<img src=""images/plus/"&Split(UserMedals(i),"&&&")(0)&""" align=""absmiddle"" alt="""&Split(UserMedals(i),"&&&")(1)&"""> "
			Next
			tmp = Replace(tmp,"{$userMedals}",Emp)
		End if
	Else
		tmp = Replace(tmp,"{$userMedals}","")
	End If
	Dim SQL,DispCount,Maxpage,PageNum,Page,iRs
	DispCount = team.execute("Select Count(ID) from ["&Isforum&"Favorites] Where UserName='"& tk_UserName &"'")(0)
	SQL = "Select id,name,url,addtime,tag,ispub,look from ["&Isforum&"Favorites] Where UserName='"& tk_UserName &"'"
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum = SqlQueryNum+1
		Maxpage = Cid(team.Forum_setting(19))		'ÿҳ��ҳ��
		PageNum = Abs(int(-Abs(DispCount/Maxpage)))	'ҳ��
		Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'��ǰҳ
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		iRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
 	If Not Isarray(iRs) Then
		emp = "<tr class=""a4""><td align=""center"" colspan=""4"">��û���ղ��κ�����</td></tr>"
	Else
		For i=0 To Ubound(iRs,2)
			emp = emp & "<tr class=""a4""><td align=""center""><input class=""checkbox"" type=""checkbox"" name=""deleteid"" value="""& irs(0,i) &"""></td><td width=""45%"" align=""left""><a href=""?action=subscriptionshow&tid="& irs(0,i) &""">"& irs(1,i) &"</td><td align=""center"">"& irs(6,i) &"</td><td align=""center"">"& irs(3,i) &"</td></tr>" & Vbcrlf
		Next
	End If
	tmp = Replace(tmp,"{$favishow}",emp)
	tmp=Replace(tmp,"{$pagecount}",PageNum)
	tmp=Replace(tmp,"{$dispcount}",DispCount)
	Echo tmp
End Sub

Sub edituserfriend
	Dim newfriend,myFriend,Rs
	Newfriend = HtmlEncode(Request("newfriend"))
	myFriend = ""
	If team.execute("Select UserName From ["&Isforum&"User] Where UserName='"&Newfriend&"'").Eof Then
		team.error " ϵͳ������ "&Newfriend&" ���û���"
	Elseif Trim(TK_UserName) = Trim(Newfriend) Then
		team.error "����������Լ�Ϊ���ѡ�"
	Else
		Set Rs = team.execute("Select Friend From ["&Isforum&"User] where ID="& team.TK_UserID)
		If Not Rs.Eof Then
			myFriend = RS(0) & Newfriend & "|"
		End if
		Rs.Close:Set Rs=Nothing
		team.execute("Update ["&IsForum&"User] Set Friend='"&myFriend&"' where ID="& team.TK_UserID)
		Session(CacheName&"_UserLogin") = ""
	End if
	team.error1 " <li> ������ӳɹ������ڽ��Զ����ء�<li> <a href=""Control.asp?action=friend"">���ؿ��������ҳ</a>�� <meta http-equiv=refresh content=3;url=""Control.asp?action=friend"">"	
End Sub

Sub delfriend
	Dim Rs,myFriend,ByName
	ByName= HRF(2,1,"byname")
	myFriend = ""
	Set Rs = team.execute("Select Friend From ["&Isforum&"User] where ID="& team.TK_UserID)
	If Not Rs.Eof Then
		myFriend = Replace(Rs(0),ByName&"|","")
	End if
	Rs.Close:Set Rs=Nothing
	team.execute("Update ["&IsForum&"User] Set Friend='"&myFriend&"' where ID="& team.TK_UserID)
	Session(CacheName&"_UserLogin") = ""
	team.error1 " <li> ����ɾ���ɹ������ڽ��Զ����ء�<li> <a href=""Control.asp?action=friend"">���ؿ��������ҳ</a>�� <meta http-equiv=refresh content=3;url=""Control.asp?action=friend"">"	
End Sub

Sub UserFriend
	Dim tmp,Rs,Ms,Ump,UserInfo,UserMedals,Emp,ExtCredits
	x2 = "<a href=""Control.asp""> ������� </a>"
	x1 = " �����б� "
	ExtCredits = Split(team.Club_Class(21),"|")
	tmp = Replace(Team.UserHtml (1),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"friends"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"uersinfo"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"usercp"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"userbank"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"favorites"))
	Set Rs = team.execute("Select UserGroupID,Levelname,Usermail,Userhome,Userface,UserCity,UserSex,Question,Answer,Honor,Birthday,Sign,Medals,UserInfo,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members,Friend From ["&Isforum&"User] where ID="& team.TK_UserID)
	If Not Rs.Eof Then
		Ms = Rs.GetString(,1, "$$##$$","","")
	End if
	Rs.Close:Set Rs=Nothing
	Ump = Split(Ms,"$$##$$")
	UserInfo = Split(Ump(13),"|")		
	'UserInfo = QQ|ICQ|YAHOO|MSN|TAOBAO|ALIPAY
	tmp = Replace(tmp,"{$usermail}",Ump(2))
	tmp = Replace(tmp,"{$userhome}",Ump(3))
	tmp = Replace(tmp,"{$userface}",Fixjs(Ump(4)))
	tmp = Replace(tmp,"{$sign}",Sign_Code(Ump(11),CID(Split(Ump(1),"||")(4))))
	tmp = Replace(tmp,"{$userqq}",IIf(UserInfo(0)<>"","<a target=""_blank"" href=""tencent://message/?uin="&UserInfo(0)&"&Site=team5.cn&Menu=yes""><img border=""0"" SRC=""http://wpa.qq.com/pa?p=1:"&UserInfo(0)&":7"" alt=""���������ҷ���Ϣ"" onerror=""javascript:this.src='images/qqerr.gif'""></a>",""))
	tmp = Replace(tmp,"{$qq}",UserInfo(0))
	tmp = Replace(tmp,"{$icq}",UserInfo(1))
	tmp = Replace(tmp,"{$yahoo}",UserInfo(2))
	tmp = Replace(tmp,"{$msn}",UserInfo(3))
	tmp = Replace(tmp,"{$taobao}",IIF(UserInfo(4)<>"","<script type=""text/javascript"">document.write('<a target=""_blank"" href=""http://amos1.taobao.com/msg.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2""><img src=""http://amos1.taobao.com/online.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2"" alt=""�Ա�����"" border=""0"" />"&UserInfo(4)&"</a> ');</script>",""))
	tmp = Replace(tmp,"{$alipay}",UserInfo(5))
	If Ump(12)<>"" Then
		UserMedals = ""
		If Instr(Ump(12),"$$$")>0 Then
			Dim i
			UserMedals = Split(Ump(12),"$$$")
			For i = 0 to Ubound(UserMedals)-1
				Emp = Emp & "<img src=""images/plus/"&Split(UserMedals(i),"&&&")(0)&""" align=""absmiddle"" alt="""&Split(UserMedals(i),"&&&")(1)&"""> "
			Next
			tmp = Replace(tmp,"{$userMedals}",Emp)
		End if
	Else
		tmp = Replace(tmp,"{$userMedals}","")
	End if
	Dim Friend,Fmp
	If Len(Ump(29))<2 Then
		tmp = Replace(tmp,"{$isfriends}","")
	Else
		If Instr(Ump(29),"|")>0 Then
			Fmp = Split(Ump(29),"|")
			for i = 0 to Ubound(Fmp)-1
				If Fmp(i) <> "" Then
					Friend = Friend & "<tr class=""tab4""><td> NO."&i+1&" </td><td> "&Fmp(i)&" </td><td> <a href=""msg.asp?action=sendpm&byname="&Fmp(i)&"""> <img src="""&team.styleurl&"/sendpm.gif"" align=""absmiddle"" border=""0"" alt=""���Ͷ���""></a> &nbsp; <a href=""?action=delfriend&byname="&Fmp(i)&"""> <img src="""&team.styleurl&"/delete.gif"" align=""absmiddle"" border=""0"" alt=""ɾ���˺���""></a></td></tr>"
				End if
			Next
		Else
			Friend = "<tr class=""tab4""><td> NO.1 </td><td> "&Ump(29)&" </td><td> <a href=""msg.asp?action=sendpm&byname="&Ump(29)&"""> <img src="""&team.styleurl&"/sendpm.gif"" align=""absmiddle"" border=""0"" alt=""���Ͷ���""></a> &nbsp; <a href=""?action=delfriend&byname="&Ump(29)&"""> <img src="""&team.styleurl&"/delete.gif"" align=""absmiddle"" border=""0"" alt=""ɾ���˺���""></a></td></tr>"
		End if
		tmp = Replace(tmp,"{$isfriends}",Friend)
	End if
	Echo tmp
End Sub

Sub buyuserbank
	Dim buys
	buys = HRF(1,2,"buys")
	If team.Forum_setting(102) = "" Or Len(team.Forum_setting(102))<7 Then
		team.error " ϵͳΪ��ͨ���ֶһ���"
	Else
		If Buys < CID(team.Forum_setting(105)) Then
			team.error " ������С��ϵͳ���� [����"&CID(team.Forum_setting(105))&"]�����ױ�ȡ����"
		Else
			If DateDiff("s",Request.Cookies("times")("buytime"),Now()) < 120 Then
				team.error " ÿ�ι���ʱ�䲻������120�� "
			Else
				team.Execute("insert into ["&Isforum&"BankLog] (bankname,buyname,buyvalue,getvalue,posttime,Makes) values ('"&Replace(Replace(Replace(now(),":","")," ",""),"-","")&team.TK_UserID&"','"&tk_UserName&"',"&CID(buys/team.Forum_setting(104))&","&buys&","&SqlNowString&",0)")
				Response.Redirect "API/Payto.asp?price="&buys
				team.error1 " <li> ����ɹ�����ȴ�ϵͳ����Ա��ˣ����ڽ��Զ����ء�<li> <a href=""Control.asp?action=usercp"">���ؿ��������ҳ</a>�� <meta http-equiv=refresh content=3;url=""Control.asp?action=usercp"">"
				Response.Cookies("times")("buytime") = Now()
			End if
		End if
	End if
End Sub

Sub edituserbank
	Dim toname,Rs,Rs1,rewardprice,Userrewardprice,ExtCredits
	ExtCredits = Split(team.Club_Class(21),"|")
	toname = HRF(1,1,"toname")
	rewardprice = HRF(1,2,"rewardprice")
	Userrewardprice = rewardprice * ( 1 + team.Forum_setting(11) )
	If CID(rewardprice) < Cid(team.Forum_setting(12)) Then
		team.error "ת�˶�ͣ��޷���ɽ��ס�"
	End if
	Set Rs = team.execute("Select * From ["&Isforum&"User] Where UserName='"&toname&"'")
	If Rs.Eof Then
		team.Error " ϵͳ�����ڴ��û��� "
	Else
		Set Rs1 = team.execute("Select Extcredits"&Cid(team.Forum_setting(99))&" From ["&Isforum&"User] Where UserName='"&TK_UserName&"'")
		If Not Rs.Eof Then
			If CID(Rs1(0)) <= CID(Userrewardprice) Then 
				team.error " ��������������ת�ˡ�"
			Else
				team.execute("Update ["&Isforum&"User] Set Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"-"&CID(Userrewardprice)&" Where UserName='"&TK_UserName&"'")
				team.execute("Update ["&Isforum&"User] Set Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"+"&CID(rewardprice)&",Newmessage=Newmessage+1 Where UserName='"&toname&"'")
				team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&TK_UserName&"','"&toname&"','��ϲ�����û�"&tk_UserName&"ת����"&rewardprice&"���"&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&"�������ʻ������½��<a href=""Control.asp?action=bank"">����ת�˹���</a>���鿴���Ļ�����',"&SqlNowString&",'����ת��֪ͨ',0)")
				team.error1 " <li> ת�˳ɹ�����ȴ�ϵͳ�Զ����ء�<li> <a href=""Control.asp?action=usercp"">���ؿ��������ҳ</a>�� <meta http-equiv=refresh content=3;url=""Control.asp?action=usercp"">"
			End if
		End if
		Rs1.Close:Set Rs1 = Nothing
	End if
	Rs.Close:Set Rs = Nothing
End Sub

Sub UserBank
	Dim tmp,Rs,Ms,Ump,UserInfo,UserMedals,Emp,ExtCredits
	x2 = "<a href=""Control.asp""> �������</a>"
	x1 = " �û�"&TK_UserName&" "
	ExtCredits = Split(team.Club_Class(21),"|")
	tmp = Replace(Team.UserHtml (1),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"userbank"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"uersinfo"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"usercp"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"friends"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"favorites"))
	Set Rs = team.execute("Select UserGroupID,Levelname,Usermail,Userhome,Userface,UserCity,UserSex,Question,Answer,Honor,Birthday,Sign,Medals,UserInfo,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members From ["&Isforum&"User] where ID="& team.TK_UserID)
	If Not Rs.Eof Then
		Ms = Rs.GetString(,1, "$$##$$","","")
	End if
	Rs.Close:Set Rs=Nothing
	Ump = Split(Ms,"$$##$$")
	UserInfo = Split(Ump(13),"|")		
	'UserInfo = QQ|ICQ|YAHOO|MSN|TAOBAO|ALIPAY
	tmp = Replace(tmp,"{$usermail}",Ump(2))
	tmp = Replace(tmp,"{$userhome}",Ump(3))
	tmp = Replace(tmp,"{$userface}",Fixjs(Ump(4)))
	tmp = Replace(tmp,"{$sign}",Sign_Code(Ump(11),CID(Split(Ump(1),"||")(4))))
	tmp = Replace(tmp,"{$userqq}",IIf(UserInfo(0)<>"","<a target=""_blank"" href=""tencent://message/?uin="&UserInfo(0)&"&Site=team5.cn&Menu=yes""><img border=""0"" SRC=""http://wpa.qq.com/pa?p=1:"&UserInfo(0)&":7"" alt=""���������ҷ���Ϣ"" onerror=""javascript:this.src='images/qqerr.gif'""></a>",""))
	tmp = Replace(tmp,"{$qq}",UserInfo(0))
	tmp = Replace(tmp,"{$icq}",UserInfo(1))
	tmp = Replace(tmp,"{$yahoo}",UserInfo(2))
	tmp = Replace(tmp,"{$msn}",UserInfo(3))
	tmp = Replace(tmp,"{$taobao}",IIF(UserInfo(4)<>"","<script type=""text/javascript"">document.write('<a target=""_blank"" href=""http://amos1.taobao.com/msg.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2""><img src=""http://amos1.taobao.com/online.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2"" alt=""�Ա�����"" border=""0"" />"&UserInfo(4)&"</a> ');</script>",""))
	tmp = Replace(tmp,"{$alipay}",UserInfo(5))
	If Ump(12)<>"" Then
		UserMedals = ""
		If Instr(Ump(12),"$$$")>0 Then
			Dim i
			UserMedals = Split(Ump(12),"$$$")
			For i = 0 to Ubound(UserMedals)-1
				Emp = Emp & "<img src=""images/plus/"&Split(UserMedals(i),"&&&")(0)&""" align=""absmiddle"" alt="""&Split(UserMedals(i),"&&&")(1)&"""> "
			Next
			tmp = Replace(tmp,"{$userMedals}",Emp)
		End if
	Else
		tmp = Replace(tmp,"{$userMedals}","")
	End if
	tmp = Replace(tmp,"{$forumride}",team.Forum_setting(11))
	tmp = Replace(tmp,"{$minpower}",Cid(team.Forum_setting(12)))
	tmp = Replace(tmp,"{$nowbanks}",IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  " ( "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" ) "," (������δ����) "))
	tmp = Replace(tmp,"{$mybanks}",Ump(20+Cid(team.Forum_setting(99))) &"  "&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(1))
	tmp = Replace(tmp,"{$buyrate}",CID(team.Forum_setting(104)))
	tmp = Replace(tmp,"{$minbuys}",CID(team.Forum_setting(105)))
	tmp = Replace(tmp,"{$isbuybank}",IIF(team.Forum_setting(102) = "" Or Len(team.Forum_setting(102))<7,"Display:none",""))
	Echo tmp
End Sub

Sub Edituserinfo
	Dim tmp
	team.ChkPost()
	If strLength(HRF(1,1,"sign"))>CID(team.Group_Browse(23)) or strLength(HRF(1,1,"sign"))>255 Then
		team.error "ǩ�����Ȳ��ܴ��� "&team.Group_Browse(23)&" �����255�ַ� ���뷵���޸ġ�"
	End if
	If HRF(1,1,"qq")<>"" Then
		If Not IsNumeric(HRF(1,1,"qq")) Then
			team.error " QQ ��ʽ���� "
		End if
	End if
	If HRF(1,1,"icq")<>"" Then
		If Not IsNumEric(HRF(1,1,"icq")) Then
			team.error " ICQ ��ʽ����"
		End if
	End if
	Dim Brithday,SQL,Questionid,Answer
	Brithday = HRF(1,1,"birthday")
	If Brithday <>"" Then
		If Not IsDate(trim(Brithday)) Then Brithday = ""
	End If
	If Len(HRF(1,1,"oldpassword"))> 1 Then
		If Len(HRF(1,1,"oldpassword"))<3 Then 
			team.Error "����Ҫ������ȷ��ԭʼ����ſ����޸ĵ�½����,�����ַ�Լ���ȫ����"
		Else
			If MD5(Trim(HRF(1,1,"oldpassword")),16) <> Trim(team.User_SysTem(1)) Then
				team.error1 "<li> �û����������ȴ�ϵͳ���غ��������� �� <meta http-equiv=refresh content=3;url=""Control.asp?action=usercp"">"
			Else
				If Trim(HRF(1,1,"newpassword")) <> Trim(HRF(1,1,"newpassword2")) Then
					team.error1 "<li> ����������û����벻��ͬ����ȴ�ϵͳ���غ��������� �� <meta http-equiv=refresh content=3;url=""Control.asp?action=usercp"">"
				Else
					If Trim(HRF(1,1,"newpassword"))<>"" And Trim(HRF(1,1,"newpassword2"))<>"" Then
						SQL = "UserPass='"&MD5(HRF(1,1,"newpassword"),16)&"',"
						tmp = tmp & "<li> �����޸ĳɹ� �� "
						Response.Cookies(Forum_sn)("username")=""
						Response.Cookies(Forum_sn)("userpass")=""
						Response.Cookies(Forum_sn)("UserMember")=""
						Response.Cookies(Forum_sn)("UserID")="0"
						Session("UserMember")= ""
						Session("Admin_Pass")=""
					End if
				End if
				Questionid = HRF(1,1,"questionid")
				Answer = HRF(1,1,"answer")
				If Questionid&""<>"" Then
					If Answer&""= "" Then 
						team.error1 " �������˰�ȫ����ѡ�������д�����Ļش�ѡ���ȴ�ϵͳ���غ��������� �� <meta http-equiv=refresh content=3;url=""Control.asp?action=usercp"">"
					End if
				Else
					If Answer&""<> "" Then 
						team.error1 " ��δ���ð�ȫ����ѡ���ȴ�ϵͳ���غ��������� �� <meta http-equiv=refresh content=3;url=""Control.asp?action=usercp"">"
					End if
				End If
				If HRF(1,1,"emailnew")&""="" Then 
					team.Error "�ʼ���ַ����Ϊ��"
				Else
					team.execute("Update ["&IsForum&"User] Set "&SQL&" Usermail='"&HRF(1,1,"emailnew")&"',Question='"&Questionid&"',Answer='"&Answer&"' where ID="& team.TK_UserID)
				End If
			End If
		End if
	End If
	Dim TSign
	TSign = HRF(1,1,"sign")
	If team.Group_Browse(21) = 0 Then
		TSign = Replace(TSign,"[","��")
	End If
	If team.Group_Browse(22) = 0 Then
		If InStr(TSign,"[img]")>0 Or InStr(TSign,"[IMG]")>0 Then
			team.Error "ǩ����֧��[img]����"
		End If
	End If 
	team.execute("Update ["&IsForum&"User] Set Userhome='"&HRF(1,1,"userhome")&"',Userface='"&Fixjs(HRF(1,1,"urlavatar"))&"',UserCity='"&HRF(1,1,"usercity")&"',UserSex='"&HRF(1,2,"usersex")&"',Honor='"&HRF(1,1,"honor")&"',Birthday='"&Brithday&"',Sign='"&TSign&"',UserInfo='"&HRF(1,1,"qq")&"|"&HRF(1,1,"icq")&"|"&HRF(1,1,"yahoo")&"|"&HRF(1,1,"msn")&"|"&HRF(1,1,"taobao")&"|"&HRF(1,1,"alipay")&"' where ID="& team.TK_UserID)
	If HRF(1,1,"tppnew") <>"" OR HRF(1,1,"pppnew")<>"" OR HRF(1,1,"msgsound")<>"" Then
		If HRF(1,1,"msgsound") <>"" Then
			Response.Cookies(Forum_sn)("msgsound") = HRF(1,1,"msgsound")
		End if
		If HRF(1,1,"tppnew") <>"" Then
			Response.Cookies(Forum_sn)("tppnew") = HRF(1,1,"tppnew")
		End if
		If HRF(1,1,"pppnew") <>"" Then
			Response.Cookies(Forum_sn)("pppnew") = HRF(1,1,"pppnew")
		End if
		'�ж�Cookies����Ŀ¼
		Dim cookies_path_s,cookies_path_d,cookies_path,i
		cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
		cookies_path_d=ubound(cookies_path_s)
		cookies_path="/"
		For i=1 to cookies_path_d-1
			cookies_path=cookies_path&cookies_path_s(i)&"/"
		Next
		Response.Cookies(Forum_sn).path = cookies_path
	End If
	Cache.DelCache("UserBirthdays")
	Application(CacheName&"_Nobady") = 0
	team.error1 " "&tmp&"<li> �û����ϱ���ɹ�����ȴ�ϵͳ�Զ����ء�<li> <a href=""Control.asp?action=usercp"">���ؿ��������ҳ</a>�� <meta http-equiv=refresh content=3;url=""Control.asp?action=usercp"">"
End Sub

Sub UserCp
	Dim tmp,Rs,Ms,Ump,UserInfo,UserMedals,Emp
	x2 = "<a href=""Control.asp""> �������</a>"
	x1 = " �û�"&TK_UserName&" "
	tmp = Replace(Team.UserHtml (1),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"usercp"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"uersinfo"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"userbank"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"friends"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"favorites"))
	tmp = Replace(tmp,"{$allface}",CID(team.Forum_setting(100)))
	Set Rs = team.execute("Select UserGroupID,Levelname,Usermail,Userhome,Userface,UserCity,UserSex,Question,Answer,Honor,Birthday,Sign,Medals,UserInfo,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members From ["&Isforum&"User] where ID="& team.TK_UserID)
	If Not Rs.Eof Then
		Ms = Rs.GetString(,1, "$$##$$","","")
	End if
	Rs.Close:Set Rs=Nothing
	Ump = Split(Ms,"$$##$$")
	UserInfo = Split(Ump(13),"|")		
	'UserInfo = QQ|ICQ|YAHOO|MSN|TAOBAO|ALIPAY
	tmp = Replace(tmp,"{$usermail}",Ump(2))
	tmp = Replace(tmp,"{$userhome}",Ump(3))
	tmp = Replace(tmp,"{$userface}",Fixjs(Ump(4)))
	tmp = Replace(tmp,"{$usercity}",Ump(5))
	tmp = Replace(tmp,"{$sex}",Ump(6))
	tmp = Replace(tmp,"{$userquest}",IIF(Ump(7)<>"","<option value="""&Ump(7)&""" checked>"&Ump(7)&"</option>",""))
	tmp = Replace(tmp,"{$answer}",Ump(8))
	tmp = Replace(tmp,"{$honor}",Ump(9))
	tmp = Replace(tmp,"{$brithday}",Ump(10))
	tmp = Replace(tmp,"{$sign}",Replace(Ump(11),"<BR>",CHR(10)))
	tmp = Replace(tmp,"{$userqq}",IIf(UserInfo(0)<>"","<a target=""_blank"" href=""tencent://message/?uin="&UserInfo(0)&"&Site=team5.cn&Menu=yes""><img border=""0"" SRC=""http://wpa.qq.com/pa?p=1:"&UserInfo(0)&":7"" alt=""���������ҷ���Ϣ"" onerror=""javascript:this.src='images/qqerr.gif'""></a>",""))
	tmp = Replace(tmp,"{$qq}",UserInfo(0))
	tmp = Replace(tmp,"{$icq}",UserInfo(1))
	tmp = Replace(tmp,"{$yahoo}",UserInfo(2))
	tmp = Replace(tmp,"{$msn}",UserInfo(3))
	tmp = Replace(tmp,"{$taobao}",IIF(UserInfo(4)<>"","<script type=""text/javascript"">document.write('<a target=""_blank"" href=""http://amos1.taobao.com/msg.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2""><img src=""http://amos1.taobao.com/online.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2"" alt=""�Ա�����"" border=""0"" />"&UserInfo(4)&"</a> ');</script>",""))
	tmp = Replace(tmp,"{$taobao1}",UserInfo(4))
	tmp = Replace(tmp,"{$alipay}",UserInfo(5))
	If Ump(12)<>"" Then
		UserMedals = ""
		If Instr(Ump(12),"$$$")>0 Then
			Dim i
			UserMedals = Split(Ump(12),"$$$")
			For i = 0 to Ubound(UserMedals)-1
				Emp = Emp & "<img src=""images/plus/"&Split(UserMedals(i),"&&&")(0)&""" align=""absmiddle"" alt="""&Split(UserMedals(i),"&&&")(1)&"""> "
			Next
			tmp = Replace(tmp,"{$userMedals}",Emp)
		End if
	Else
		tmp = Replace(tmp,"{$userMedals}","")
	End if
	tmp = Replace(tmp,"{$signmax}",team.Group_Browse(23))
	tmp = Replace(tmp,"{$msgsound}",IIf(Request.Cookies(Forum_Sn)("msgsound")="",1,Request.Cookies(Forum_Sn)("msgsound")))
	tmp = Replace(tmp,"{$tppnew}",IIF(Request.Cookies(Forum_Sn)("tppnew")="","<option value="""&CID(team.Forum_setting(19))&""" selected=""selected"">- ʹ��Ĭ�� -</option>","<option value="""&CID(Request.Cookies(Forum_Sn)("tppnew"))&""" selected=""selected"">"&CID(Request.Cookies(Forum_Sn)("tppnew"))&"</option>"))
	tmp = Replace(tmp,"{$pppnew}",IIF(Request.Cookies(Forum_Sn)("pppnew")="","<option value="""&CID(team.Forum_setting(20))&""" selected=""selected"">- ʹ��Ĭ�� -</option>","<option value="""&CID(Request.Cookies(Forum_Sn)("pppnew"))&""" selected=""selected"">"&CID(Request.Cookies(Forum_Sn)("pppnew"))&"</option>"))
	Echo tmp
End Sub

Sub Main()
	Dim tmp,Rs,UserInfo,Emp,Post,i,PostTmp
	x1 = "<a href=""Control.asp"">�������</a>"
	tmp = Replace(Team.UserHtml (1),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"uersinfo"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"usercp"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"userbank"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"friends"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"favorites"))
	tmp = Replace(tmp,"{$username}",TK_UserName)
	tmp = Replace(tmp,"{$uid}",team.TK_UserID)
	tmp = Replace(tmp,"{$myurls}",team.Club_Class(2))
	Set Rs = Team.Execute("Select Top 5 id,Author,MsgTopic,Sendtime From ["&IsForum&"Message] Where incept='"&TK_UserName&"' and isbak=0 Order By Sendtime Desc ")
	If Not Rs.Eof Then
		Post = Rs.GetRows(5)
	End If
	RS.Close:Set Rs=Nothing
	If IsArray(Post) Then
		For i=0 To Ubound(Post,2)
			PostTmp = PostTmp & "<tr class=""tab4""><td align=""left""><a href=""Msg.asp?tid="&Post(0,i)&"""  target=""_blank"">"&Post(2,i)&"</a></td><td>"&Post(1,i)&"</td><td>"&Post(3,i)&"</td></tr>"
		Next
	End If
	tmp = Replace(tmp,"{$newmsg}",PostTmp)
	Post = "" : PostTmp = ""
	Set Rs = Team.Execute("Select Top 5 id,topic,Lasttime,Views,Replies,GoodTopic From ["&IsForum&"Forum] Where Deltopic=0 And UserName='"&TK_UserName&"' Order By LastTime Desc ")
	If Not Rs.Eof Then
		Post = Rs.GetRows(5)
	End If
	RS.Close:Set Rs=Nothing
	If IsArray(Post) Then
		For i=0 To Ubound(Post,2)
			PostTmp = PostTmp & "<tr class=""tab4""><td align=""left""><a href=""thread.asp?tid="&Post(0,i)&""" target=""_blank"">"&Post(1,i)&"</a> "
			If Post(5,i) = 1 Then PostTmp = PostTmp & " <img Src="""&team.styleurl&"/f_good.gif"" align=""absmiddle"" alt=""�����Ѿ������뾫����""> "
			PostTmp = PostTmp & " </td><td>"&Post(3,i)&" / "&Post(4,i)&"</td><td>"&Post(2,i)&"</td></tr>"
		Next
	End If
	tmp = Replace(tmp,"{$newtopic}",PostTmp)
	Post = "" : PostTmp = ""
	Set Rs=Team.Execute("Select DISTINCT Top 5 B.id,B.Topic,B.Lasttime,B.Views,B.Replies,U.Topicid From "&IsForum & team.Club_Class(11)&" U Inner Join ["&IsForum&"Forum] B On U.Topicid=B.ID Where U.UserName='"&TK_UserName&"' and B.Deltopic=0 Order By B.Lasttime Desc")
	If Not Rs.Eof Then
		Post = Rs.GetRows(5)
	End If
	RS.Close:Set Rs=Nothing
	If IsArray(Post) Then
		For i=0 To Ubound(Post,2)
			PostTmp = PostTmp & "<tr class=""tab4""><td align=""left""><a href=""thread.asp?tid="&Post(0,i)&""" target=""_blank"">"&Post(1,i)&"</a></td><td>"&Post(3,i)&" / "&Post(4,i)&"</td><td>"&Post(2,i)&"</td></tr>"
		Next
	End If
	tmp = Replace(tmp,"{$newretopic}",PostTmp)
	Set Rs = team.execute("Select UserGroupID,Levelname,Usermail,Userhome,Userface,UserCity,UserSex,Question,Answer,Honor,Birthday,Sign,Medals,UserInfo,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members From ["&Isforum&"User] where ID="& team.TK_UserID)
	If Not Rs.Eof Then
		Ms = Rs.GetString(,1, "$$##$$","","")
	End if
	Rs.Close:Set Rs=Nothing
	Ump = Split(Ms,"$$##$$")
	UserInfo = Split(Ump(13),"|")			
	'UserInfo = QQ|ICQ|YAHOO|MSN|TAOBAO|ALIPAY
	tmp = Replace(tmp,"{$usermail}",Ump(2))
	tmp = Replace(tmp,"{$userhome}",Ump(3))
	tmp = Replace(tmp,"{$userface}",Fixjs(Ump(4)))
	tmp = Replace(tmp,"{$usercity}",Ump(5))
	tmp = Replace(tmp,"{$sex}",Ump(6))
	tmp = Replace(tmp,"{$userquest}","<option value="""&Ump(7)&""" checked>"&Ump(7)&"</option>")
	tmp = Replace(tmp,"{$answer}",Ump(8))
	tmp = Replace(tmp,"{$honor}",Ump(9))
	tmp = Replace(tmp,"{$brithday}",Ump(10))
	tmp = Replace(tmp,"{$sign}",Sign_Code(Ump(11),CID(Split(Ump(1),"||")(4))))
	tmp = Replace(tmp,"{$userqq}",IIf(UserInfo(0)<>"","<a target=""_blank"" href=""tencent://message/?uin="&UserInfo(0)&"&Site=team5.cn&Menu=yes""><img border=""0"" SRC=""http://wpa.qq.com/pa?p=1:"&UserInfo(0)&":7"" alt=""���������ҷ���Ϣ"" onerror=""javascript:this.src='images/qqerr.gif'""></a>",""))
	tmp = Replace(tmp,"{$icq}",UserInfo(1))
	tmp = Replace(tmp,"{$yahoo}",UserInfo(2))
	tmp = Replace(tmp,"{$msn}",UserInfo(3))
	tmp = Replace(tmp,"{$taobao}",IIF(UserInfo(4)<>"","<script type=""text/javascript"">document.write('<a target=""_blank"" href=""http://amos1.taobao.com/msg.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2""><img src=""http://amos1.taobao.com/online.ww?v=2&amp;uid='+encodeURIComponent('"&UserInfo(4)&"')+'&amp;s=2"" alt=""�Ա�����"" border=""0"" />"&UserInfo(4)&"</a> ');</script>",""))
	tmp = Replace(tmp,"{$alipay}",UserInfo(5))
	If Ump(12)<>"" Then
		Dim Ms,Ump,UserMedals
		UserMedals = ""
		If Instr(Ump(12),"$$$")>0 Then
			UserMedals = Split(Ump(12),"$$$")
			For i = 0 to Ubound(UserMedals)-1
				Emp = Emp & "<img src=""images/plus/"&Split(UserMedals(i),"&&&")(0)&""" align=""absmiddle"" alt="""&Split(UserMedals(i),"&&&")(1)&"""> "
			Next
			tmp = Replace(tmp,"{$userMedals}",Emp)
		End if
	Else
		tmp = Replace(tmp,"{$userMedals}","")
	End if
	Echo tmp
End Sub

Function JsMYFace(s)
	Dim Str
	Str = s
	face=Fixjs(Replace(face,"'",""))
	face=Replace(face,"..","")
	face=Replace(face,"\","/")
	face=Replace(face,"^","")
	face=Replace(face,"#","")
	face=Replace(face,"%","")
	face=Replace(face,"|","")
	face=Server.htmlencode(Left(face,200))
End Function

%>
