<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Call testUser()
Dim x1,x2,Fid
Dim sID
sID = HRF(2,2,"sid")
team.Headers(" ���� / PM ")
Select Case Request("action")
	Case "sendpm"
		Call SendPm
	Case "sendpmok"
		Call sendpmok
	Case "readmsg"
		Call readmsg
	Case "deletes"
		Call deletes
	Case "delmsgs"
		Call delmsgs
	Case "nowpostmsg"
		Call nowpostmsg
	Case else
		Call Main()
End Select
team.footer()

Sub delmsgs
	Dim ho
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"Message] Where incept='"& tk_UserName &"' and ID="&Int(ho))
	next
	team.error1 "<li>��Ϣ�Ѿ�ɾ�������������� <a href=""Msg.asp"">���ض�����</a> ��ȴ�ϵͳ�Զ����� <meta http-equiv=refresh content=3;url=Msg.asp> "
End Sub

Sub nowpostmsg
	If team.execute("Select * from ["&Isforum&"Message] Where ID="&sID).Eof Then
		team.error " ָ���Ĳ�������"
	Else
		UpdateUserpostExc()
		team.execute("Update ["&Isforum&"Message] Set isbak=0 Where ID="&sID)
		team.error1 "<li>��Ϣ�Ѿ����͡����������� <a href=""Msg.asp"">���ض�����</a> ��ȴ�ϵͳ�Զ����� <meta http-equiv=refresh content=3;url=Msg.asp> "
	End if
End Sub

Sub UpdateUserpostExc()
	'�û����ֲ���
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
				UExt = UExt &"Extcredits0=Extcredits0-"&MustSort(5)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"-"&MustSort(5)&""
			End If
			If (team.User_SysTem(14+U)-MustSort(5))-MustSort(8)<0 Then
				team.Error "����"&ExtSort(0)&" ["& team.User_SysTem(14+U) - MustSort(5) &"] ���ڻ��ֲ�������ֵ ["& MustSort(8)&"] �������޷����д˲�����"
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where ID = "& team.TK_UserID)
End Sub

Sub sendpmok
	Dim Umsg,i
	If Len(HRF(1,1,"subject"))<2  Then
		team.error2 "���ű��ⲻ������2���ַ�!"
	End If
	UpdateUserpostExc()
	If HRF(1,1,"msgto") = tk_userName Then
		team.Error "���ն�����Ϊ�Լ�"
	End if
	If Request("chkall") = "on" Then
		Umsg = Split(Replace(HRF(1,1,"msgtobuddys")," ",""),",")
		for i = 0 to Ubound(Umsg)
			team.Execute("Update ["&Isforum&"User] set Newmessage=Newmessage+1 Where UserName='"&HtmlEncode(Umsg(i))&"'")
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&TK_UserName&"','"&HTmlEncode(Umsg(i))&"','"&HRF(1,1,"message")&"',"&SqlNowString&",'"&HRF(1,1,"subject")&"',"&HRF(1,2,"saveoutbox")&")")
		Next
	End if
	team.Execute("Update ["&Isforum&"User] set Newmessage=Newmessage+1 Where UserName='"&HRF(1,1,"msgto")&"'")
	team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&TK_UserName&"','"&HRF(1,1,"msgto")&"','"&HRF(1,1,"message")&"',"&SqlNowString&",'"&HRF(1,1,"subject")&"',"&HRF(1,2,"saveoutbox")&")")
	If HRF(1,2,"saveoutbox") = 1 Then
		team.error1 "<li>��Ϣ�Ѿ�����ݸ��� ������Ҫ���Ͷ��ţ���鿴���Ĳݸ��䡣���������� <a href=""Msg.asp"">���ض�����</a> ��ȴ�ϵͳ�Զ����� <meta http-equiv=refresh content=3;url=Msg.asp> "
	Else
		team.error1 "<li>��Ϣ�Ѿ����ͳɹ������������� <a href=""Msg.asp"">���ض�����</a> ��ȴ�ϵͳ�Զ����� <meta http-equiv=refresh content=3;url=Msg.asp> "
	End if
End Sub

Sub deletes
	If sID = "" or Not IsNumeric(sID) Then 
		team.error "��������"
	Else
		team.execute("Delete From ["&IsForum&"Message] Where incept='"& tk_UserName &"' and  ID="& sID)
		team.error1 "<li>��Ϣ�Ѿ�ɾ���������� <a href=""Msg.asp"">���ض�����</a> ��ȴ�ϵͳ�Զ����� <meta http-equiv=refresh content=3;url=Msg.asp> "
	End if
End Sub

Sub readmsg
	Dim tmp,incept,IsPage
	Dim Rs,sID
	sID = HRF(2,2,"sid")
	InCept = HRF(2,1,"incept")
	X1="<a href=""Msg.asp"">�鿴���ж���</a>"
	if team.Newmessage>0 then
		Team.execute("update ["&IsForum&"user] Set Newmessage=0 Where ID="& team.TK_UserID)
		Session(CacheName&"_UserLogin")=""
	End if
	tmp = Replace(Team.UserHtml (2),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"readpm"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"newpm"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"pages"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"sendpm"))
	Set Rs = team.execute("Select ID,incept,author,msgtopic,Content,Sendtime From ["&IsForum&"Message] Where incept= '"&TK_UserName&"' and ID="&sID)
	If Rs.Eof And Rs.Bof Then
		team.error "ָ����ID�����ڻ������ܲ鿴�����û��Ķ������ݡ�"
	Else
		tmp = Replace(tmp,"{$sid}",Rs(0))
		tmp = Replace(tmp,"{$msgname}",Rs(2))
		tmp = Replace(tmp,"{$msgtitle}",Rs(3))
		tmp = Replace(tmp,"{$msgcontent}",Ubb_Code(Replace(Rs(4),"'","''")))
		tmp = Replace(tmp,"{$msgtime}",Rs(5))
		tmp = Replace(tmp,"{$send}",IIF(Request("send")=1,"- <a href=""Msg.asp?action=nowpostmsg&sid="&Rs(0)&""">��������</a>",""))
	End if
	Rs.close:Set Rs=Nothing
	IsPage = team.execute("Select Count(ID) From ["&IsForum&"Message] Where incept= '"&TK_UserName&"'")(0)
	If IsPage<1 or Not IsNumeric(IsPage) Then IsPage = 1
	tmp = Replace(tmp,"{$countmessage}",IsPage)
	tmp = Replace(tmp,"{$messcount}",CID(team.Group_Browse(12)))
	Dim MyMsg
	MyMsg = CID(team.Group_Browse(12))
	If MyMsg = 0 Then MyMsg = 1
	tmp = Replace(tmp,"{$widse}",IsPage*100/MyMsg)
	tmp = Replace(tmp,"{$messcount}",CID(team.Group_Browse(12)))
	Echo tmp
End Sub

Sub SendPm
	Dim tmp,incept,TWhere,i,mmp,SQL
	Dim IsPage,Page,RS,mRs,Maxpage,PageNum
	InCept = HRF(2,1,"incept")
	X1="<a href=""Msg.asp"">�鿴���ж���</a>"
	tmp = Replace(Team.UserHtml (2),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"newpm"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"pages"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"sendpm"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"readpm"))
	IsPage = team.execute("Select Count(ID) From ["&IsForum&"Message] Where incept= '"&TK_UserName&"'")(0)
	If IsPage<1 or Not IsNumeric(IsPage) Then IsPage = 1
	tmp = Replace(tmp,"{$countmessage}",IsPage)
	tmp = Replace(tmp,"{$messcount}",CID(team.Group_Browse(12)))
	Dim MyMsg
	MyMsg = CID(team.Group_Browse(12))
	If MyMsg = 0 Then MyMsg = 1
	tmp = Replace(tmp,"{$widse}",IsPage*100/MyMsg)
	tmp = Replace(tmp,"{$msgtitle}",IIF(HRF(2,1,"msgtitle")="","","[�ظ�]��" & HRF(2,1,"msgtitle")))
	tmp = Replace(tmp,"{$byname}",IIF(HRF(2,1,"byname")="","",HRF(2,1,"byname")))
	If HRF(2,1,"shows") = "" Then
		tmp = Replace(tmp,"{$showcontent}","")
	Else
		Set Rs = team.execute("Select Content From ["&IsForum&"Message] Where isbak=0 and incept= '"&TK_UserName&"' and ID="&HRF(2,2,"sid"))
		If Rs.Eof Then
			tmp = Replace(tmp,"{$showcontent}","")
		Else
			tmp = Replace(tmp,"{$showcontent}","[B]ת����Ϣ��[/B] [br] "& CHR(10) & "[quote]"& Rs(0) & "[/quote]")
		End if
		Rs.Close:Set Rs=Nothing
	End if
	If team.User_SysTem(23)="" Then
		tmp = Replace(tmp,"{$allbody}","")
	Else
		Dim Umsg,Rmsg
		Umsg = Split(team.User_SysTem(23),"|")
		for i = 0 to Ubound(Umsg)-1
			Rmsg = Rmsg & " <input class=""checkbox"" type=""checkbox"" name=""msgtobuddys"" value="""&Umsg(i)&"""> "&Umsg(i)&""
		Next
		tmp = Replace(tmp,"{$allbody}",Rmsg)
	End if
	Echo tmp
End Sub


Sub Main()
	Dim tmp,incept,TWhere,i,mmp,SQL,forsend
	Dim IsPage,Page,RS,mRs,Maxpage,PageNum
	InCept = HRF(2,1,"incept")
	X1="<a href=""Msg.asp"">�鿴���ж���</a>"
	tmp = Replace(Team.UserHtml (2),"{$weburl}",team.MenuTitle)
	if team.Newmessage>0 then
		Team.execute("update ["&IsForum&"user] Set Newmessage=0 Where ID="& team.TK_UserID)
		Session(CacheName&"_UserLogin")=""
	End if
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"sendpm"))
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"pages"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"newpm"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"readpm"))
	Select Case Request("send")
		Case 1
			TWhere=" isbak=1 and author= '"&TK_UserName&"'"
			tmp=Replace(tmp,"{$pmname}","���Ͷ���")
			forsend = "&send=1"
		Case 2
			TWhere=" isbak=0 and author= '"&TK_UserName&"'"
			tmp=Replace(tmp,"{$pmname}","���Ͷ���")
			forsend = ""
		Case Else
			TWhere=" isbak=0 and incept= '"&TK_UserName&"'"
			tmp=Replace(tmp,"{$pmname}","����")
			forsend = ""
	End Select
	IsPage = team.execute("Select Count(ID) From ["&IsForum&"Message] Where "&TWhere&"")(0)
	If IsPage<1 or Not IsNumeric(IsPage) Then IsPage = 1
	SQL="Select ID,incept,author,msgtopic,Content,Sendtime From ["&IsForum&"Message] Where "&TWhere&" Order By Sendtime DESC"
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum=SqlQueryNum+1
		Maxpage = 20								'ÿҳ��ҳ��
		PageNum = Abs(int(-Abs(IsPage/Maxpage)))	'ҳ��
		Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'��ǰҳ
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		mRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
	If Not Isarray(mRs) Then
		tmp=Replace(tmp,"{$msgcontent}","")
	Else
		For i=0 To Ubound(mRs,2)
			mmp = mmp & "<tr class=""tab4"">"
			mmp = mmp & " <td> <input Name=""newid"" type=""hidden"" value="""&mRS(0,i)&"""><input type=""checkbox"" name=""deleteid"" value="""&mRS(0,i)&""" class=""checkbox"" " 
			If INt(Request("send"))=2 Then
				mmp = mmp & "disabled=disabled"
			End if
			mmp = mmp & "></td>"
			mmp = mmp & " <td align=""left""><a href=""Msg.asp?action=readmsg&sid="&mRS(0,i) & forsend & """>"&mRS(3,i)&"</td>"
			mmp = mmp & " <td>"&mRS(2,i)&"</td>"
			mmp = mmp & " <td>"&mRS(5,i)&"</td>"
			mmp = mmp & "</tr>"
		Next
		tmp=Replace(tmp,"{$msgcontent}",mmp)
	End if
	tmp = Replace(tmp,"{$countmessage}",IsPage)
	tmp = Replace(tmp,"{$messcount}",CID(team.Group_Browse(12)))
	Dim MyMsg
	MyMsg = CID(team.Group_Browse(12))
	If MyMsg = 0 Then MyMsg = 1
	tmp = Replace(tmp,"{$widse}",IsPage*100/MyMsg)
	tmp = Replace(tmp,"{$TotalPage}",IsPage)
	tmp = Replace(tmp,"{$allpage}",PageNum)
	Echo tmp
End Sub
%>