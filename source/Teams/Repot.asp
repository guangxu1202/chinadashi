<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim tID,retopicid,x1,x2,fid
tID = HRF(2,2,"tid")
fID = HRF(2,2,"fid")
Retopicid = HRF(2,2,"retopicid")
TestUser()
If HRF(2,1,"action") = "repot" Or HRF(2,1,"action") = "repotok" Then
	team.Headers(Team.Club_Class(1) & " - ��������")
Else
	team.Headers(Team.Club_Class(1) & " - ��������")
End If
Select Case HRF(2,1,"action")
	Case "repot"
		Call repot
	Case "repotok"
		Call repotok
	Case "upext"
		Call upext
	Case "upextok"
		Call upextok
End Select
team.footer

Function tonames(a,b)
	Dim n,i
	n = team.GroupManages
	tonames = True 
	If team.ManageUser Then
		If CID(team.Forum_setting(61)) = 1 Then
			If team.BoardMaster Then
				If IsArray(n) Then
					For i = 0 To UBound(n,2)
						If Trim(n(1,i)) = Trim(tk_UserName) And CID(n(2,i)) <> fID Then
							tonames = False 
						End If
					Next
				End If
			End If
		End If
		Exit Function
	Else
		If a > 0 Then
			If CID(team.Group_Browse(10)) = 0 Then tonames = False 
			If DateDiff("h",b,Now())>a Then tonames = False 
		End If
	End if
End Function

Sub upext
	Dim Rs,nUser,ReList,Topic,Posttime,i
	Set Rs = team.execute("Select topic,ReList,UserName,Posttime From ["&IsForum&"Forum] Where deltopic=0 and CloseTopic=0 and ID="& Int(tID) )
	If Rs.Eof And Rs.Bof Then
		team.Error "ָ�������Ӳ����ڻ��Ѿ���ɾ����"
	Else	
		Topic = RS(0)
		ReList = Rs(1)
		nUser = Rs(2)
		Posttime = Rs(3)
	End If
	Rs.close:Set Rs = Nothing
	If Retopicid > 0 Then
		Set Rs = team.execute("Select UserName,Posttime From ["&IsForum & ReList &"] Where ID="& Int(retopicid))
		If Rs.Eof And Rs.Bof Then
			team.Error "ָ���Ļ��������ڻ��Ѿ���ɾ����"
		Else
			nUser = Rs(0)
			If Not tonames(CID(team.Forum_setting(60)),Rs(1)) Then
				team.Error "�����ӷ����ʱ�򳬹�"&team.Forum_setting(60)&"Сʱ���޷��������ֲ�����"
			End if
		End If
	Else
		If Not tonames(CID(team.Forum_setting(60)),Posttime) Then
			team.Error "�����ӷ����ʱ�򳬹�"&team.Forum_setting(60)&"Сʱ���޷��������ֲ�����"
		End If
	End If
	If Trim(nUser) = Trim(tk_UserName)  Then
		team.Error "�����ܶ��Լ�����"
	End if
	x1 = "��������"
	x2 = "<a href=""Thread.asp?tid="&tid & IIF(Retopicid>0,"#"&Retopicid&"","")&""">"&Topic&"</a>"
	Echo team.MenuTitle
	Echo "<form method=""post"" action=""?action=upextok&tid="&tID&"&retopicid="&retopicid&"""><input type='hidden' name='nName' value='" & nUser & "'/><input type='hidden' name='tid' value='" & tid & "'/><table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""80%"" align=""center"" class=""a2"">"
	Echo "<tr class=""tab1""><td colspan=""2""> �������� </td></tr>"
	Echo "<tr class=""a4""><td>����������</td><td>" & tk_UserName & " </td></tr>"
	Echo "<tr class=""a4""><td>��  �ߣ�</td><td>" & nUser & " </td></tr>"
	Echo "<tr class=""a4""><td>��  �⣺</td><td>" & Topic & " </td></tr>"
	Echo "<tr class=""a4""><td>��  �֣�</td><td> <select name=""score"" style=""width: 8em"">"
	Dim ExtCredits,ExtSort
	ExtCredits= Split(team.Club_Class(21),"|")
	ExtSort=Split(ExtCredits(CID(team.Forum_setting(46))),",")
	For i = -5 To 5
		If i = 0 Then 
			Echo "<option value="""&I&""" SELECTED>"&ExtSort(0) & I &"</option>"
		Else 
			Echo "<option value="""&I&""" >"&ExtSort(0) & I &"</option>"
		End if
	Next
	Echo "</select> *����Ϊ0ʱΪ��������</td></tr>"
	Echo "<tr class=""a4""><td valign=""top"">����ԭ��<BR>�������������ɲ��ܽ��в���<BR><BR><input type=""checkbox"" name=""sendreasonpm"" value=""1"" checked class=""radio""> ������Ϣ֪ͨ����<br /><input type=""checkbox"" name=""sendposted"" value=""1"" checked class=""radio""> ����������ʾ </td><td><textarea name=""reason"" style=""height: 8em; width: 25em""></textarea></td></tr>"
	Echo "</table><br><center><input type=""submit"" name=""ratesubmit"" value=""�� &nbsp; ��""></center></form>"
End Sub

Sub upextok
	Dim reason,score,nName,t,l,mytid,i,tmp,rs,tid
	TID = HRF(1,2,"tid")
	score = Request.Form("score")
	reason = HRF(1,1,"reason")
	nName = HRF(1,1,"nName")
	Dim ExtCredits,ExtSort
	ExtCredits= Split(team.Club_Class(21),"|")
	ExtSort=Split(ExtCredits(CID(team.Forum_setting(46))),",")
	If reason & "" ="" Then
		team.Error "�������ɲ���Ϊ��"
	End If
	If Not IsNumeric(score) Then 
		team.Error "��������"
	End if
	mytid = ""
	If retopicid > 0 Then
		mytid = retopicid
	Else
		mytid = tid
	End If
	If CID(Request.Cookies("Class")("UserpostExt"&mytid)) >0 Then
		team.Error "�����ܶ�һ�������ظ�����"
	End if
	If HRF(1,2,"sendreasonpm") = 1 Then
		'IntoMsg ������,������,����,����
		team.IntoMsg "ϵͳ��Ϣ",nName,"ϵͳ���ּ�¼"& vbcrlf &"�������ӱ��û�"&tk_UserName&" ���������ֲ���������"& ExtSort(0) & score & ExtSort(1) &" "& vbcrlf &"��������: "& reason &" "& vbcrlf &" [url=Thread.asp?tid="& TID &"][�����ֵ���������][/url]","���ּ�¼����"
	End If
	If HRF(1,2,"sendposted") = 1 Then
		If retopicid > 0 Then
			Set Rs = team.execute("select Content from ["&IsForum & team.Club_Class(11)&"] Where lock=0 and ID="& Int(retopicid))
			If Not Rs.eof Then
				If score = 0 Then
					tmp = ClearCode(team.Checkstr(Rs(0))) & "[fieldset]"&ReCode(Rs(0))&"<li><strong><a href=""Profile.asp?username="&tk_username&""">"&tk_username&"</a></strong> �� "& now &" <cite style=""color:red;"">�������ѣ�</cite><em title="""& reason &""">"& Left(reason,20) &"</em></li>[/fieldset]"
				Else
					tmp = ClearCode(team.Checkstr(Rs(0))) & "[fieldset]"&ReCode(Rs(0))&"<li><strong><a href=""Profile.asp?username="&tk_username&""">"&tk_username&"</a></strong> �� "& now &" <cite> ���飺"& ExtSort(0) &" "& IIF(score>0,"+"&score,score) &"</cite>:<em title="""& reason &""">"& Left(reason,20) &"</em></li>[/fieldset]"
				End If
				team.execute("Update ["&IsForum & team.Club_Class(11)&"] Set Content='"& tmp &"' Where lock=0 and ID="& Int(retopicid))
			End If
			Rs.Close: Set RS=Nothing
		Else
			Set Rs = team.execute("select Content from ["&IsForum&"Forum] Where deltopic=0 and CloseTopic=0 and ID="& Int(tid))
			If Not Rs.eof Then
				If score = 0 Then
					tmp = ClearCode(team.Checkstr(Rs(0))) & "[fieldset]"&ReCode(Rs(0))&"<li><strong><a href=""Profile.asp?username="&tk_username&""">"&tk_username&"</a></strong> �� "& now &"<cite style=""color:red;"">�������ѣ�</cite><em title="""& reason &""">"& Left(reason,20) &"</em></li>[/fieldset]"
				Else
					tmp = ClearCode(team.Checkstr(Rs(0))) & "[fieldset]"&ReCode(Rs(0))&"<li><strong><a href=""Profile.asp?username="&tk_username&""">"&tk_username&"</a></strong> �� "& now &"<cite> ���飺"& ExtSort(0) &" "& IIF(score>0,"+"&score,score) &"</cite>:<em title="""& reason &""">"& Left(reason,20) &"</em></li>[/fieldset]"
				End If 
				team.execute("Update ["&IsForum&"Forum] Set Content='"& tmp &"' Where deltopic=0 and CloseTopic=0 and ID="& Int(tid))
			End If
			Rs.Close: Set RS=Nothing
		End if
	End If
	If HRF(1,2,"sendreasonpm") = 0 And HRF(1,2,"sendposted")=0 Then
		team.Error "���Ż�ҳ��֪ͨ���߱�ѡһ"
	End if
	If score > 0 Then
		team.execute("Update ["&IsForum&"User] Set Extcredits"&CID(team.Forum_setting(46))&"=Extcredits"&CID(team.Forum_setting(46))&"+"&score&" Where UserName='"&nName&"'")
	Else
		team.execute("Update ["&IsForum&"User] Set Extcredits"&CID(team.Forum_setting(46))&"=Extcredits"&CID(team.Forum_setting(46))&"-"&abs(score)&" Where UserName='"&nName&"'")
	End If
	If CID(team.Forum_setting(62)) = 0 Then
		Response.Cookies("Class")("UserpostExt"&mytid) = 1
		Dim cookies_path_s,cookies_path_d,cookies_path
		cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
		cookies_path_d=ubound(cookies_path_s)
		cookies_path="/"
		For i=1 to cookies_path_d-1
			cookies_path=cookies_path&cookies_path_s(i)&"/"
		Next
		Response.Cookies("Class").Expires = DateAdd("s",360, Now())
		Response.Cookies("Class").Path = cookies_path
	End if
	team.error1 "<li>���ֳɹ����������û���"& ExtSort(0) & score & ExtSort(1) &"��</li><li> <a href=Thread.asp?tid="& tID &">��������</a><li><a href=""Default.asp"">������̳��ҳ</a><meta http-equiv=refresh content=3;url=Thread.asp?tid="& tID & ">"
End Sub

Function ClearCode(strContent)
	Dim re
	Set re = new RegExp
	re.IgnoreCase = True
	re.Global = True
	re.Pattern = "\[fieldset\]([\s\S]+?)\[\/fieldset]"
	strContent = re.Replace(strContent,"")
	set re = Nothing
	ClearCode = strContent
End Function

Function ReCode(s)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	If InStr(s,"[fieldset]")>0 Then
		re.Pattern = "([\s\S]+?)\[fieldset\]([\s\S]+?)\[\/fieldset]"
		ReCode=re.Replace(s,"$2")
	Else
		ReCode = ""
	End if
	set re=Nothing
End Function

Sub repotok
	Dim Rs,rUser,ReList,Topic,nTitle,nUser,i
	If Request.Form("postname") = "" Then
		team.Error "��û��ѡ����Ҫ����Ĺ�����Ա"
	End if
	rUser = Split(HtmlEncode(Replace(Request.Form("postname")," ","")),",")
	Set Rs = team.execute("Select topic,ReList,UserName From ["&IsForum&"Forum] Where deltopic=0 and CloseTopic=0 and ID="& Int(tID) )
	If Rs.Eof And Rs.Bof Then
		team.Error "ָ�������Ӳ����ڻ��Ѿ���ɾ����"
	Else	
		Topic = RS(0)
		ReList = Rs(1)
		nUser = Rs(2)
	End If
	Rs.close:Set Rs = Nothing
	nTitle = ""
	If retopicid > 0 Then
		Set Rs = team.execute("Select ID,UserName From ["&IsForum & ReList &"] Where ID="& Int(retopicid))
		If Rs.Eof Then
			team.Error "ָ���Ļ��������ڻ��Ѿ���ɾ����"
		Else
			nTitle = "�������ӣ�<a href=""Thread.asp?tid="&tid&"#"&retopicid&""" target=""_blank"">"& Topic &" </a><BR>�ظ��û���<a href=""Profile.asp?username="&Rs(1)&""">"&RS(1)&"</a>"
		End If
		Rs.close:Set Rs = Nothing
	Else
		nTitle = "�������ӣ�<a href=""Thread.asp?tid="&tid&""" target=""_blank"">"& Topic &" </a><BR>�����û���<a href=""Profile.asp?username="&nUser&""">"&nUser&"</a>"
	End if
	For i=0 To Ubound(rUser)
		If Not (Trim(tk_UserName) = Trim(rUser(i))) Then
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('"&tk_UserName&"','"&rUser(i)&"','��ã��ҷ����������ӵ����Ӳ�������̳����Ҫ��ϣ�������Դ����¡�<BR> "&nTitle&"',"&SqlNowString&",'�������ӱ���')")
			team.execute("Update ["&Isforum&"User] set Newmessage=Newmessage+1 where UserName='"&rUser(i)&"'")
		End if
	Next
	team.error1 "<li>���ӱ���ɹ���</li><li> <a href=Thread.asp?tid="& tID &">��������</a><li><a href=""Default.asp"">������̳��ҳ</a><meta http-equiv=refresh content=3;url=Thread.asp?tid="& tID & ">"
End Sub

Sub repot
	Dim Rs,tWhere,forumid,mMaster,i
	If team.Forum_setting(63)=0 Then
		team.Error "ϵͳ��ʱ�ر��˱������ӵĹ��ܡ��������Ա����վ�ڶ��Ų�ѯ��"
	End If
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""80%"" align=""center"" class=""a2"">"
	Echo "<tr class=""tab1""><td> TEAM's��ʾ </td></tr>"
	Echo "<tr class=""a4""><td> <B>���ӱ���</B> ���������й����û�����ѡȡ����Ҫ����Ĺ�����Ա��Ȼ���ڵ������һ�������������ӱ���Ĳ����� </td></tr>"
	Echo "</table><br>"
	Echo "<form method=""Post"" action=""?action=repotok&tid="&tid&"&retopicid="&retopicid&"""><table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""80%"" align=""center"" class=""a2"">"
	Echo "<tr class=""tab1""><td width=""10%""> ѡȡ </td><td width=""50%""> ������Ա���� </td><td width=""40%""> �û��ȼ� </td></tr>"
	Set Rs = team.execute("Select forumid From ["&IsForum&"Forum] Where deltopic=0 and CloseTopic=0 and ID="& Int(tID) )
	If Rs.Eof And Rs.Bof Then
		team.Error "ָ�������Ӳ����ڻ��Ѿ���ɾ����"
	Else	
		forumid = RS(0)
	End If
	Rs.close:Set Rs = Nothing
	mMaster = team.GroupManages
	If team.Forum_setting(63)= 1 Then
		tWhere = "UserGroupID=3"
	ElseIf team.Forum_setting(63)= 2 Then
		tWhere = "UserGroupID in (2,3)"
	ElseIf team.Forum_setting(63)= 3 Then
		tWhere = "UserGroupID in (1,2,3)"
	End if
	Set Rs = team.execute("Select UserName,Levelname,UserGroupID From ["&IsForum&"User] Where "& tWhere &" Order By UserGroupID Asc")
	Do While Not Rs.Eof
		If Int(Rs(2)) = 3 Then
			If IsArray(mMaster) Then
				For i = 0 To ubound(mMaster,2)
					If Trim(mMaster(1,i)) = Trim(Rs(0)) And (mMaster(2,i))=Int(forumid) Then
						Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""postname"" class=""radio"" value="""&Rs(0)&"""></td><td><a href=""Profile.asp?username="&Rs(0)&""">"&Rs(0)&"</a></td><td> "& Split(RS(1),"||")(0) &"  </td></tr>"
					End if
				Next 
			End If
		End if
		If Int(Rs(2))=1 Or Int(Rs(2))=2 Then
			Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""postname"" class=""radio"" value="""&Rs(0)&"""></td><td><a href=""Profile.asp?username="&Rs(0)&""">"&Rs(0)&"</a></td><td> "& Split(RS(1),"||")(0) &"  </td></tr>"
		End if
		Rs.MoveNext
	Loop
	Rs.close:Set Rs = Nothing
	Echo "</table><br><center><input type=""submit"" name=""submit"" value=""��һ��""></center><br></form>"
End Sub

%>
