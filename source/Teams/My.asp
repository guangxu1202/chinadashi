<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim tID,fID,x1,x2
tID = HRF(2,2,"tid")
fID = HRF(2,2,"fid")
team.Headers(Team.Club_Class(1) & "- �ҵ��ղؼ�")

Select Case HRF(2,1,"action")
	Case "getnews"
		Call getnews
	Case "upnews"
		Call upnews
	Case "myopen"
		Call myopen
	Case Else
		Call Main()
End Select

Sub myopen()
	Dim Rs,sRoot,i
	x1="<A href=My.asp>�ҵĿ�ݲ˵�</a> "
	Echo team.MenuTitle
	Echo "<div class=""left"" style=""width:20%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""a1""><td>�ҵĿ�ݲ˵�</td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp"">�ҵ�����</a></td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp?action=myopen"">�ҵĻظ�</a></td></tr>"
	Echo "</table>"
	Echo "</div>"
	Echo "<div style=""margin-left:5px;width:100%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""tab1""><td width=""40%"">����</td><td width=""15%"">��̳</td><td width=""25%"">��󷢱�</td><td width=""10%"">״̬</td></tr>"
	Set Rs=Team.Execute("Select Top 50 B.id,B.Topic,B.Lasttime,B.Views,B.Replies,B.forumid,B.Locktopic,B.CloseTopic,B.Goodtopic,U.Topicid,U.ID,B.LastText From "&IsForum & team.Club_Class(11)&" U Inner Join ["&IsForum&"Forum] B On U.Topicid=B.ID Where B.Auditing=0 and U.UserName='"&TK_UserName&"' and B.UserName<>'"&TK_UserName&"' and B.Deltopic=0 Order By B.Lasttime Desc")
	sRoot = team.BoardList
	Do While Not Rs.eof
		Echo "	<tr class=""tab4""><td align=""left""><a href=""thread.asp?tid="& Rs(0) &"#"& RS(10) &""" target=""_blank"">"& Rs(1) &"</a> (�鿴:"& RS(3) &"/�ظ�:"&RS(4)&")</td><td align=""left"">"
		For i = 0 To UBound(sRoot,2)
			If Int(sRoot(0,i)) = Int(RS(5)) Then Echo sRoot(1,i)
		Next
		Echo "</td><td align=""left"">"& Rs(2) &" by <a href=""Profile.asp?username="& Split(Rs(11),"$@$")(0) &""">"& Split(Rs(11),"$@$")(0) &"</a></td><td>"
		If Rs(6) = 1 Then
			Echo "�ر�"
		ElseIf Rs(7) = 1 Then
			Echo "����"
		ElseIf Rs(8) = 1 Then
			Echo "�ѱ��Ӿ�"
		End if
		Echo "</td></tr>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo "</table>"
	Echo "</div>"
End Sub



Sub Main()
	Dim Rs,sRoot,i
	x1="<A href=My.asp>�ҵĿ�ݲ˵�</a> "
	Echo team.MenuTitle
	Echo "<div class=""left"" style=""width:20%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""a1""><td>�ҵĿ�ݲ˵�</td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp"">�ҵ�����</a></td></tr>"
	Echo "	<tr class=""a4""><td><a href=""my.asp?action=myopen"">�ҵĻظ�</a></td></tr>"
	Echo "</table>"
	Echo "</div>"
	Echo "<div style=""margin-left:5px;width:100%;"">"
	Echo "<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" align=""center"" class=""a2"">"
	Echo "	<tr class=""tab1""><td width=""40%"">����</td><td width=""15%"">��̳</td><td width=""25%"">��󷢱�</td><td width=""10%"">״̬</td></tr>"
	Set Rs=Team.Execute("Select top 50 ID,Topic,Views,Replies,forumid,Lasttime,Locktopic,CloseTopic,Goodtopic,LastText From ["&IsForum&"Forum] Where Auditing=0 and deltopic=0 and UserName='"& tk_UserName &"' Order By Lasttime Desc")
	sRoot = team.BoardList
	Do While Not Rs.eof
		Echo "	<tr class=""tab4""><td align=""left""><a href=""thread.asp?tid="& Rs(0) &""" target=""_blank"">"& Rs(1) &"</a> (�鿴:"& RS(2) &"/�ظ�:"&RS(3)&")</td><td align=""left"">"
		For i = 0 To UBound(sRoot,2)
			If Int(sRoot(0,i)) = Int(RS(4)) Then Echo sRoot(1,i)
		Next
		Echo "</td><td align=""left"">"& Rs(5) &" by <a href=""Profile.asp?username="& Split(Rs(9),"$@$")(0) &""">"& Split(Rs(9),"$@$")(0) &"</a></td><td>"
		If Rs(6) = 1 Then
			Echo "�ر�"
		ElseIf Rs(7) = 1 Then
			Echo "����"
		ElseIf Rs(8) = 1 Then
			Echo "�ѱ��Ӿ�"
		End if
		Echo "</td></tr>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo "</table>"
	Echo "</div>"
End Sub

Sub getnews
	Dim Rs
	If team.UserLoginED Then
		If tid = 0 Then
			team.Error "���Ӳ�������"
		Else
			Set Rs = team.execute("select topic From ["& isforum &"Forum] Where id="& tid)
			If Rs.eof Then
				team.Error "�����ڵ�����"
			Else
				If team.execute("select * From ["& isforum &"Favorites] Where ispub="& tid & " and username='"& tk_UserName &"' ").eof Then
					team.Execute("insert into ["&Isforum&"Favorites] (username,name,url,addtime,tag,ispub,look) values ('"& tk_UserName &"','"& Rs(0) &"','thread.asp?tid="& tid &"',"&SqlNowString&",'',"& tid &",0)")
					team.Error1 "��ѡ��������Ѿ��ɹ��ղ�,��ȴ�ϵͳ����.<meta http-equiv=refresh content=3;url="""& Request.ServerVariables("http_referer") &""">"
				Else
					team.Error1 "����ȥ�Ѿ��ղع��������,��ȴ�ϵͳ����.<meta http-equiv=refresh content=3;url="""& Request.ServerVariables("http_referer") &""">"
				End If 
			End If
		End If
	Else
		team.Error "����δ<a href=login.asp>��¼</a>��̳"
	End if
End Sub

team.footer
%>
