<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim Thing,myreforum,isre
Call Master_Us()
Header()
Dim Admin_Class
Admin_Class=",4,"
Call Master_Se()
team.SaveLog ("������̳ͳ�� [������������̳ͳ�� ] ")
Myreforum = Request.form("Reforum")
Select Case Request("menu")
	Case "UP_clear"
		Application.Contents.RemoveAll()
		SuccessMsg("ϵͳ�����Ѿ����ؽ�!")
	Case "updateids"
		Dim eid,oid
		Eid = Request.Form("u1")
		oid = Request.Form("u2")
		For i = Eid To oID
			Set Rs=team.execute("Select Count(ID) From Reforum Where TopicID="& i )
			If not Rs.Eof Then
				team.execute("Update ["&IsForum&"Forum] Set replies = "& Rs(0) &" Where ID="& i )
			End If
			Rs.Close:Set Rs=Nothing
		Next 
		SuccessMsg " ������ɣ� "
	Case "update1"
		Dim SQL,SQL1,SQL2,oldday
		SQL2 = "update ["&IsForum&"ClubConfig] set "
		If Request.Form("u1")=1 Then
			Set Rs=team.Execute("Select ID From ["&IsForum&"bbsconfig]")
			Do While not rs.eof
				SQL=Team.Execute("Select count(*) from ["&IsForum&"forum] where deltopic=0 and forumid="&rs(0))(0)
				If SQL>0 Then
					SQL1=Team.Execute("Select sum(replies) from ["&IsForum&"forum] where deltopic=0 and forumid="&rs(0))(0)
				Else
					SQL1=0
				End If
				Team.Execute("update ["&IsForum&"bbsconfig] set toltopic="&SQL&",tolrestore="&SQL+SQL1&" where ID="& rs(0))
				rs.movenext
			Loop
			SQL2 = SQL2 & "PostNum="&team.execute("select count(id) from "&IsForum&"forum where deltopic=0")(0)&","
			SQL2 = SQL2 & "RepostNum="&team.execute("select sum(replies) from "&IsForum&"forum where deltopic=0")(0)&","
		End If
		If Request.Form("u2")=1 Then
			Dim UserNum
			UserNum = team.Execute("Select count(id) from ["&IsForum&"user]")(0)
			SQL2 = SQL2 & "UserNum="&UserNum&","
		End If
		If Request.Form("u3")=1 Then
			If IsSqlDataBase = 1 Then
				today = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff(d,Posttime,"&SqlNowString&")=0 ")(0)
			Else
				today = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff('d',Posttime,"&SqlNowString&")=0 ")(0)
			End If
			If IsSqlDataBase = 1 Then
				today = int(today) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff(d,Posttime,"&SqlNowString&")=0")(0)
			Else
				today = int(today) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff('d',Posttime,"&SqlNowString&")=0")(0)
			End If
			SQL2 = SQL2 & "Today="&Today&""
		End If
		If Request.Form("u4")=1 Then
			If IsSqlDataBase = 1 Then
				Oldday = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff(d,Posttime,"&SqlNowString&")=1 ")(0)
				Oldday = int(Oldday) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff(d,Posttime,"&SqlNowString&")=1")(0)
			Else
				Oldday = team.execute("select count(id) from ["&IsForum&"forum] where deltopic=0 and datediff('d',Posttime,"&SqlNowString&")=1 ")(0)
				Oldday = int(Oldday) + team.execute("select count(id) from ["&IsForum&""&myreforum&"] where datediff('d',Posttime,"&SqlNowString&")=1")(0)
			End If
			SQL2 = SQL2 & ",Oldday="&Oldday&""
		End If
		If Request.Form("u5")=1 Then
			SQL2 = SQL2 & ",newreguser='"&team.Execute("Select Top 1 UserName From ["&IsForum&"User] Order by regtime Desc")(0)&"'"
		End If
		Team.Execute(SQL2)
		Application.Contents.RemoveAll()
		SuccessMsg("����̳����ͳ�����ݸ��³ɹ�!")
	Case "upnew"
		dim uid,toltopic,tolrestore,rs1,rs,today,trs,rs3,ismytoday,p
		uid=int(request("uid"))
		toltopic=0:tolrestore=0
		Set Rs=team.execute("select toltopic,tolrestore from ["&IsForum&"bbsconfig] where id="& uid)
		If Rs.Eof Then
			SuccessMsg("����Ĳ���")
		Else
			p =0
			Set rs3 = team.execute("select ID from "&IsForum&"forum where deltopic=0 and forumid="& uid)
			do while not rs3.eof
				If IsSqlDataBase = 1 Then
					ismytoday = team.execute("select count(ID) from "&IsForum&""&team.Club_Class(11)&" where datediff(d,Posttime,"&SqlNowString&")=0 and topicid="& RS3(0))(0)
				else
					ismytoday = team.execute("select count(ID) from "&IsForum&""&team.Club_Class(11)&" where datediff('d',Posttime,"&SqlNowString&")=0 and topicid="& RS3(0))(0)
				end if
				p= p+ismytoday
				rs3.movenext
			loop
			toltopic = team.execute("select count(*) from ["&IsForum&"forum] where deltopic=0 and forumid="& uid)(0)
			tolrestore = team.execute("select sum(replies) from ["&IsForum&"forum] where deltopic=0 and forumid="& uid)(0)
			if tolrestore="" or not isNumeric(tolrestore) then tolrestore=0
			Team.Execute("update ["&IsForum&"bbsconfig] set today="&p&",toltopic="&toltopic&",tolrestore="&tolrestore&" where id="& uid)
			Set Trs=team.Execute("Select Top 1 id,topic,username,lasttime From ["&IsForum&"Forum] Where deltopic=0 and forumid="& uid&" Order by lasttime Desc")
			If Not (Trs.Eof And Trs.Bof) Then
				Team.Execute("update ["&IsForum&"bbsconfig] set Board_Last='<a href=Thread.asp?tid="&Trs(0)&">"&Trs(1)&"</a>$@$"&Trs(2)&"$@$"&Trs(3)&"' where id="& uid)
			Else
				Team.Execute("update ["&IsForum&"bbsconfig] set Board_Last='��������$@$ - $@$"&SqlNowString&"' where id="& uid)
			End If
		End if
		Application.Contents.RemoveAll()
		SuccessMsg("ͳ�����ݸ��³ɹ�. ������"&toltopic&",������"&tolrestore&",��������"&p&" .")
	Case Else
		UP_main
		Footer()
End Select

Sub UP_main
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing=1 cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1">TEAM's ��ʾ</td>
  </tr>
  <tr>
    <td class="a4" height="30">
	<li>�����еĲ������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�
	<li>��������̳����  = �����¼���������̳����������ͻظ���,���������ȵ�ͳ�����ݣ�����ÿ��һ��ʱ������һ�Ρ�
	<li>���·ְ������� = ���ｫ���¼���ָ���������������ͻظ��������ظ���Ϣ�ȣ�����ÿ��һ��ʱ������һ�Ρ�
	</td>
  </tr>
  </table>
<br />

<form name="form1" method="POST" action="?menu=update1">
<table cellspacing="1" cellpadding="6" width="90%" border="0" class="a2" align="center">
 <tr class="a1"><td colspan="2">������̳������</td></tr>
 <tr class="a4">
    <td>
	<input type="checkbox" name="u1" value="1">����/������&nbsp;
	<input type="checkbox" name="u2" value="1">�û���&nbsp;
	<input type="checkbox" name="u3" value="1" checked>������&nbsp;
	<input type="checkbox" name="u4" value="1" checked>������&nbsp;
	<input type="checkbox" name="u5" value="1">���ע���û�&nbsp;
	</td></tr>
	<tr class=a4>
	<td>
	<input size="10" name="Reforum" value="Reforum">&nbsp;&nbsp;<input type="submit" name="Submit"value="������̳������"> <BR>ע:����д��ǰ�����������,��ѡ����Ҫ����ͳ�Ƶ�ѡ��<br> 
	</td>
  </tr>
</table><br /></form>

<form name="form1" method="POST" action="?menu=updateids">
<table cellspacing="1" cellpadding="6" width="90%" border="0" class="a2" align="center">
 <tr class="a1"><td colspan="2">���µ�����������</td></tr>
 <tr class="a4">
    <td>
	����д��Ҫ����ͳ�Ƶ�ID��  <input type="text" name="u1" value="1">&nbsp; &nbsp;<input type="text" name="u2" value="1">
	</td></tr>
	<tr class="a4">
	<td>
	<input type="submit" name="Submit"value="��ʼ����">
	</td>
  </tr>
</table><br /></form>

<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1">���·ְ�������</td>
  </tr>
	<tr class="a4">
	<td>
		<table cellspacing="1" cellpadding="5" width="99%" border="0" class="a2">
		<%ForumList(0)%>
		</table>
		</td>
  </tr>
</table><br><br>
<form name="form1" method="POST" action="?menu=UP_clear">
<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1" colspan="2">�ؽ�ϵͳ����</td>
  </tr>
  <tr>
    <td class="a4" colspan="2"><li>ϵͳ��ʹ����Application���� <%=Application.Contents.Count%> ����Session���� <%=Session.Contents.Count%> </td>
  </tr> 
	<%
For Each Thing in Application.Contents
	Response.Write "<tr class=a4><td>" & thing & "</td><td>״̬��"
	If isObject(Application.Contents(Thing)) Then
		Set Application.Contents(Thing) = Nothing
		Application.Contents(Thing) = null
		Response.Write "����ɹ��ر�"
	ElseIf isArray(Application.Contents(Thing)) Then
		Set Application.Contents(Thing) = Nothing
		Application.Contents(Thing) = null
		Response.Write "����ɹ��ͷ�"
	Else
		Response.Write Application.Contents(Thing)
		Application.Contents(Thing) = null
	End If
	Response.Write "</td></tr>"
Next
%></table><br>
<center><input type="submit" name="Submit1" value=" �ͷŻ��� "></form> <br>


<%
End sub

dim ii
ii=0
Sub ForumList(V)
	Dim SQL,RS,Style,S,T,Sty
	Set Rs=team.Execute("Select ID,BbsName,SortNum,Hide,Board_Model From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		Select Case RS(3)
			Case 1
				T="ֻ���ο�����"
			Case 2
				T="����"
			Case Else
				T="����"
		End Select
		If V=0 then	
			Response.Write"<tr class=a4><td width=""5%""></td><td><a target=_blank href=../Forums.asp?Fid="&RS(0)&">��<b>"&RS(1)&"</b></a>  </td><td> [״̬: <b>"&T&"</b>]</a>  </td><td> <a href=""?menu=upnew&uid="&RS(0)&""">����ͳ��</a></span></td></tr>"
		Else
			Response.Write"<tr class=a4><td width=""5%""></td><td>"&String(ii*2,"��")&" ��<a target=_blank href=../Forums.asp?Fid="&RS(0)&"><b>"&RS(1)&"</b></a> </td><td> [״̬: <b>"&T&"</b>]</a>  </td><td> <a href=""?menu=upnew&uid="&RS(0)&""">����ͳ��</a></span></td></tr>"
		End If
		ii=ii+1
		ForumList RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub
%>