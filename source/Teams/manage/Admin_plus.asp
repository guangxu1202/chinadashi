<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Public boards
Dim Admin_Class,Uid
Call Master_Us()
Uid = Cid(Request("uid"))
Header()
Admin_Class=",9,"
Call Master_Se()
team.SaveLog (" �������")
Select Case Request("action")
	Case "makeonline"
		Call makeonline
	Case "makeonlineok"
		Call makeonlineok
	Case "usermenuok"
		Call usermenuok
	Case "menuadd"
		Call menuadd
	Case "buyalipays"
		Call buyalipays
	Case "makebanks"
		Call makebanks
	Case Else
		Call usermenu
End Select

Sub makebanks
	Dim ho,newid,i,getvalue
	for each ho in request.form("setid")
		team.execute("Update ["&Isforum&"BankLog] Set Makes = 1,settime="&SqlNowString&" Where ID="&ho)
	next
	newid=Split(Replace(Request.Form("newid")," ",""),",")
	getvalue=Split(Replace(Request.Form("getvalue")," ",""),",")
	For i=0 To Ubound(newid)
		team.Execute("Update ["&Isforum&"User] set Extcredits"&Cid(team.Forum_setting(99))&"=Extcredits"&Cid(team.Forum_setting(99))&"+"&CID(getvalue(i))&",Newmessage=Newmessage+1 Where UserName='"&newid(i)&"' ")
		team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic,isbak) values ('"&TK_UserName&"','"&newid(i)&"','��ϲ����������Ļ��� [��"&CID(getvalue(i))&"] ���Ѿ����ʣ����½��[url=Control.asp?action=bank]����ת�˹���[/url]���鿴���Ļ�����',"&SqlNowString&",'���ֵ���֪ͨ',0)")
	Next
	SuccessMsg " ����������ɣ���ȴ�ϵͳ�Զ����ص� <a href=Admin_Plus.asp?action=buyalipays&makes=1>�鿴�Ѵ����� </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Plus.asp?action=buyalipays&makes=1>�� "
End Sub

Sub buyalipays %>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
 <table Cellpadding="5" Cellspacing="1" Border="0" Width="95%" class="a2">
    <tr class="a3">
      <td align="center"><a href="?action=buyalipays&makes=0">�鿴δ������</a></td>
	  <td align="center"><a href="?action=buyalipays&makes=1">�鿴�Ѵ�����</a></td>
	  <td colspan="5" align="center"> ��������.. </td>
    </tr>
</table>
<BR>
<Form Action="?action=makebanks" method="post">
<%
Dim Rs,ExtCredits
ExtCredits = Split(team.Club_Class(21),"|")
If CID(Request("makes")) = 0 Then
%>
  <table Cellpadding="5" Cellspacing="1" Border="0" Width="95%" class="a2">
    <tr class="a1">
      <td colspan="7" align="center">���ֹ��򶩵�[δ����]</td>
    </tr>
    <tr align="center" class="a3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio">
        ��������?</td>
      <td> ������ </td>
      <td> ������ </td>
      <td> ֧����� </td>
      <td> ������ </td>
	  <td> ����ʱ�� </td>
	  <td> ����ʱ�� </td>
    </tr>
	<%
	Set Rs = team.execute ("Select ID,bankname,buyname,buyvalue,getvalue,posttime,settime From ["&Isforum&"BankLog] Where Makes = 0 Order By posttime Desc ")
	Do While Not Rs.Eof
		Echo "<tr align=""center"" class=""tab4""> "
		Echo "	<td><input Name=""newid"" type=""hidden"" value="&RS(2)&"><input Name=""getvalue"" type=""hidden"" value="&RS(4)&"><input type=""checkbox"" name=""setid"" value="&RS(0)&" class=""radio""></td>"
		Echo "	<td>"&Rs(1)&"</td>"
		Echo "	<td>"&Rs(2)&"</td>"
		Echo "	<td>"&Rs(3)&" Ԫ/����� </td>"
		Echo "	<td>"&Rs(4)&"  "&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&"</td>"
		Echo "	<td>"&Rs(5)&"</td>"
		Echo "	<td>"&IIF(Rs(6)<>"",RS(6),"NA")&"</td>"
		Echo "</tr>"
		Rs.movenext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
  </table>
  <%Else%>
<table Cellpadding="5" Cellspacing="1" Border="0" Width="95%" class="a2">
    <tr class="a1">
      <td colspan="7" align="center">���ֹ��򶩵�[�Ѵ���]</td>
    </tr>
    <tr align="center" class="a3">
      <td> ����ID</td>
      <td> ������ </td>
      <td> ������ </td>
      <td> ֧����� </td>
      <td> ������ </td>
	  <td> ����ʱ�� </td>
	  <td> ����ʱ�� </td>
    </tr>
	<%
	Set Rs = team.execute ("Select ID,bankname,buyname,buyvalue,getvalue,posttime,settime From ["&Isforum&"BankLog] Where Makes = 1 Order By posttime Desc ")
	Do While Not Rs.Eof
		Echo "<tr align=""center"" class=""tab4""> "
		Echo "	<td> NO."&Rs(0)&"</td>"
		Echo "	<td>"&Rs(1)&"</td>"
		Echo "	<td>"&Rs(2)&"</td>"
		Echo "	<td>"&Rs(3)&" Ԫ/�����</td>"
		Echo "	<td>"&Rs(4)&"  "&Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&"</td>"
		Echo "	<td>"&Rs(5)&"</td>"
		Echo "	<td>"&IIF(Rs(6)<>"",RS(6),"NA")&"</td>"
		Echo "</tr>"
		Rs.movenext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
  </table>
  <%End if%>
  <BR>
  <center>
  <input type="Submit" value="����" name="forumlinksubmit">
</form>

<%
End Sub

Sub makeonline
	Dim regonline,maxonlies
	regonline = team.execute("Select count(*) from ["&isforum&"Online] where username<>''")(0)
	maxonlies = Application(CacheName&"_UserNum")-regonline		%>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>���������������Խ���̳��δ��½�û�����Ϊ������Ա��������̳������
      </ul>
      <ul>
        <li>ÿ�������������𳬹�100�������ͬʱ���������û�����ķѼ����ϵͳ��Դ�����·����������½� ��
      </ul>
      <ul>
        <li>ÿ�����������ò�ͬ��IP��ַ��ͷ��ʹϵͳ������ͬ��IP�Σ�����ʵ�������û���
      </ul></td>
  </tr>
</table>
<BR><BR>
<Form Action="?action=makeonlineok" method="post">
  <table Cellpadding="5" Cellspacing="1" Border="0" Width="95%" class="a2">
    <tr class="a1">
      <td colspan="2" align="center">������������</td>
    </tr>
    <tr class="a4">
		<td colspan="2" align="left" class="a2"><BR>
		 <ul>
		 <li>��̳�� <%=Application(CacheName&"_UserNum")%>�û������������û���  <%=regonline%> �����������ڿ�����������ֵ������ <%=maxonlies%> ��</li>
		 </ul>
		 </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8" width="60%">
		<B>����������Ҫ������������� --> </B>
		<BR> ÿ�β������û�����������𳬹�Ĭ�����ֵ��
	  </td>
      <td bgcolor="#FFFFFF">
		<input size="40" name="getname" value="<%=maxonlies%>">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8"><B>����������Ҫ���������IP�� --></B><br>
        IP����*�Ž���,ϵͳ���Զ������� 220. ��ͷ������IP��ַ .</td>
      <td bgcolor="#FFFFFF"><input size="40" name="getip" value="220.*.*.*">
      </td>
    </tr>
    <tr>
		<td bgcolor="#F8F8F8"><B>ѡ����Ӧ�İ��</B><br>��������ָ���İ��������Ӧ������ .</td>
		<td bgcolor="#FFFFFF">
			<select name="classid">
				<option value="0">��ҳ</option>
				<%BBSListshow(0)%>
			</Select>
			</td>
    </tr> 
  </table>
  <BR>
  <center>
  <input type="Submit" value="��ʼ����" name="forumlinksubmit">
</form>
<%
End Sub

Sub makeonlineok
	Dim getname,getip,i,Ismyip,rs,alluser,Ran,killname,classid,regonline
	Dim iswhere,Levelname,regalluser,rs1,u,Ran1,maxonlies
	regonline = team.execute("Select count(*) from ["&isforum&"Online] where username<>''")(0)
	maxonlies = Application(CacheName&"_UserNum")- regonline
	getname = CID(Request.Form("getname"))
	classid = CID(request.form("classid"))
	Getip = Request.Form("getip")
	If Getip="" or Instr(Getip,".")<=0 then 
		SuccessMsg " ip��ַ����"
	End if
	if getname > 100 then 
		SuccessMsg " ÿ�����������������뱣����100������, �Է�ϵͳ��Ϊͬʱ�����̫�������! "
	End if
	If getname> maxonlies Then
		SuccessMsg " ���������������ϵͳӵ�е��û���. "
	End if
	Set Rs = team.execute("Select UserName from ["&isforum&"online] Where Eremite=0 ")
	If Not (Rs.eof And Rs.bof) then
		Regalluser = Rs.GetRows(-1)
	End If
	Rs.close:Set Rs = Nothing
	iswhere = ""
	if isarray(regalluser) then
		for u=0 to ubound(regalluser,2)
			iswhere = iswhere & " And Not ( UserName='"&trim(Regalluser(0,u))&"')  "
		next
	end if
	if classid = 0  then
		Ran  = "��ҳ"
		Ran1 = "/default.asp?"
	Else
		Ran = team.execute("select bbsname from ["&IsForum&"Bbsconfig] where id="& classid )(0)
		Ran1 = "/Forums.asp?tid="&classid&""
	end if
	Set Rs=team.execute("Select ID,UserName,Levelname from ["&isforum&"User] Where UserGroupID >4 "& iswhere &" Order By Landtime Desc")
	If Not (Rs.eof And Rs.bof) then
		alluser = Rs.GetRows(-1)
	end if
	rs.close:set rs=nothing
	Ismyip = split(getip,".")
	if isarray(alluser) then
		For i=0 to ubound(alluser,2)
			randomize
			Levelname = Split(alluser(2,i),"||")
			team.Execute("Insert Into ["&IsForum&"Online](Forumid,Sessionid,Username,Ip,Eremite,Bbsname,Act,Acturl,Cometime,Lasttime,Levelname) Values ('"&CID(classid)&"','"&alluser(0,i)&"','"&alluser(1,i)&"','"&Ismyip(0)&"."&CInt(Rnd * 253)+1&"."&CInt(Rnd * 253)+1&"."&CInt(Rnd * 253)+1&"','0','"& team.Club_Class(1) &"','"&Ran&"','"&Ran1&"',"&SqlNowString & "," & SqlNowString & ",'"&Levelname(0)&"')" )
			if i > getname then exit for
		next
	end If
	Cache.DelCache("ShowLines"&CID(classid))
	Cache.DelCache("UserOnlineCache")
	SuccessMsg " ���������������, �ɹ������� "&getname&" ���ˡ�"
End Sub

Sub menuadd
	Dim Rs
	If Request("edit") = 1 then
		Set Rs=team.execute("Select Name,url,followid,SortNum,Newtype From ["&Isforum&"Menu] Where ID="&UID) 
		If Rs.Eof then
			SuccessMsg "  ��������"
		Else
%>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" name="settings" action="?action=usermenuok&updates=1">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <input type="hidden" name="uid" value="<%=uid%>">
    <tr align="center" class="a1">
      <td>&nbsp;</td>
      <td> ��ʾ˳�� </td>
      <td> ���� </td>
      <td> URL </td>
      <td> ���� </td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td></td>
      <td><input type="text" size="3" name="newid" value="<%=RS(3)%>"></td>
      <td><input type="text" size="15" name="newname"  value="<%=RS(0)%>"></td>
      <td><input type="text" size="15" name="newurl"  value="<%=RS(1)%>"></td>
      <td>
	  <input type="radio" name="newtype" value="1" <%If CID(Rs(4))=1 then%>checked<%End if%> class="radio">ǰ̨�˵� &nbsp; &nbsp;
      <input type="radio" name="newtype" value="0" <%If CID(Rs(4))=0 then%>checked<%End if%> class="radio"> ��̨�˵� </td>
    </tr>
  </table>
  <BR>
  <center>
    <input type="submit" name="forumlinksubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%		End if
	Else
%>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" name="settings" action="?action=usermenuok&newsinto=1">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <input type="hidden" name="fid" value="<%=request("fid")%>">
    <input type="hidden" name="Mid" value="<%=request("Mid")%>">
    <tr align="center" class="a1">
      <td>&nbsp;</td>
      <td> ��ʾ˳�� </td>
      <td> ���� </td>
      <td> URL </td>
      <td> ���� </td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td>����:</td>
      <td><input type="text" size="3" name="newid" value="0"></td>
      <td><input type="text" size="15" name="newname"></td>
      <td><input type="text" size="15" name="newurl"></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <BR>
  <center>
    <input type="submit" name="forumlinksubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
	End if
End Sub

Sub usermenuok
	Dim ho,newid,i,Myid,MySortNum
	If Request("updates") = 1 then
		If Request.Form("newname")&""=""  or Request.Form("newurl")&""="" then
			SuccessMsg "  ����д��Ҫ��������"
		Else
			team.execute("Update ["&Isforum&"Menu] set Name='"&Replace(Request.Form("newname"),"'","")&"',url='"&Replace(Request.Form("newurl"),"'","")&"',SortNum="&Cid(Request.Form("newid"))&",Newtype="&Cid(Request.Form("newtype"))&" Where ID="&UID)
		End if
	ElseIf Request("newsinto")=1 then
		If Request.Form("newname")&""=""  or Request.Form("newurl")&""="" then
			SuccessMsg "  ����д��Ҫ��������"
		Else
			If Request.form("Mid") = 0 then
				Newid = 0
			Else
				Newid = 1
			End if
			team.execute("insert into ["&Isforum&"Menu] (Name,url,followid,SortNum,Newtype) values ('"&Replace(Request.Form("newname"),"'","")&"','"&Replace(Request.Form("newurl"),"'","")&"',"&CID(Request("fid"))&","&Cid(Request.Form("newid"))&","&Newid&") ")
		End if
	Else
		for each ho in request.form("deleteid")
			team.execute("Delete from ["&Isforum&"Menu] Where ID="&ho)
		next
		If Request.form("deleteid")="" then
			Myid=Split(Request.Form("UID"),",")
			MySortNum=Split(Request.Form("SortNum"),",")
			For i=0 to Ubound(Myid)
				team.Execute("Update "&IsForum&"Menu set SortNum="&MySortNum(i)&" where ID="&Myid(i))
			Next
			If Request.Form("newname")<>"" then
				team.execute("insert into ["&Isforum&"Menu] (Name,url,followid,SortNum,Newtype) values ('"&Replace(Request.Form("newname"),"'","")&"','"&Replace(Request.Form("newurl"),"'","")&"',0,"&CID(Request.Form("newid"))&","&Cid(Request.Form("newtype"))&") ")
			End if
		End if
	End If
	Cache.DelCache("MenuLoad")
	SuccessMsg " �˵�������� ����ȴ�ϵͳ�Զ����ص� <a href=Admin_plus.asp>�˵�����</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_plus.asp>�� "
End Sub
Sub usermenu 

%>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>ǰ̨�˵�������Ӵ��࣬Ȼ����������¼����ࡣ����ֻ����Ϊ�˵�Ŀ¼���޷������ӡ�</li>
        <li>ǰ̨�˵�����ʾ�˲˵����ڵ����˵�����ʾ��</li>
        <li>��̨�˵�����ʾ�ں�̨��Ӳ˵��������Զ����������ļ����ӡ�</li>
		<li>ǰ̨�Ĳ˵���ʾ��Ҫ����������õ� <a href="http://localhost/1/Manage/Admincp.asp#��������ʾ��ʽ"><B>��������ʾ��ʽ</B></a>������ <B>��ʾ�Զ��������˵�</B>�� </li>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" name="settings" action="?action=usermenuok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="5">��ӵ����˵�</td>
    </tr>
    <tr align="center" class="a3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio">
        ɾ?</td>
      <td> ��ʾ˳�� </td>
      <td> ���� </td>
      <td> URL </td>
      <td> ���� </td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td>����[����]:</td>
      <td><input type="text" size="3" name="newid" value="0"></td>
      <td><input type="text" size="15" name="newname"></td>
      <td><input type="text" size="15" name="newurl"></td>
      <td>
	  <input type="radio" name="newtype" value="1" checked class="radio">ǰ̨�˵� &nbsp; &nbsp;
      <input type="radio" name="newtype" value="0" class="radio">��̨�˵� </td>
    </tr>
  </table>
  <BR>
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td> �༭�����˵� </td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td><%Call Menus %></td>
    </tr>
  </table>
  <BR>
  <center>
    <input type="submit" name="forumlinksubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
End Sub


Sub Menus()
	Dim SQL,RS,tmp
	Echo " <table cellspacing=""1"" cellpadding=""4"" width=""98%"" align=""center"" class=""a2""> "
	Set Rs=team.Execute("Select ID,Name,url,followid,SortNum,Newtype From "&IsForum&"Menu Where Followid=0 Order By SortNum")
	If Rs.Eof then
		Echo "<tr><td><BR><ul><center> Ŀǰû������κβ˵� </center></ul></td></tr> "
	End if
	Do While Not RS.Eof
		Echo "<tr class=""a4"" align=""center""><td width=""10""> <Input Name=UID value="&RS(0)&" type=hidden> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&" class=""checkbox""></td><td width=""100""> ����  <input type=text name=SortNum Value="&RS(4)&" Size=""1""> </td><td width=""50%"" align=""left""> �� <a target=_blank href=../"&RS(2)&"><b>"&RS(1)&"</b></a> </td><td> "
		tmp =""
		If Cid(Rs(5))=1 then	
			Echo "ǰ̨�˵�"
			tmp = "<a href=""?action=menuadd&fid="&RS(0)&"&Mid="&Rs(5)&""" title=""��ӱ�������¼��˵�"">[���]</a> <a href=""?action=menuadd&uid="&RS(0)&"&edit=1&Mid="&Rs(5)&""" title=""�༭���˵�����"">[�༭]</a>"
		Else
			Echo "��̨�˵�"
			tmp = "  <a href=""?action=menuadd&uid="&RS(0)&"&edit=1&Mid="&Rs(5)&""" title=""�༭���˵�����"">[�༭]</a> "
		End if
		Echo " </td><td> "&tmp&" </td></tr>"
		Call Menus_1(Rs(0))
		Echo " "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
	Echo "</table>"
End Sub

Sub Menus_1(a)
	Dim SQL,RS,Style,S,t,sty
	Set Rs=team.Execute("Select ID,Name,url,followid,SortNum,Newtype From "&IsForum&"Menu Where Followid="&a&" Order By SortNum")
	Do While Not RS.Eof
		Echo "<tr class=""a4"" align=""center""><td width=""10""> <Input Name=UID value="&RS(0)&" type=hidden> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&" class=""radio""></td><td width=""100""> ����  <input type=text name=SortNum Value="&RS(4)&" Size=""1""> </td><td width=""50%"" align=""left"">���� ��<a target=_blank href=../"&RS(2)&"><b>"&RS(1)&"</b></a> </td><td> "
		If Rs(5)=1 then	
			Echo "ǰ̨�˵�"
		Else
			Echo "��̨�˵�"
		End if
		Echo " </td><td> <a href=""?action=menuadd&uid="&RS(0)&"&edit=1&Mid="&Rs(5)&""" title=""�༭���˵�����"">[�༭]</a> </td></tr>"
		Call Menus_1(Rs(0))
		Echo " "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub

sub BBSListshow(selec)
	Dim SQL,ii,RS2,aa
	sql="Select ID,bbsname From ["&isforum&"bbsconfig] where followid="&selec&" order by SortNum"
	Set Rs2=team.Execute(sql)
	do while not rs2.eof
		aa="����"
		If selec = 0 then aa="��"
		Response.write "<option value="&rs2(0)&">"&aa&""&rs2(1)&"</option>"
		ii=ii+1
		BBSListshow rs2(0)
		ii=ii-1
		rs2.movenext
	loop
	Rs2.close: Set Rs2 = Nothing
End Sub

footer()
%>
