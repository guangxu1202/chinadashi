<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<!-- #include file="../inc/MD5.asp" -->
<%
Dim Admin_Class,Uid
Call Master_Us()
Uid = Cid(Request("uid"))
Header()
Admin_Class=",6,"
Call Master_Se()
team.SaveLog ("�û����� [�������༭�û�������û� ���ϲ��û� ������û� �����ʹ��� ]  ")
Select Case Request("action")
	Case "adduser"
		Call adduser
	Case "adduserok"
		Call adduserok
	Case "setuser"
		Call Setuser
	Case "setuserok"
		Call setuserok
	Case "findmembers"
		Call findmembers
	Case "members"
		Call members
	Case "editgroups"
		Call editgroups
	Case "editgroupsok"
		Call editgroupsok
	Case "editcredits"
		Call editcredits
	Case "editcreditsok"
		Call editcreditsok
	Case "editmedals"
		Call editmedals
	Case "editmedalsok"
		Call editmedalsok
	Case "annonces"
		Call annonces
	Case "edituserexc"
		Call edituserexc
	Case "Activation"
		Call Activation
	Case "activationok"
		Call Activationok
	Case "getmoney"
		Call getmoney
	Case "getmoneyok"
		Call getmoneyok
	Case Else
		Call Master_Se()
		Call Main()
End Select

Sub getmoneyok
	Dim ho,newMembers,newWageMach,Gs
	NewMembers = Cid(Request.Form("newMembers"))
	NewWageMach = Cid(Request.Form("newWageMach"))
	for each ho in request.form("wagid")
		Team.execute("Delete from ["&Isforum&"Wages] Where id="&ho)
	next
	If Request.form("wagid")="" Then
		If NewMembers =""  Then SuccessMsg " �����Ʋ���Ϊ�ա�"
		If team.execute("Select Members from ["&Isforum&"Wages] where Members='"&NewMembers&"' ").Eof Then
			Set Gs = team.execute("Select ID,GroupName From ["&IsForum&"UserGroup] Where ID = "& NewMembers)
			If Not Gs.Eof Then
				team.execute("insert into ["&Isforum&"Wages] (Members,WageMach,WageGroupID) values ('"&Gs(1)&"',"&NewWageMach&","&Gs(0)&") ")
			End if
			Gs.Close:Set Gs=Nothing
		Else
			SuccessMsg  " ���û����Ѿ�����! "
		End If
	End if
	SuccessMsg  " ����ͼ��������ɡ� "
End Sub

Sub getmoney 
	%>
	<br>
	<br>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?action=getmoneyok">
	<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr class="a1">
	 <td>������ʾ</td>
	</tr>
	<tr class="a4">
	 <td><br>
      <ul>
        <li>TEAM's ֧�ֶԸ��û��鷢�Ź��ʣ��˹�����Ҫ�ڻ���ѡ�����濪�����׻������� ��
		<li>�򿪴˹��ܺ�ϵͳ����ÿ�µĵ�һ���Զ����Ź��ʣ���������Ϊ���׻�������ֵ�����û��˻� ��
      </ul></td>
	 </tr>
	</table>
	<BR>
	<table cellspacing="1" cellpadding="3" border="0" width="95%" align="center" class="a2">
	<tr class="a1">
		<td align="center" colspan="3">���¹��ʶ�ȹ���</td>
	</tr>
	<tr class="a3">
		<td align="center" width="80"><input type="checkbox" name="chkall" onclick="checkall(this.form, 'wagid')" class="a3">ɾ? </td>
		<td align="center">�����</td><td align="center">���ʶ��</td>
	</tr>
	<tbody <%if team.Forum_setting(96)=0 Then%>disabled<%end if%>>
	<%
	Dim Rs,WsValue,m
	i = 0
	Set Rs= team.execute("Select ID,Members,WageMach,WageGroupID From ["&Isforum&"Wages]")
	If Rs.Eof Then
		Echo "<tr class=""a4"" align=""center""><td colspan=""3"">Ŀǰû����Ҫ�ķ��Ź��ʵ������</td></tr>"
	Else
		Do While Not Rs.Eof 
			i = i+1
			Echo "<tr class=""a4"" align=""center""> "
			Echo "	<td><input type=""checkbox"" name=""wagid"" value="""&Rs(0)&"""></td> <td bgcolor=""#F8F8F8""> "&Rs(1)&" </td><td bgcolor=""#FFFFFF"">"&Rs(2)&"</td></tr> "
			Rs.MoveNext
		Loop
	End if
	Rs.Close:Set Rs=Nothing	%>
	<tr><td colspan="3" class="a4" height="2"></td></tr>
	<tr class="a4" align="center">
		<td>����:</td>
		<td><select name="newMembers" style="width:100%">
		<option value=""> ��ѡ���û��� </option>
		<%
		Dim Gs
		Set Gs = team.execute("Select ID,GroupName From ["&IsForum&"UserGroup] Where ID<>5 and ID<>6 and ID<>7 and ID<>28 Order By ID DEsc")
		Do While Not Gs.Eof
			Echo "<option value="""&Gs(0)&""">"&Gs(1)&" </option> "
			Gs.MoveNext
		Loop
		Gs.Close:Set Gs=Nothing
		%>
		</select></td>
		<td><input type="text" name="newWageMach" size="20" value="100"></td>
	</tr>
	</tbody>
	</table>
	<br><center>
	<input type="submit" name="medalsubmit" value="�� ��" <%if team.Forum_setting(96)=0 Then%>disabled<%end if%>>
	</center></form>
<%
End Sub

Sub Activationok
	Dim Ho
	for each ho in request.form("checkuid")
		team.execute("Update ["&Isforum&"User] Set UserGroupID=27,Levelname='��Сһ�꼶||||||0||0',Members='ע���û�' Where ID="&ho)
	Next
	team.SaveLog ("����û�����")
	SuccessMsg " ����û������ɹ�����ȴ�ϵͳ�Զ����ص� <a href=Admin_User.asp?action=Activation>����û�  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_User.asp?action=Activation>�� "
End Sub

Sub Activation %>
	<br>
	<br>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?action=activationok">
	<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr class="a1">
	 <td>������ʾ</td>
	</tr>
	<tr class="a4">
	 <td><br>
      <ul>
        <li>TEAM's �ṩ��3�����û�ע����֤��ʽ������鿴 <a href="http://localhost/1/Manage/Admincp.asp#ע������ʿ���"> ע������ʿ��� ��</a>
		<li>������������˹���ˣ���ô����Ҫ�ֶ����û����м��� ��
      </ul></td>
	 </tr>
	</table>
	<BR>
	<table cellspacing="1" cellpadding="3" border="0" width="95%" align="center" class="a2">
	<tr class="tab1">
		<td width="15%"><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio"> ��� </td><td>������û��б�</td><td>Ŀǰ�ȼ�</td>
	</tr>
	<%
	Dim Rs
	Set Rs = team.execute("Select * From ["&IsForum&"User] Where UserGroupID=5")
	Do While Not Rs.Eof
		Echo "<tr class=""a4"">"
		Echo "	<td align=""center""><input type=""checkbox"" name=""checkuid"" value="""&RS("ID")&""" class=""radio""></td><td title=""����ɲ鿴�û�����ϸ����""><a href=""Admin_User.asp?action=editgroups&uid="&RS("ID")&""">"&Rs("UserName")&"</a></td><td>"&Split(Rs("Levelname"),"||")(0)&"</td>"
		Echo "</tr>"
		Rs.MoveNext
	Loop
	Echo "</table></table><br><center><input type=""submit"" name=""onlinesubmit"" value=""�� ��""></center></form>"
	Rs.Close:Set Rs=Nothing
End Sub

Sub editmedalsok
	Dim tmp,newid
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " �������� "
	Else
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			If CID(Request.Form("medals"&i))>0 Then
				If Request.Form("reason"&i)&"" = "" Then 
					SuccessMsg "�����������������ɡ�"
				End if
				tmp = tmp & Request.Form("medalname"&i) &"&&&"&Request.Form("reason"&i) & "$$$"
			End if
		Next
		team.Execute("Update ["&Isforum&"User] set Medals='"&tmp&"' Where ID= "& Uid )
	end if
	SuccessMsg " �û���ѫ��������ɡ� "
End Sub

Sub editmedals
	Dim Rs,Rs1,Medals,u,i,Medalsinfo,Medaltext,MyMedals
	Dim SetMys,SetMsg
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " �������� "
	Else
		Set Rs = team.execute("Select UserName,Medals From ["&Isforum&"User] Where ID="& Uid)
		If Rs.Eof Then
			SuccessMsg " ָ�����û������ڡ� "
		Else
			Set Rs1=team.execute("Select ID,MedalName,Medalimg From ["&Isforum&"Medals] Where MedalSet=1 Order By ID asc")
			If Rs1.Eof Then
				SuccessMsg " Ŀǰû�����õ�ѫ�£��뵽 <A HREF=""Admin_Change.asp?action=medals"">ѫ�±༭</A> �������趨���õ�ѫ�º��ٱ༭��  "
			Else
				MyMedals = Rs1.GetRows(-1)
			End If
			Rs1.Close:Set Rs1=Nothing
			Echo "<br><br> "
			Echo "<body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10"">"
			Echo "<form method=""post"" action=""?action=editmedalsok&uid="&UID&""">"
			Echo "<table cellspacing=""1"" cellpadding=""4"" width=""95%"" align=""center"" class=""a2"">"
			Echo "<tr class=""a1"">"
			Echo "	<td colspan=""4"">ѫ�±༭ - "&Rs(0)&"</td>"
			Echo "</tr>"
			Echo "<tr class=""a4"" align=""center"">"
			Echo "		<td>ѫ��ͼƬ</td><td>����</td><td>�����ѫ��</td><td>��ѫ����</td>"
			Echo "</tr>"
			If IsArray(MyMedals) Then
				For i = 0 To UBound(MyMedals,2)
					Echo "<tr align=""center""><Input Name=""newid"" type=""hidden"" value="""&MyMedals(0,i)&""">"
					Echo "	<td bgcolor=""#F8F8F8""><Input Name=""medalname"&i&""" type=""hidden"" value="""&MyMedals(2,i)&"""><img src=""../images/plus/"&MyMedals(2,i)&" "" align=""absmiddle""></td><td bgcolor=""#FFFFFF"">"&MyMedals(1,i)&"</td>" 
					SetMys = "" : SetMsg = ""
					If InStr(RS(1),"$$$")>0 Then
						Medals = Split(RS(1),"$$$")
						for U = 0 to ubound(Medals)-1
							Medalsinfo = Split(Medals(u),"&&&")
							If Trim(Medalsinfo(0)) = MyMedals(2,i) Then
								SetMys = "checked"
								SetMsg = Medalsinfo(1)
							End If
						Next
					End If
					Echo "<td bgcolor=""#F8F8F8""><input type=""checkbox"" name=""medals"&i&""" class=""radio"" value="""&MyMedals(0,i)&""" "&SetMys&"><td bgcolor=""#FFFFFF""><textarea name=""reason"&i&""" rows=""5"" cols=""30"">"& SetMsg &"</textarea></td></td></tr> "
				Next 
			End If
			Echo "</table><BR><br><center>"
			Echo "<input type=""submit"" name=""medalsubmit"" value=""�� ��"">"
			Echo "</center></form><br><br>"
		End if
	End if
	Rs.Close:Set Rs=Nothing
End Sub
Sub editcreditsok
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " �������� "
	Else
		Dim Exters,i
		For i = 0 to 7
			If i = 0 Then
				Exters = "Extcredits"&i&"="&Cid(Request.Form("extcreditsnew"&i&""))&""
			Else
				Exters = Exters & ",Extcredits"&i&"="&Cid(Request.Form("extcreditsnew"&i&""))&""
			End if
		Next
		team.execute("Update ["&Isforum&"User] Set "&Exters&" Where ID="& Uid)
	End if
	SuccessMsg " ����������� ��"
End Sub

Sub editcredits	
	Dim Gs,Value,i,m,Rs,u,UserInfo
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " �������� "
	Else
		Set Rs = team.execute("Select Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,UserName,LevelName From ["&Isforum&"User] Where ID="& Uid)
		If Rs.Eof Then
			SuccessMsg " ָ�����û������ڡ� "
		Else
%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>TEAM's ֧�ֶ��û� 8 ����չ���ֵ����ã�ֻ�б����õĻ��ֲ����������б༭��
		<li>���û��Ļ��֣�����Ϊ�������ͷ�Ϊ���� ��
      </ul></td>
  </tr>
</table>
<br>
<form name="input" method="post" action="?action=editcreditsok&uid=<%=UID%>">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="10">�༭�û����� - <%=Rs(8)%>(<%=Split(RS(9),"||")(0)%>)</td>
    </tr>
    <tr class="a3" align="center">
      <td width="14%">�û���ϸ����</td>
	  <%
Dim ExtCredits,ExtSort
ExtCredits= Split(team.Club_Class(21),"|")
For U=0 to Ubound(ExtCredits)
	ExtSort=Split(ExtCredits(U),",")
	Echo " <td bgcolor=""#F8F8F8""> "
	If ExtSort(3)="1" Then 
		Echo ExtSort(0)
	Else
		Echo " ExtCredits"&U&" "
	End if
	Echo " </td> "
Next 
Echo " </tr><tr align=""center"" class=""a4""><td bgcolor=""#F8F8F8""> N/A </td>"

For U=0 to Ubound(ExtCredits)
	ExtSort=Split(ExtCredits(U),",")
	Echo " <td bgcolor=""#F8F8F8"">  <input name=""extcreditsnew"&u&""" type=""text"" size=""3"" "
	If ExtSort(3)="1" Then 
		Echo " value="""&RS(u)&""""
	Else
		Echo " value=""N/A"" disabled "
	End if
	Echo " ></td> "
Next 
%>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="creditsubmit" value="�� ��">
  </center>
</form>
<%		End if
	End if
End Sub

Sub editgroupsok
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " �������� "
	Else
		Dim SQL,tmp,UserInfo
		If Ucase(Trim(Request.Form("olduser"))) <> Ucase(Trim(Request.Form("usernamenew"))) Then

			team.execute("Update ["&Isforum&"User] Set  UserName='"&Trim(Request.Form("usernamenew"))&"' Where ID="& Uid )
			Team.Execute("Update ["&Isforum&"Forum] Set UserName='"&Trim(Request.Form("usernamenew"))&"' Where UserName='"&Trim(Request.Form("olduser"))&"'")

			Set RS = team.execute("Select TableName from TableList")
			If Rs.Eof Then
				Team.Execute("Update ["&Isforum&"Reforum] Set UserName='"&Trim(Request.Form("usernamenew"))&"' Where UserName='"&Trim(Request.Form("olduser"))&"'")
			Else
				Do While Not Rs.Eof
					Team.Execute("Update ["&Isforum&""&RS(0)&"] Set UserName='"&Trim(Request.Form("usernamenew"))&"' Where UserName='"&Trim(Request.Form("olduser"))&"'")
					Rs.Movenext
				Loop
			End If
			Rs.Close:Set Rs=nothing
			Team.Execute("Update ["&Isforum&"message] Set author='"&Trim(Request.Form("usernamenew"))&"' Where author='"&Trim(Request.Form("olduser"))&"'")
			Team.Execute("Update ["&Isforum&"message] Set incept=''"&Trim(Request.Form("usernamenew"))&"' Where incept='"&Trim(Request.Form("olduser"))&"'")
			Team.Execute("Update ["&Isforum&"upfile] Set UserName='"&Trim(Request.Form("usernamenew"))&"' Where UserName='"&Trim(Request.Form("olduser"))&"'")
			tmp = "�û� "&Trim(Request.Form("olduser"))&" �Ѿ�����Ϊ "&Trim(Request.Form("usernamenew"))&" "
		End if
		If Request.Form("clearquestion") = 1 Then
			team.execute("Update ["&Isforum&"User] Set Question='',Answer=''  Where ID="& Uid)
		End if
		If not IsValidEmail(Request.Form("emailnew")) Then
			SuccessMsg "�ʼ���ַ����"
		Else
			Dim passwordnew,Us,UserGroupID
			passwordnew = Request.Form("passwordnew")
			UserInfo = Request.Form("qqnew") &"|"& Request.Form("icqnew") &"|"& Request.Form("yahoonew") &"|"& Request.Form("msnnew") &"|"& Request.Form("taobao") &"|"& Request.Form("alibuy")
			Set Us = team.execute("Select UserGroupID From ["&Isforum&"User] Where UserName='"& tk_userName &"'")
			If Not Us.Eof Then
				UserGroupID = Int(Us(0))
			End If 
			Us.Close:Set Us=Nothing
			If int(Request.Form("mygroups"))=1 And Trim(tk_userName)<>Trim(WebSuperAdmin) Then
				SuccessMsg "��û����������Ա��Ȩ��"
			ElseIf int(Request.Form("mygroups"))=2 Then
				If UserGroupID = 2 Then SuccessMsg "��û�������û�����ͬ�ȼ���Ȩ��"
			End If 
			If passwordnew <> "" Then
				team.execute("Update ["&Isforum&"User] Set  UserPass='"&MD5(Request.form("passwordnew"),16)&"',UserGroupID ="&Cid(Request.Form("mygroups"))&",Posttopic="&Cid(Request.Form("postsnew"))&",Postrevert="&Cid(Request.Form("postsre"))&",Goodtopic="&Cid(Request.Form("digestpostsnew"))&",Usermail='"&team.Checkstr(Request.Form("emailnew"))&"',Userhome='"&team.Checkstr(Request.Form("sitenew"))&"',Userface='"&team.Checkstr(Request.Form("avatarnew"))&"',UserCity='"&team.Checkstr(Request.Form("locationnew"))&"',UserSex="&Cid(Request.Form("gendernew"))&",Honor='"&team.Checkstr(Request.Form("honor"))&"',Birthday='"&Request.Form("bdaynew")&"',Sign='"&HtmlEncode(Request.Form("signaturenew"))&"',Degree="&Cid(Request.Form("totalnew"))&",RegIP='"&team.Checkstr(Request.Form("regipnew"))&"',Regtime='"&Request.Form("regdatenew")&"',Landtime='"&Request.Form("lastvisitnew")&"',UserInfo='"&team.Checkstr(UserInfo)&"' Where ID="& Uid )
			Else
				team.execute("Update ["&Isforum&"User] Set  UserGroupID = "&Cid(Request.Form("mygroups"))&",Posttopic="&Cid(Request.Form("postsnew"))&",Postrevert="&Cid(Request.Form("postsre"))&",Goodtopic="&Cid(Request.Form("digestpostsnew"))&",Usermail='"&team.Checkstr(Request.Form("emailnew"))&"',Userhome='"&team.Checkstr(Request.Form("sitenew"))&"',Userface='"&team.Checkstr(Request.Form("avatarnew"))&"',UserCity='"&team.Checkstr(Request.Form("locationnew"))&"',UserSex="&Cid(Request.Form("gendernew"))&",Honor='"&team.Checkstr(Request.Form("honor"))&"',Birthday='"&Request.Form("bdaynew")&"',Sign='"&HtmlEncode(Request.Form("signaturenew"))&"',Degree="&Cid(Request.Form("totalnew"))&",RegIP='"&team.Checkstr(Request.Form("regipnew"))&"',Regtime='"&Request.Form("regdatenew")&"',Landtime='"&Request.Form("lastvisitnew")&"',UserInfo='"&team.Checkstr(UserInfo)&"' Where ID="& Uid )
			End if
			If Cid(Request.Form("oldgroup")) <> Cid(Request.Form("mygroups")) Then
				Call SetUserMamdber(Cid(Request.Form("mygroups")),Uid)
			End if
		End if
		SuccessMsg tmp & "<BR> �û���Ϣ���³ɹ���"
	End if
End sub

Sub SetUserMamdber(s,m)
	Dim Rs
	Set Rs= team.execute("Select GroupName,UserColor,UserImg,rank,Members From ["&IsForum&"UserGroup] Where ID="& Int(s))
	If Rs.Eof Then
		SuccessMsg "�û�Ȩ�ޱ��𻵣������µ��롣"
	Else
		team.Execute("Update ["&Isforum&"User] Set Levelname='"&Rs(0)&"||"&Rs(1)&"||"&Rs(2)&"||"&Rs(3)&"||0',Members='"&Rs(4)&"' Where ID="& Int(m))
	End if
End Sub

Sub editgroups	
	Dim Gs,Value,i,m,Rs,u,UserInfo,uGID,Us,UserGroupID
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " �������� "
	Else
		Set Us = team.execute("Select UserGroupID From ["&Isforum&"User] Where UserName='"& tk_userName &"'")
		If Not Us.Eof Then
			UserGroupID = Int(Us(0))
		End If 
		Us.Close:Set Us=Nothing 
		Set Rs = team.execute("Select UserName,UserGroupID,Posttopic,Postrevert,Deltopic,Goodtopic,Usermail,Userhome,Userface,UserCity,UserSex,Honor,Birthday,Sign,Degree,RegIP,Regtime,Landtime,UserInfo,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members From ["&Isforum&"User] Where ID="& Uid)
		If Rs.Eof Then
			SuccessMsg " ָ�����û������ڡ� "
		Else
			If Int(Rs(1))=1 Then
				If Not (Trim(tk_userName) = Trim(WebSuperAdmin)) Then SuccessMsg "����Ա�ȼ��û�������ֻ�ܱ����õ���վ����Ա�޸�, ���ù���Ա���������Conn.asp,�޸������"
			Elseif Int(Rs(1))=2 Then
				If UserGroupID=2 Then  SuccessMsg "�������ں�̨�޸ĸ��ڻ�����һ���ȼ����û�����,�������Լ�������"
			End if
			UserInfo = Split(Rs(18),"|")
%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>�����Ҫ�����û�Ȩ�޵�����Ա�����Conn.asp �޸����� Const WebSuperAdmin = "admin"		'Ĭ�����ù���Ա, 
		WebSuperAdmin �������Ĺ���Ա�����ƣ� ֻ�����ó�ΪĬ�����ù���Ա���û����������������û�������Ա��Ȩ�ޡ�
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=editgroupsok&uid=<%=UID%>">
  <input type="hidden" value="<%=Rs(0)%>" name="olduser">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�û����������</td>
    </tr>
    <tr class="a4">
      <td><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr class="a4">
		  <input type="hidden" name="oldgroup" value="<%=Rs(1)%>"> <%
	Set Gs = team.execute("Select Max(ID) as ID,GroupName From ["&IsForum&"UserGroup] Group By GroupName")
	If Gs.Eof or Gs.Bof Then
		SuccessMsg " �û�Ȩ�ޱ�������,���ֶ������±�! "
	Else
		u=0
		Do While Not Gs.Eof 
			u = u+1
			Response.write "<td><input type=""radio"" name=""mygroups"" value="""&Gs(0)&""""
			If Rs(1) =  Gs(0) Then Response.write " checked "
			Response.write "> "&Gs(1)&" </td>"
			If U= 5 Then 
				Echo "</tr><tr>"
				U=0
			End If
			Gs.MoveNext
		Loop
	End If
	Gs.Close:Set Gs=Nothing%>
        </table></td>
    </tr>
  </table>
  <BR>
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�༭�û� [ <%=Rs(0)%> ] </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�û���:</b><br>
        <span class="a3">�粻���ر���Ҫ���벻Ҫ�޸��û���</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="usernamenew" value="<%=Rs(0)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������:</b><br>
        <span class="a3">�������������˴�������</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="passwordnew" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����û���ȫ����:</b><br>
        <span class="a3">ѡ���ǡ�������û���ȫ���ʣ����û�������Ҫ�ش�ȫ���ʼ��ɵ�¼��ѡ�񡰷�Ϊ���ı��û��İ�ȫ��������</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="clearquestion" value="1">
        �� &nbsp; &nbsp;
        <input type="radio" name="clearquestion" value="0" checked>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ͷ��:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="honor" value="<%=Rs(11)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�Ա�:</b></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="gendernew" value="1" <%If Rs(10)=1 Then%>checked<%End if%>>
        ��
        <input type="radio" name="gendernew" value="2" <%If Rs(10)=2 Then%>checked<%End if%>>
        Ů
        <input type="radio" name="gendernew" value="0" <%If Rs(10)=0 Then%>checked<%End if%>>
        ����</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>Email:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="emailnew" value="<%=Rs(6)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="postsnew" value="<%=Rs(2)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="postsre" value="<%=RS(3)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��������:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="digestpostsnew" value="<%=Rs(5)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������ʱ��(����):</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="totalnew" value="<%=Rs(14)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>ע�� IP:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="regipnew" value="<%=Rs(15)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>ע��ʱ��:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="regdatenew" value="<%=Rs(16)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�ϴη���:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="lastvisitnew" value="<%=Rs(17)%>">
      </td>
    </tr>
  </table>
  <br>
  <br>
  <%
'  0        1              2       3        4         5        6         7        8          9 
'UserName,UserGroupID,Posttopic,Postrevert,Deltopic,Goodtopic,Usermail,Userhome,Userface,UserCity,
'   10     11      12    13    14     15   16        17
'UserSex,Honor,Birthday,Sign,Degree,RegIP,Regtime,Landtime
',Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7

'UserInfo 0.qq 1.icq 2.yahoo 3.msn 4.taobao 5.alipay 
%>
  <a name="�û�����"></a>
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�û�����</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��ҳ:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="sitenew" value="<%=Rs(7)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>QQ:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="qqnew" value="<%=UserInfo(0)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>ICQ:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="icqnew" value="<%=UserInfo(1)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>Yahoo:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="yahoonew" value="<%=UserInfo(2)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>MSN:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="msnnew" value="<%=UserInfo(3)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�Ա�:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="taobao" value="<%=UserInfo(4)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>֧����:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="alibuy" value="<%=UserInfo(5)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="locationnew" value="<%=Rs(9)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="bdaynew" value="<%=Rs(12)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>ͷ��:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="avatarnew" value="<%=Rs(8)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>ǩ��:</b></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="signaturenew" cols="30" style="height:70;overflow-y:visible;"><%=Rs(13)%></textarea></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="searchsubmit" value="�༭" <%If Session("UserMember") = 2 and  Rs(1) = 1 Then%>disabled<%end if%>>
  </center>
</form>
<br>
<%
		End if
		Rs.close:Set Rs=Nothing
	End if
End Sub

Sub members
	Dim DelForm,ho,Rs,Rs1,UserName,LastSubject,SQL
	If Request.Form("deleteid") = "" Then
		SuccessMsg "��û��ѡ����Ҫɾ�����û�ID"
	Else
		for each ho in Request.Form("deleteid")
			UserName = team.execute("Select UserName from ["&Isforum&"User] Where ID="& ho)(0)
			Set Rs = team.execute("Select ReList From  ["&isforum&"Forum] Where UserName='"&UserName&"'")
			Do While Not Rs.Eof
				team.execute("Delete from ["&isforum & Rs(0) &"] Where UserName='"&UserName&"'")
				Rs.MoveNext
			Loop
			Rs.close:Set Rs=Nothing
			team.execute("Delete from ["&isforum & team.Club_Class(11) &"] Where UserName='"&UserName&"'")
			team.execute("Delete from ["&isforum&"Forum] Where UserName='"&UserName&"'")
			team.execute("Delete from ["&isforum&"Message] Where author='"&UserName&"' or incept ='"&UserName&"'")
			If IsSqlDataBase = 1 Then
				SQL = " Board_Last Like '%"&UserName&"%'"
			Else
				SQL = "InStr(1,LCase(Board_Last),LCase('$@$"&UserName&"$@$'),0)<>0"
			End if
			Set Rs = team.execute("Select ID From ["&isforum&"Bbsconfig] Where "& SQL&"")
			Do While Not Rs.Eof
				LastSubject = team.Execute("Select Max(ID) from ["&IsForum&"forum] where deltopic=0 and Forumid="& Rs(0))(0)
				Set Rs1=team.Execute("Select ID,topic,username,posttime from ["&IsForum&"forum] where deltopic=0 and id="& LastSubject )
				If Not Rs1.eof then
					team.execute("Update ["&IsForum&"bbsconfig] set Board_Last='<A href=Thread.asp?tid="&rs1(0)&" target=""_blank"">"&Cutstr(rs1(1),200)&"</a> ��$@$"&TK_UserName&"$@$"&Now()&"' where id="&RS(0))
				End If
				Rs.MoveNext
			Loop
			team.execute("Delete from ["&isforum&"User] Where ID="&ho)
			Cache.DelCache("BoardLists")
		next
		SuccessMsg " ָ�����û��Ѿ�ɾ����"
	End If
End Sub

Sub findmembers
	Dim username,lookperm,Looks,findpages,usermail,regip,usersign,maxpost,maxlogin,regtime
	Dim CountNum,Rs,PageSNum,i,NextSeach
	Dim tmp,u,ExtCredits,ExtSort
	PageSNum = Cid(Request.Form("findpages"))		'��ѯ�����ҳ
	UserName = HtmlEncode(Request.Form("username"))
	Lookperm = Trim(Replace(Request.form("lookperm")," ",""))
	NextSeach = 0
	If PageSNum = "" or Not isNumeric(PageSNum) Then PageSNum = 10
	tmp = "	Where "
	If UserName<>"" Then 
		If IsSqlDataBase = 1 Then
			tmp = tmp & "UserName Like '%"&UserName&"%' "
		Else
			tmp = tmp & " InStr(1,LCase(UserName),LCase('"&UserName&"'),0)<>0 "
		End if
		NextSeach = 1
	End If
	If Lookperm<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If Instr(Lookperm,",")>0 Then
			Looks = Split(Lookperm,",")
			For i=0 To Ubound(Looks)
				If i = 0 Then
					tmp = tmp &  " InStr(1,LCase(LevelName),LCase('"&Looks(i)&"'),0)>0 "
				Else
					tmp = tmp &  " or InStr(1,LCase(LevelName),LCase('"&Looks(i)&"'),0)>0 "
				End If
			Next
		Else
			tmp = tmp &  " InStr(1,LCase(LevelName),LCase('"&Lookperm&"'),0)>0  "
		End If
		NextSeach = 1
	End If
	If IsValidEmail(Request.Form("usermail")) Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If IsSqlDataBase = 1 Then
			tmp = tmp & " Usermail Like '%"&Request.Form("usermail")&"%' "	
		Else
			tmp = tmp & "InStr(1,LCase(Usermail),LCase('"&Request.Form("usermail")&"'),0)<>0"
		End if
		NextSeach = 1
	End if
	If Request.Form("regip")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If IsSqlDataBase = 1 Then
			tmp = tmp & " RegIP Like '%"&Request.Form("regip")&"%' "	
		Else
			tmp = tmp & "InStr(1,LCase(RegIP),LCase('"&Request.Form("regip")&"'),0)<>0"
		End if
		'tmp = tmp & " RegIP Like '%"&Request.Form("regip")&"%' "	
		NextSeach = 1
	End if
	If Request.Form("usersign")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If IsSqlDataBase = 1 Then
			tmp = tmp & " Sign Like '%"&Request.Form("usersign")&"%' "	
		Else
			tmp = tmp & "InStr(1,LCase(Sign),LCase('"&Request.Form("usersign")&"'),0)<>0"
		End if
		'tmp = tmp & " Sign Like '%"&Request.Form("usersign")&"%' "	
		NextSeach = 1
	End if
	If Request.Form("userlogin")="1" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If IsSqlDataBase=1 Then
			tmp = tmp & "  Datediff(d,LandTime, " & SqlNowString & ") =0 "
		Else
			tmp = tmp & "  Datediff('d',LandTime, " & SqlNowString & ")=0"
		End If
		NextSeach = 1
	End If
	If Request.Form("newuserreg")="1" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If IsSqlDataBase=1 Then
			tmp = tmp & "  Datediff(d,RegTime, " & SqlNowString & ") =0 "
		Else
			tmp = tmp & "  Datediff('d',RegTime, " & SqlNowString & ")=0"
		End If
		NextSeach = 1
	End If
	If Request.Form("maxpost")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If Request.Form("maxpost1") = 1 Then
			tmp = tmp & " Posttopic+Postrevert > "&Request.Form("maxpost")&" "	
		Else
			tmp = tmp & " Posttopic+Postrevert < "&Request.Form("maxpost")&" "	
		End if
		NextSeach = 1
	End if
	If Request.Form("maxlogin")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If Request.Form("maxlogin1") = 1 Then
			tmp = tmp & " Degree * 60 > "&Request.Form("maxlogin")&" "	
		Else
			tmp = tmp & " Degree * 60 < "&Request.Form("maxlogin")&" "	
		End if
		NextSeach = 1
	End if
	If Request.Form("regtime")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If Request.Form("regtime1") = 0 Then
			If IsSqlDataBase=1 Then
				tmp = tmp & "  Datediff(d, RegTime, " & Request.Form("regtime") & ") > 0 "
			Else
				tmp = tmp & "  Datediff('d', RegTime, " & Request.Form("regtime") & " ) > 0"
			End If
		ElseIf Request.Form("regtime1") = 1 Then
			If IsSqlDataBase=1 Then
				tmp = tmp & "  Datediff(d, RegTime, " & Request.Form("regtime") & ") = 0 "
			Else
				tmp = tmp & "  Datediff('d', RegTime, " & Request.Form("regtime") & " ) = 0"
			End If
		Else
			If IsSqlDataBase=1 Then
				tmp = tmp & "  Datediff(d, RegTime, " & Request.Form("regtime") & ") < 0 "
			Else
				tmp = tmp & "  Datediff('d',RegTime, " & Request.Form("regtime") & " ) < 0"
			End If
		End if
		NextSeach = 1
	End if
	If Request.Form("MyCred0")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If Request.Form("Nums0") = 1 Then
			tmp = tmp & " Extcredits0 > "&Request.Form("MyCred0")&" "	
		Else
			tmp = tmp & " Extcredits0 < "&Request.Form("MyCred0")&" "	
		End if
		NextSeach = 1
	End if
	If Request.Form("MyCred1")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If Request.Form("Nums1") = 1 Then
			tmp = tmp & " Extcredits1 > "&Request.Form("MyCred1")&" "	
		Else
			tmp = tmp & " Extcredits1 < "&Request.Form("MyCred1")&" "	
		End if
		NextSeach = 1
	End if
	If Request.Form("MyCred2")<>"" Then
		If NextSeach = 1 Then tmp = tmp & " and "
		If Request.Form("Nums2") = 1 Then
			tmp = tmp & " Extcredits2 > "&Request.Form("MyCred2")&" "	
		Else
			tmp = tmp & " Extcredits2 < "&Request.Form("MyCred2")&" "	
		End if
		NextSeach = 1
	End if
	If tmp = "	Where " Then tmp = ""
	Dim TopUser
	Set Rs=team.Execute("Select top "&PageSNum&" ID,UserName,UserGroupID,Levelname,Posttopic,Postrevert,Regtime,Landtime,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,UserMail From ["&Isforum&"User] "&tmp&" Order By ID Asc")
	TopUser = team.Execute("Select Count(ID) From ["&Isforum&"User] "&tmp&" ")(0)
	If Request.Form("searchsubmit") = "�����û�" Then
		Echo "<body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10"">"
		Echo "<table cellspacing=""1"" cellpadding=""4"" width=""95%"" align=""center"" class=""a2"">"
		Echo " <form method=""post"" action=""?action=members"">"
		Echo "<tr align=""center"" class=""a1"">"
		Echo "	<td width=""48""><input type=""checkbox"" name=""chkall"" onclick=""checkall(this.form, 'delete')"">ɾ?</td>"
		Echo "	<td>�û���</td>"
		Echo "	<td>������</td><td>�û���</td><td>ע��ʱ��</td><td>��½ʱ��</td><td>�༭</td></tr>"
		If Rs.Eof Then
			Echo "<tr align=""center"" class=""a4""><td Colspan=""7""> �Բ���û���ҵ������������û���</td></tr>"
		End if
		Do While Not Rs.Eof
			Echo " <tr align=""center"" class=""a4"">"
			Echo "		<td><input type=""checkbox"" name=""deleteid"" value="&RS(0)&"	"
			If Rs(2)>88 Then Echo "disabled"
			Echo "		></td>"
			Echo "		<td><a href=""../Profile.asp?username="&RS(1)&""" target=""_blank"">"&Rs(1)&"</a></td>"
			Echo "		<td>"&Rs(4)+Cid(Rs(5))&"</td>"
			Echo "		<td>"
			If Rs(2)>=77 Then Echo "<b>"
			Echo Split(Rs(3),"||")(0)
			If Rs(2)>=77 Then Echo "</b>"
			Echo "		</td>"
			Echo "		<td>"&RS(6)&"</td>"
			Echo "		<td>"&RS(7)&"</td>"
			Echo "		<td><a href=""?action=editgroups&uid="&RS(0)&""">[�û�����]</a> <a href=""?action=editcredits&uid="&RS(0)&""">[����]</a> <a href=""?action=editmedals&uid="&RS(0)&""">[ѫ��]</a> </td>"
			Echo "	</tr>"
			Rs.MoveNext
		Loop
		Echo "</table><br><center> <input type=""submit"" name=""searchsubmit"" value=""ɾ���û�""></center></form>"
	End If
	If Request.Form("newslettersubmit") = "��̳֪ͨ" Then
		Echo " <BR><BR><form method=""post"" action=""?action=annonces"">"
		Do While Not Rs.Eof
			Echo  "	<input type=""hidden"" name=""msgid"" value="""&RS(0)&""">"
		Rs.MoveNext
		Loop
		%>
		<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
		<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
		<tr class="a1"><td colspan="9">���������Ļ�Ա��: <%=TopUser%></td></tr>
		<tr>
			<td bgcolor="#F8F8F8">����:</td>
			<td bgcolor="#FFFFFF"><input type="text" name="subject" size="80" value=></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8" valign="top">����:</td><td bgcolor="#FFFFFF">
			<textarea cols="80" rows="10" name="message"></textarea></td></tr>
		<tr>
			<td bgcolor="#F8F8F8">���ͷ�ʽ:</td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="email" name="sendvia"> Email<input type="radio" value="pm" checked name="sendvia"> ����Ϣ</td>
		</tr>
		</table><br>
		<center><input type="submit" name="sendsubmit" value="�� ��"></center></form><br><br>
		<%
	End If
	If Request.Form("creditsubmit") = "���ֽ���" Then 
		Echo " <BR><BR><form method=""post"" action=""?action=edituserexc"">"
		Do While Not Rs.Eof
			Echo  "	<input type=""hidden"" name=""msgid"" value="""&RS(0)&""">"
			Rs.MoveNext
		Loop
	%>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	 <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
		<tr class="a1">
		  <td colspan="10">���������Ļ�Ա��: <%=TopUser%></td>
		</tr>
		<tr class="a3" align="center">
			<td width="14%">�û���ϸ����</td>
		<%
			ExtCredits= Split(team.Club_Class(21),"|")
			For U=0 to Ubound(ExtCredits)
				ExtSort=Split(ExtCredits(U),",")
				Echo " <td bgcolor=""#F8F8F8""> "
				If ExtSort(3)="1" Then 
					Echo ExtSort(0)
				Else
					Echo " ExtCredits"&U&" "
				End if
				Echo " </td> "
			Next 
			Echo " </tr><tr align=""center"" class=""a4""><td bgcolor=""#F8F8F8""> ������ֵ </td>"
			For U=0 to Ubound(ExtCredits)
				ExtSort=Split(ExtCredits(U),",")
				Echo " <td bgcolor=""#F8F8F8"">  <input name=""extcreditsnew"&u&""" type=""text"" size=""3"" "
				If ExtSort(3)="1" Then 
					Echo " value=""0"" "
				Else
					Echo " value=""N/A"" disabled "
				End if
				Echo " ></td> "
			Next %>
			</tr>
		</table>
	 <br>
	 <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
		<tr class="a1"><td colspan="9"><input class="a1" type="checkbox" name="sendcreditsletter" value="1">���ͻ��ֱ��֪ͨ</td></tr>
		<tr>
			<td bgcolor="#F8F8F8">����:</td>
			<td bgcolor="#FFFFFF"><input type="text" name="subject" size="80" value=></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8" valign="top">����:</td><td bgcolor="#FFFFFF">
			<textarea cols="80" rows="10" name="message"></textarea></td></tr>
		<tr>
			<td bgcolor="#F8F8F8">���ͷ�ʽ:</td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="email" name="sendvia"> Email<input type="radio" value="pm" checked name="sendvia"> ����Ϣ</td>
		</tr>
		</table><br>
	 <center><input type="submit" name="creditsubmit" value="�� ��"></center></form><%
	End If
	Rs.Close:Set Rs=Nothing
End Sub

Sub edituserexc
	If Request.Form("msgid") = "" Then
		SuccessMsg "  ��������Ҫ���͵��û���"
	Else
		Dim Exters,i,ho
		For i = 0 to 7
			If i = 0 Then
				Exters = "Extcredits"&i&"="&Cid(Request.Form("extcreditsnew"&i&""))&""
			Else
				Exters = Exters & ",Extcredits"&i&"="&Cid(Request.Form("extcreditsnew"&i&""))&""
			End if
		Next
		for each ho in Request.Form("msgid")
			team.execute("Update ["&Isforum&"User] Set "&Exters&" Where ID="&  ho)
		next
		If Request.Form("sendcreditsletter") = 1 Then
			If request.Form("sendvia") = "pm" Then
				for each ho in Request.Form("msgid")
					team.execute("Update ["&isforum&"User] Set Newmessage=Newmessage+1 Where ID="& ho)
					team.Execute( "insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('ϵͳ��Ϣ','"&team.Execute("Select UserName From ["&isforum&"User] Where ID="& ho)(0)&"','"&Request.Form("message")&"',"&SqlNowString&",'"&Request.Form("subject")&"')" )
				next
			Else
				for each ho in Request.Form("msgid")
					If IsValidEmail(team.Execute("Select UserMail From ["&isforum&"User] Where ID="& ho)(0)) Then
						Call Emailto ( team.Execute("Select UserMail From ["&isforum&"User] Where ID="& ho)(0), Request.Form("subject") , Request.Form("message"))
					End if
				next
			End if
		End if
	End if
	SuccessMsg " ����������ɣ���ȴ�ϵͳ�Զ����ص� <a href=Admin_User.asp>�༭�û� </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_User.asp> ��"
End Sub


Sub annonces
	Dim MsgName,ho,msgmail
	If Request.Form("msgid") = "" Then
		SuccessMsg "  ��������Ҫ���͵��û���"
	Else
		If Len(Request.Form("message"))<5 or Request.Form("subject") = "" Then 
			SuccessMsg " ���ݻ���ⲻ��Ϊ�� ��"
		Else
			If request.Form("sendvia") = "pm" Then
				for each ho in Request.Form("msgid")
					team.execute("Update ["&isforum&"User] Set Newmessage=Newmessage+1 Where ID="&ho)
					team.Execute( "insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('ϵͳ��Ϣ','"&team.Execute("Select UserName From ["&isforum&"User] Where ID="& ho)(0)&"','"&Request.Form("message")&"',"&SqlNowString&",'"&Request.Form("subject")&"')" )
				next
			Else
				for each ho in Request.Form("msgid")
					If IsValidEmail(team.Execute("Select UserMail From ["&isforum&"User] Where ID="& ho)(0)) Then
						Call Emailto ( team.Execute("Select UserMail From ["&isforum&"User] Where ID="& ho)(0), Request.Form("subject") , Request.Form("message"))
					End if
				next
			End if
		End if
		SuccessMsg " ��Ϣ�ѷ��ͳɹ�����ȴ�ϵͳ�Զ����ص� <a href=Admin_User.asp>�༭�û� </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_User.asp> ��"
	End If
End Sub

Sub setuserok
	Dim source1,source2,source3,target,UserTrg,rs,MsgTrg,MsgTrg1
	Source1 = HtmlEncode(Trim(Request.Form("source1")))
	Source2 = HtmlEncode(Trim(Request.Form("source2")))
	Source3 = HtmlEncode(Trim(Request.Form("source3")))
	Target = HtmlEncode(Trim(Request.Form("target")))
	If Source1 & Source2 & Source3 &"" = "" Then
		SuccessMsg "ԭ�û�����������Ҫһ�����ݣ�����ȫ��Ϊ�ա�"
	ElseIf Source1&"" = "" and Source2 & Source3 &"" <>"" Then 
		SuccessMsg "�û���������� <FONT COLOR=""red""><B>ԭ�û��� 1</B></FONT> ����ʼ��" 	
	Else	
		If Source1 & ""<>"" Then
			If team.execute("Select * from ["&Isforum&"User] Where UserName='"&Source1&"'").Eof Then
				SuccessMsg " ϵͳ��������Ϊ"&Source1&"���û��� ��" 
			End If
		End If
		If Source2 & ""<>"" Then
			If team.execute("Select * from ["&Isforum&"User] Where UserName='"&Source2&"'").Eof Then
				SuccessMsg " ϵͳ��������Ϊ"&Source2&"���û��� ��" 
			End if
		End If
		If Source3 & ""<>"" Then
			If team.execute("Select * from ["&Isforum&"User] Where UserName='"&Source3&"'").Eof Then
				SuccessMsg " ϵͳ��������Ϊ"&Source3&"���û��� ��" 	
			End If
		End If
		UserTrg = " UserName='"&Source1&"' "
		MsgTrg = " author = '"&Source1&"' "
		MsgTrg1 = " incept = '"&Source1&"' "
		If Source2&"" <>"" Then 
			UserTrg = UserTrg & " or UserName='"&Source2&"' "
			MsgTrg = MsgTrg & " or author = '"&Source2&"' "
			MsgTrg1 = MsgTrg1 & " or incept = '"&Source2&"' "
		End If
		If Source3&"" <>"" Then 
			UserTrg = UserTrg & " or UserName='"&Source3&"' "
			MsgTrg = MsgTrg & " or author = '"&Source3&"' "
			MsgTrg1 = MsgTrg1 & " or incept = '"&Source3&"' "
		End If
		Dim Uppost,UpRepost,GoodPost,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7
		Uppost=team.execute("Select Sum(Posttopic) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		UpRepost=team.execute("Select Sum(Postrevert) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits0=team.execute("Select Sum(Extcredits0) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits1=team.execute("Select Sum(Extcredits1) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits2=team.execute("Select Sum(Extcredits2) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits3=team.execute("Select Sum(Extcredits3) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits4=team.execute("Select Sum(Extcredits4) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits5=team.execute("Select Sum(Extcredits5) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits6=team.execute("Select Sum(Extcredits6) From ["&Isforum&"User] Where "&UserTrg&" ")(0)
		Extcredits7=team.execute("Select Sum(Extcredits7) From ["&Isforum&"User] Where "&UserTrg&" ")(0)

		Team.Execute("Update ["&Isforum&"User] Set Posttopic=Posttopic+"&Cid(Uppost)&",Postrevert=Postrevert+"&Cid(UpRepost)&",Goodtopic=Goodtopic+"&Cid(GoodPost)&",Extcredits0=Extcredits0+"&Cid(Extcredits0)&",Extcredits1=Extcredits1+"&Cid(Extcredits1)&",Extcredits2=Extcredits2+"&Cid(Extcredits2)&",Extcredits3=Extcredits3+"&Cid(Extcredits3)&",Extcredits4=Extcredits4+"&Cid(Extcredits4)&",Extcredits5=Extcredits5+"&Cid(Extcredits5)&",Extcredits6=Extcredits6+"&Cid(Extcredits6)&",Extcredits7=Extcredits7+"&Cid(Extcredits7)&" Where UserName='"&Target&"'")
		Team.Execute("Update ["&Isforum&"Forum] Set UserName='"&Target&"' Where "&UserTrg&" ")
		Set RS = team.execute("Select TableName from TableList ")
		If Rs.Eof Then
			Team.Execute("Update ["&Isforum&"Reforum] Set UserName='"&Target&"' Where "&UserTrg&" ")
		Else
			Do While Not Rs.Eof
				Team.Execute("Update ["&Isforum&""&RS(0)&"] Set UserName='"&Target&"' Where "&UserTrg&" ")
				Rs.Movenext
			Loop
		End If
		Rs.Close:Set Rs=nothing
		Team.Execute("Update ["&Isforum&"message] Set author='"&Target&"' Where  "&MsgTrg&" ")
		Team.Execute("Update ["&Isforum&"message] Set incept='"&Target&"' Where "&MsgTrg1&" ")
		Team.Execute("Update ["&Isforum&"upfile] Set UserName='"&Target&"' Where "&UserTrg&" ")
		If Trim(Source1) = TK_UserName  or Trim(Source2) = TK_UserName or Trim(Source3) = TK_UserName Then
			Response.Cookies(Forum_sn)("username")= Target
		End If
		'ɾ��ԭ�û�
		Team.Execute("Delete From ["&Isforum&"User] Where "&UserTrg&" ")
	End If
	SuccessMsg " �û��β��ɹ���ԭ�û������������֣���ȫ��ת��Ŀ���û���ͬʱԭ�û��ѱ�ɾ�� ��"
End Sub

Sub  Setuser%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?action=setuserok">
  <table cellspacing="1" cellpadding="4" width="85%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�ϲ��û� - ԭ�û������ӡ�����ȫ��ת��Ŀ���û���ͬʱɾ��ԭ�û�</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">ԭ�û��� 1:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="source1" size="20"></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">ԭ�û��� 2:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="source2" size="20"></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">ԭ�û��� 3:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="source3" size="20"></td>
    </tr>
    <tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">Ŀ���û���:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="target" size="20"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="mergesubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
End Sub



Sub adduserok
	Dim ExtCredits
	Dim newusername,newpassword,newemail,emailnotify,CheckStr,i
	NewUserName = HtmlEncode(Trim(Request.Form("newusername")))
	Newpassword = team.Checkstr(Trim(Request.Form("newpassword")))
	Newemail = Trim(Request.Form("newemail"))
	If NewUserName = "" Or IsNull(NewUserName) Then
		SuccessMsg "�û�������Ϊ��"
	End If
	If Newpassword = "" Or IsNull(Newpassword) Then
		SuccessMsg "���벻��Ϊ��"
	End If	
	If Not IsValidEmail(Newemail) Then
		SuccessMsg "�ʼ���ʽ���� !"
	End If
	CheckStr=Array("=","%",chr(32),"?","&",";",",","'",",",chr(34),chr(9),"��","$","|")
	For i=0 To Ubound(CheckStr)
		If Instr(NewUserName,CheckStr(i))>0 then SuccessMsg "�û����в��ܺ����������"
	Next
	If Not team.execute("Select * From ["&Isforum&"User] Where UserName='"&NewUserName&"'").Eof Then
		SuccessMsg " �û����ظ�������������һ���û�����"
	End If
	ExtCredits= Split(team.Club_Class(21),"|")
	team.Execute( "insert into ["&Isforum&"User] (UserName,Userpass,UserGroupID,Usermail,UserSex,Newmessage,Posttopic,Postrevert,Deltopic,Goodtopic,Extcredits0,Extcredits1,Extcredits2,Regtime,Landtime,Postblog,UserInfo,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,LevelName,Members) values('"&NewUserName&"','"&MD5(Newpassword,16)&"',26,'"&Newemail&"',0,0,0,0,0,0,"&Cid(Split(ExtCredits(0),",")(2))&","&Cid(Split(ExtCredits(1),",")(2))&","&Cid(Split(ExtCredits(2),",")(2))&","&SqlNowString&","&SqlNowString&",0,'|||||',"&Cid(Split(ExtCredits(3),",")(2))&","&Cid(Split(ExtCredits(4),",")(2))&","&Cid(Split(ExtCredits(5),",")(2))&","&Cid(Split(ExtCredits(6),",")(2))&","&Cid(Split(ExtCredits(7),",")(2))&",'��Сһ�꼶||||||0||0','ע���Ա')" )

	If Request.Form("emailnotify") = "yes" Then
		Dim Mailtopic,Body
		Mailtopic="��ȷ���û���ע��֪ͨ��"
		Body="�װ���"&NewUserName&", ����!"&vbCrlf&""&vbCrlf&" [���ʼ���ϵͳ�Զ�����] ��ϲ���õ� "&team.Club_Class(1)&" ��ע���û�Ȩ�ޡ�"&vbCrlf&""&vbCrlf&"��* �����ʺ���:"&NewUserName&"��������:"&Newpassword&" "&vbCrlf&""&vbCrlf&"��* "&vbCrlf&""&vbCrlf&" * ���, �м���ע�����������μ�"&vbCrlf&"1�������ء��������Ϣ�������������ȫ��������취�����һ�й涨��"&vbCrlf&"2��ʹ�����ɶ������Ļ��⣬�����벻Ҫ�漰���Ρ��ڽ̵����л��⡣"&vbCrlf&"3���е�һ����������Ϊ��ֱ�ӻ��ӵ��µ����»����·������Ρ�"&vbCrlf&""&vbCrlf&""&vbCrlf&"��̳������ "&team.Club_Class(1)&"("&team.Club_Class(2)&") �ṩ ��"&vbCrlf&"[����̳Դ������:TEAM5.CN�ṩ]"&vbCrlf&""&vbCrlf&""&vbCrlf&""
		Call Emailto(Newemail,Mailtopic,Body)
	End If
	SuccessMsg " ���û� "&NewUserName&" �Ѿ������ɣ�Ĭ������Ϊ "&Newpassword&"�� "
End Sub

Sub AddUser %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=adduserok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="2">������û�</td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">�û���:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="text" name="newusername"></td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">����:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="text" name="newpassword"></td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">Email:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="text" name="newemail"></td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">����֪ͨ��������ַ:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="checkbox" name="emailnotify" value="yes"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="addsubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub Main()
	Dim Gs,Value,i,u
	Set Gs = team.execute("Select ID,GroupName From "&IsForum&"UserGroup Where GroupRank>0 Order By ID ASC")
	If Gs.Eof or Gs.Bof Then
		SuccessMsg " �û�Ȩ�ޱ�������,���ֶ������±�! "
	Else
		Value = Gs.GetRows(-1)
	End If
	Gs.Close:Set Gs=Nothing
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
  <tr class="a1">
    <td>TEAM's ��ʾ</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li> ϵͳ�����û� <B><%=Application(CacheName&"_UserNum")%></B> λ��</li>
        <li> �����û�����ʹ��ֱ�������û���Ҳ����ʹ�����������ģ�����ҡ�</li>
        <li> ϵͳ�ṩ3���ύ��ť���ܣ����û��Ŀ�ݹ������ͨ���˴�ֱ��ִ�С�</li>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=findmembers">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr class="a1">
      <td colspan="2" align="center">�û����� </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">�û���������</td>
      <td bgcolor="#FFFFFF"><input type="text" name="username" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">24Сʱ�ڵ�¼���û���</td>
      <td bgcolor="#FFFFFF"><input type="checkbox" name="userlogin" value="1">
        ѡ�� </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">24Сʱ��ע����û���</td>
      <td bgcolor="#FFFFFF"><input type="checkbox" name="newuserreg" value="1">
        ѡ�� </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">���û�����ѯ��</td>
      <td bgcolor="#FFFFFF"><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr>
            <%If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				Echo "<td><input type=""checkbox"" name=""lookperm"" value="""&Replace(Value(1,i),"'","''")&""">"&Value(1,i)&"</td> "
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table></td>
    </tr>
  </table>
  <BR>
  <center>
    <input type="submit" name="searchsubmit" value="�����û�">
    &nbsp;
    <input type="submit" name="newslettersubmit" value="��̳֪ͨ">
    &nbsp;
    <input type="submit" name="creditsubmit" value="���ֽ���">
    &nbsp;
  </center>
  <BR>
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr class="a1">
      <td align="center" colspan="2">�߼���ѯ</td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">��߲�ѯ����</td>
      <td bgcolor="#FFFFFF"><input type="text" name="findpages" size="40" value="20">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">Email������</td>
      <td bgcolor="#FFFFFF"><input type="text" name="usermail" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">ע��IP������</td>
      <td bgcolor="#FFFFFF"><input type="text" name="regip" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">�û�ǩ��������</td>
      <td bgcolor="#FFFFFF"><input type="text" name="usersign" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">����+����������
        <input type="radio" name="maxpost1" value="1" checked>
        &nbsp;����&nbsp;
        <input type="radio" name="maxpost1" value="0">
        &nbsp;����</td>
      <td bgcolor="#FFFFFF"><input type="text" name="maxpost" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">����ʱ����
        <input type="radio" name="maxlogin1" value="1" checked>
        &nbsp;����&nbsp;
        <input type="radio" name="maxlogin1" value="0">
        &nbsp;����</td>
      <td bgcolor="#FFFFFF"><input type="text" name="maxlogin" size="40" value="">
        (����)</td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">ע�����ڣ�
        <input type="radio" name="regtime1" value="0" checked>
        &nbsp;����&nbsp;
        <input type="radio" name="regtime1" value="1">
        &nbsp;����&nbsp;
        <input type="radio" name="regtime1" value="2">
        &nbsp;���� (yyyy-mm-dd):</td>
      <td bgcolor="#FFFFFF"><input type="text" name="regtime" size="40" value="">
      </td>
    </tr>
    <%
	Dim ExtCredits,ExtSort
	ExtCredits= Split(team.Club_Class(21),"|")
	For U=0 to 2
		ExtSort=Split(ExtCredits(U),",")
		Echo "<tr>"
		Echo " <td bgcolor=""#F8F8F8"">"&ExtSort(0)&"��"
		Echo "	<input type=radio name=""Nums"&U&""" value=""1"" checked>&nbsp;����&nbsp; "
        Echo "	<input type=radio name=""Nums"&U&""" value=""0"">&nbsp;����</td>"
		Echo "	<td bgcolor=""#FFFFFF""><input type=""text"" name=""MyCred"&U&""" size=""40"" value=""""></td>"
		Echo "</tr>"
	Next
	%>
  </table>
  <BR>
  <center>
    <input type="submit" name="searchsubmit" value="�����û�">
    &nbsp;
    <input type="submit" name="newslettersubmit" value="��̳֪ͨ">
    &nbsp;
    <input type="submit" name="creditsubmit" value="���ֽ���">
    &nbsp;
  </center>
  <BR>
</form>
<%
end sub

footer()
%>
