<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<!-- #include file="../inc/MD5.asp" -->
<%
Public boards
Dim Admin_Class,Uid
Call Master_Us()

Uid = Cid(Request("uid"))
If Cid(Session("UserMember")) <> 1 Then 
	SuccessMsg "�Բ���ֻ�й���Ա���ɲ鿴�˰������ �� "
End if
Header()
Select Case Request("action")
	Case "masterupdate"
		Call masterupdate
	Case "manages"
		Call manages
	Case "edmaster"
		Call edmaster
	Case "edmasterok"
		Call edmasterok
	Case "upkey"
		Call upkey
	Case "upkeyok"
		Call upkeyok
	Case "killmaster"
		Call killmaster
	Case Else
		Call Main
End Select

Sub killmaster
	If Uid="" or Not IsNumeric(Uid) Then
		SuccessMsg "��������"
	Else
		team.Execute("Delete From ["&isforum&"Admin] Where ID="&UID)
		SuccessMsg "�˺�̨��½�û��Ѿ���ɾ����"
	End If
End Sub

Sub upkeyok
	If Uid="" or Not IsNumeric(Uid) Then
		SuccessMsg "��������"
	Else
		team.Execute("Update ["&isforum&"Admin] Set adminname='"&request.Form("adminname")&"',forumname='"&request.Form("forumname")&"',adminpass='"&Md5(request.Form("adminpass"),16)&"' Where ID="&UID)
		SuccessMsg "��̨���������޸���ɡ�"
	End If
End Sub

Sub upkey
	Dim Rs
	Set Rs=TEAM.Execute("Select id,adminname,forumname From ["&isforum&"admin] Where ID="& UID)
	If Not Rs.Eof Then
	%>
<BR>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form name="form" method="post" action="?action=upkeyok">
  <input name="uid" type="hidden" value="<%=RS(0)%>">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr class="a1">
      <td colspan="3">�޸ĺ�̨����Ա����</td>
    </tr>
    <tr class="a3">
      <td align="center" width="40%">��̨��½���ƣ�</td>
      <td width="30%"><input name="adminname" size="30" value="<%=RS(1)%>"></td>
      <td width="30%">(����ע������ͬ) </td>
    </tr>
    <tr class="a4">
      <td align="center">��̨��½���룺</td>
      <td><input name="adminpass" size="30"></td>
      <td>(����ע�����벻ͬ)</td>
    </tr>
    <tr class="a3">
      <td align="center">ǰ̨�û����ƣ�</td>
      <td><input name="forumname" size="30" value="<%=RS(2)%>"></td>
      <td>һ��ǰ̨�û����ɰ󶨶����̨���Ƶ�½!</td>
    </tr>
  </table>
  <br>
  <center>
  <input type=submit value="ȷ��">
  <center>
</form>
<%
	End if
End Sub

Sub edmasterok	
	Dim Admin_s
	Admin_s = Replace(Request.Form("Admin_Pass")," ","")
	If Uid="" or Not IsNumeric(Uid) Then
		SuccessMsg "��������"
	Else
		team.execute("Update ["&isforum&"admin] set adminclass='"&Admin_s&"' Where ID="&UID)
		SuccessMsg "��̨Ȩ��������ɡ�"
	End If
End Sub

Sub edmaster
	dim menu(4,3),trs,k
	menu(0,0)=" ��̳��̨����Ȩ�޷��� "
	menu(0,1)="<a href=Admincp.asp> ����ѡ�� </a> [������������Ŀ ]@@1"
	menu(1,1)="<a href=Admin_Forum.asp> ��̳����  </a> [�������༭���  ]@@2"
	menu(1,2)="<a href=Admin_Manage.asp> ��ݹ���  </a> [��������ݹ����������]@@3"
	menu(1,3)="<a href=Admin_Update.asp> ������̳ͳ��  </a> [������������̳ͳ�� ]@@4"
	menu(2,1)="<a href=Admin_Group.asp> �����뼶��   </a> [������������ ���û��� ]@@5"
	menu(2,2)="<a href=Admin_User.asp> �û�����  </a> [�������༭�û�������û� ���ϲ��û� ������û� �����ʹ��� ]@@6"
	menu(3,1)="<a href=admin_skins.asp> ������ </a> [�������༭ģ�� ��ģ�嵼�룬ģ�嵼�� ]@@7"
	menu(3,2)="<a href=Admin_Change.asp>��������  </a> [��������̳���� ���������� ��ѫ�±༭ �������� �������б��� ]@@8"
	menu(3,3)="<a href=Admin_plus.asp>������� </a> [�������˵����� ������������Ա ]@@9"
	menu(4,1)="<a href=Admin_dbmake.asp> ��̳ά�� </a> [���������ݿ���� �����ݿ����� ������������ ���������� �����Ź��� ��������¼ ]@@10"
	menu(4,2)="<a href=Admin_Path.asp> ͳ����Ϣ  </a> [������������������ �����֧�������ͳ��ռ�ÿռ� ]@@11"

	dim j,tmpmenu,menuname,menurl,rs
	Set Rs=Team.Execute("Select forumname,adminclass From "&IsForum&"Admin Where ID="&UID )	
	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>TEAMϵͳ���Խ����̨����ĵȼ���Ĭ���� <B>����Ա</B> �� �������� ������Աӵ������Ȩ�ޣ� ����������ӵ�г��˶Ժ�̨�����û������ɾ���ȹ�����֮�������Ȩ�ޣ���������Ȩ�ޱ�����<FONT COLOR="red">Ȩ�޸���</FONT>������²ſ���ִ�С����Խ���Գ�����������һ���ĺ�̨Ȩ�ޣ��Խ��͹���Ա�Ĺ���ǿ�� ��
      </ul></td>
  </tr>
</table>
<br>
<form action="?action=edmasterok" method="post">
  <table cellpadding="3" cellspacing="1" border="0" width="95%" class="a2" align="center">
    <tr class="a1">
      <td><b>����ԱȨ�޹���</b>( ��ѡ����Ӧ��Ȩ�޷��������Ա <%=rs("forumname")%> )</td>
    </tr>
    <tr>
      <td class="a3"><table cellpadding="3" cellspacing="1" border="0" width="95%" class="a2" align="center">
          <tr>
            <td class="a1">ȫ��Ȩ��</td>
          </tr>
          <%
	for i=0 to ubound(menu,1)
		Echo" <tr><td class=""a3"">"& menu(i,0) &"</td></tr>"
		on error resume next
		for j=1 to ubound(menu,2)
			if isempty(menu(i,j)) then exit for
			tmpmenu=split(menu(i,j),"@@")
			menuname=tmpmenu(0)
			menurl=tmpmenu(1)
			response.write	"<tr><td class=""a4""> <input type=""checkbox"" name=""Admin_Pass"" value="&menurl&" "
			if instr(","&rs(1)&",",","&menurl&",")>0 then response.write "checked" 
			response.write ">"
			Echo "" & menurl &" . "&menuname&" </td></tr> "
			next
	next
	%>
          <tr>
            <td class="a4"><input type="hidden" name="uid" value="<%=UID%>">
              <input type="checkbox" name="chkall" onClick="checkall(this.form,'Admin_Pass')">
              ѡ������Ȩ��</td>
          </tr>
        </table>
        <BR>
        <center>
          <input type="submit" name="Submit" value="����">
        </center>
        <BR>
        <BR>
      </td>
    </tr>
  </table>
  <BR>
  <BR>
</form>
<%
	Rs.Close:Set RS=Nothing
End Sub

Sub manages %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>������ڴ˽��к�̨����Ա����Ӻͺ�̨����ĸ��ģ��鿴������Ա�ĵ�½������޸Ĺ���Ա��Ȩ�޵ȵȡ�
      </ul>
      <ul>
        <li>�˰�Ĺ���ֻ��ӵ�й���Ա�ȼ����û����ɵ�½����
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=masterupdate">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
<tr align="center" class="a1">
  <td colspan="7">�鿴��̨�����½״�� </td>
</tr>
<tr class="a3" align="center">
  <td>��½�û�</td>
  <td>ǰ̨�û���</td>
  <td>��½IP</td>
  <td>����½ʱ��</td>
  <td>Ȩ��</td>
  <td>����</td>
  <td>����</td>
</tr>
<%
	Dim RS
	Set Rs=Team.Execute("Select ID,adminname,forumname,Loginip,Logintime From ["&Isforum&"admin] Order By Logintime Desc")
	Do While Not RS.EOF
		Echo "<tr class=a4 align=""center"">"
		Echo "	<td>"&Rs(1)&"</td>"
		Echo "	<td>"&Rs(2)&"</td>"
		Echo "	<td>"&Rs(3)&"</td>"
		Echo "	<td>"&Rs(4)&"</td>"
		Echo "	<td> <a href=""?action=edmaster&uid="&RS(0)&""">�༭Ȩ��</a> </td>"
		Echo "	<td> <a href=""?action=upkey&uid="&RS(0)&""">�޸�����</a> "
		Echo "	<td> <a href=""?action=killmaster&uid="&RS(0)&""">ɾ��?</a></td>	"
		Echo "</tr>"
		RS.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo " </table><br><center><input type=""submit"" name=""detailsubmit"" value=""�� ��""></form><br>"
End Sub



Sub masterupdate
	Dim Rs,adminname,adminpass,forumname,Us
	Adminname = Replace(Request.Form("Adminname"),"'","''")
	Adminpass = Replace(Request.Form("adminpass"),"'","''")
	Forumname = Replace(Request.Form("forumname"),"'","''")
	Set Us = team.Execute("Select UserGroupID from ["&Isforum&"User] Where UserName='"&Forumname&"'")
	If Us.Eof And Us.bof Then
		SuccessMsg "��ǰ̨�û���������,���������á�"
	Else
		Set Rs= team.execute("Select adminname,forumname from ["&Isforum&"admin] ")
		Do While Not Rs.Eof
			If LCase(Rs(0))=Lcase(Adminname) Then 
				SuccessMsg " �˺�̨�û����Ѿ�����,����������! �������Ҫ�޸�����,��ʹ���޸�����Ĺ��� ��"
			End If
			If LCase(Rs(1))=Lcase(Forumname) Then 
				SuccessMsg " ���û����Ѿ������˺�̨�����û����� ��"
			End if
			Rs.Movenext
		Loop
		Rs.Close:Set Rs=Nothing
		If Len(Adminpass)< 6  Then 
			SuccessMsg " �������벻������6λ�� ��"
		Else
			team.Execute( "insert into "&Isforum&"admin (adminname,adminpass,forumname) values ('"&Adminname&"','"&MD5(Adminpass,16)&"','"&Forumname&"')" )
			If Int(Us(0)) >2 Then
				team.Execute("Update ["&Isforum&"User] Set UserGroupID=2,LevelName='��������||||||18||0',Members='��������' Where UserName='"&Forumname&"'")
			End if
			SuccessMsg " ��̨����Ա��ӳɹ���<BR> Ĭ����ӵ��û����ڳ��������飬�������Ҫ�����û��������Ա��,�� ת��<A HREF=""Admin_User.asp""><B>�༭�û� </B></A>ѡ�������ȴ�ϵͳ�Զ����ص� <a href=Admin_maste.asp?action=manages>����Ȩ������  </a> ҳ�棬��ѡ�� <B>�༭Ȩ��</B> ѡ��Ժ�̨�û����к�̨����Ȩ�ޱ༭ ��<meta http-equiv=refresh content=3;url=Admin_maste.asp?action=manages>�� "
		End If
	End If
	Us.Close:Set Us = nothing
End Sub

Sub Main
	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>���ô�ѡ����Ը���ǰ̨�û������̨�����Ȩ�ޣ������ʹ�ô�Ȩ�ޡ�
      </ul>
      <ul>
        <li>ǰ̨�û������Ϊ��̨����Ա�����Զ����������ĳ��������飬ӵ�г��������Ȩ�ޡ�
      </ul>
      <ul>
        <li>�������ú�̨�����û���������Ĭ���û��� admin ��½����̨��Ȼ����һ�²���ִ�У�
          <ul>
            <li> 1. <A HREF="Admin_User.asp?action=adduser"><B>����µ��û�</B></A> </li>
          </ul>
          <ul>
            <li> 2. Ȼ��ת��<A HREF="Admin_User.asp"><B>�༭�û� </B></A>ѡ��������û���������ɺ��� <B>�û�����</B>���� <B>�û����������</B> ����Ϊ <B>����Ա</B> ��</li>
          </ul>
          <ul>
            <li> 3. ת�� <A HREF="Admin_maste.asp"><B>����Ա��� </B></A>ѡ��,����µĺ�̨�����û����ƣ���<B>ǰ̨�û�����</B>����Ϊ�ղ���ӵ��û����ơ�</li>
          </ul>
          <ul>
            <li> 4. ���ת�� <A HREF="Admin_maste.asp?action=manages"><B>����Ȩ������ </B></A>ѡ��,������ӵĺ�̨�û�������ϸ�ĺ�̨����Ȩ�ޡ�</li>
          </ul>
        </li>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=masterupdate">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
      <td colspan="2">��Ӻ�̨����Ա</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="60%"><B>��̨��½���ƣ�</B><BR>
        ����Ա��½��̨ʱ��ʹ�õĵ�½�����˵�½������ǰ̨��ע�������Բ�ͬ�����ǲ������Ѿ����ڵĺ�̨�û������ظ���ÿ�������߶�ӵ�ж����ĺ�̨��½���ƣ����û������ڵ�½��̨����ʱ��Ч�� </td>
      <td bgcolor="#F8F8F8"><input name="adminname" size="30" value="<%=TK_UserName%>">
      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="60%"><B>��̨��½���룺</B><BR>
        ����Ա��½����̨��Ҫ��������룬ÿ����̨�����߶�ӵ��һ�������Ĺ������룬����Ա�����Ҫ��½����̨��������������ȷ����ſ��Ե��롣 </td>
      <td bgcolor="#F8F8F8"><input name="adminpass" size="30" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="60%"><B>ǰ̨�û����ƣ�</B><BR>
        ÿ����̨�������ƶ���Ӧһ��ǰ̨�Ĺ���Ա���ƣ� �� Ĭ������ <FONT COLOR="RED">admin</FONT> ���Ӧ���û�Ϊ����Ա <FONT COLOR="RED">admin</FONT> ���˴���������Ʊ���Ϊǰ̨���û���������û������ڣ������� <A HREF="Admin_User.asp?action=adduser"><B>����û�</B></A> ѡ������µ��û���Ȼ�������ð����ơ� </td>
      <td bgcolor="#F8F8F8"><input name="forumname" size="30" value="">
      </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="�� ��">
</form>
<br>
<%
End Sub

footer()
%>
