<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim ii,ID
Dim Admin_Class
Call Master_Us()
Header()
ii=0:ID=Request("id")
Admin_Class=",3,"
Call Master_Se()
team.SaveLog ("��ݹ��� [��������ݹ����������] ")
Select Case Request("Action")
	Case "deltopicok"
		Call deltopicok
	Case "delforumok"
		Call delforumok
	Case "delusertopicok"
		Call delusertopicok
	Case "delliketopicok"
		Call delliketopicok
	Case "delretopicok"
		Call delretopicok
	Case "deluserretopicok"
		Call deluserretopicok
	Case "UniForum"
		Call UniForum	'�ϲ����
	Case "Forumsmerge"
		Call Forumsmerge
	Case "uniteok"
		Call uniteok
	Case "readkey"
		CAll readkey
	Case "readkeyok"
		Call readkeyok
	Case Else
		Call Main()
End Select

Sub readkeyok
	Dim ho
	For each ho in request.form("checktid")
		team.execute("Update ["&Isforum&"Forum] Set Auditing=0 Where ID="&ho)
	Next
	team.SaveLog ("����������")
	SuccessMsg " ���������ɣ���ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp?Action=readkey>������� </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp?Action=readkey>�� "	
End sub

Sub readkey%>
	<br>
	<br>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?action=readkeyok">
	<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr class="a1">
	 <td>������ʾ</td>
	</tr>
	<tr class="a4">
	 <td><br>
      <ul>
        <li>�������ڴ�����û���������ӻ������
		<li>ֻ���ڰ��������<B>�������</B>���ܣ����Ӳ���Ҫ��ˡ�ͨ������˵����Ӳſ���������ʾ�ڰ������ ��
      </ul></td>
	 </tr>
	</table>
	<BR>
	<table cellspacing="1" cellpadding="3" border="0" width="95%" align="center" class="a2">
	<tr class="tab1"> 
		<td align="center" width="10%"><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio">���</td> <td align="center">����</td>
	</tr>
	<%
	Dim Rs
	Set Rs=team.execute("Select * From ["&IsForum&"Forum] Where Auditing=1 and Deltopic=0")
	Do While Not Rs.Eof
			Echo "<tr class=""a4"">"
			Echo "	<td align=""center""><input type=""checkbox"" name=""checktid"" class=""radio"" value="""&RS(0)&"""></td>"
			Echo "	<td> <a href=""../SeeDeltop.asp?tid="&Rs("ID")&""" target=""_blank""> "& Rs("topic") &"</a> </td>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo "</table><br><center><input type=""submit"" name=""onlinesubmit"" value=""�� ��""></center></form>"
End Sub

Sub uniteok
	Dim Source,Target,UserTips
	Source = Request.Form("source")
	Target = Request.Form("target")
	If Source="" or Target="" Then Error2 "Դ��̳��Ŀ����̳����Ϊ��!"
	If CID(Source) = CID(Target) Then Error2 "Դ��̳��Ŀ����̳������ͬ!"
	if Request.Form("postname") <> "" then 
		UserTips =" and UserName='"&HtmlEncode(Trim(Request.Form("postname")))&"'"
	End If
	If IsSqlDataBase=1 Then
		team.Execute( "Update ["&Isforum&"forum] Set Forumid="&Target&" Where Forumid="&Source&" and Datediff(d,Posttime, " & SqlNowString & ") > " & Cid(Request.Form("posttime"))&" "&UserTips&" ")
	Else
		team.Execute( "Update ["&Isforum&"forum] Set Forumid="&Target&" Where Forumid="&Source&" and Datediff('d',Posttime, " & SqlNowString & ") > " & Cid(Request.Form("posttime"))&" "&UserTips&" ")
	End If
	SuccessMsg "��̳�ƶ��ɹ�����ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> "
End Sub

Sub Forumsmerge
	Dim Source,Target
	Source = Request.Form("source")
	Target = Request.Form("target")
	If Source="" or Target="" Then Error2 "Դ��̳��Ŀ����̳����Ϊ��!"
	If CID(Source) = CID(Target) Then Error2 "Դ��̳��Ŀ����̳������ͬ!"
	team.execute("Update "&IsForum&"forum Set Forumid="&Target&" Where Forumid="&Source)
	team.execute("Delete from "&IsForum&"Bbsconfig Where id="&Source)
	SuccessMsg "��̳�ϲ��ɹ�,Դ��̳�Ѿ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> "
End Sub

Sub deluserretopicok
	Call Master_Se()
	Dim BbsID,FindBoard,BoardName,Rs
	BbsID = Request.Form("bbsid")
	If Request.Form("postname") = "" Then 
		Error2 "��û�������û�����"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then
			Set Rs=Team.execute("Select ID From "&Isforum&"Forum Where forumid="& BbsID)
			While Not RS.EOF
				team.Execute( "Delete From ["&Isforum&""& Request.Form("Reforumname") &"] Where UserName = '"&HtmlEncode(Trim(Request.Form("postname")))&"' and Topicid="&RS(0) )
				Rs.MoveNext
			Wend
		Else
			team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where UserName = '"&HtmlEncode(Trim(Request.Form("postname")))&"' ")
		End If
		
		SuccessMsg " �Ѿ��������� [ "&Request.Form("Reforumname")&" ] �����û�"&Request.Form("postname")&" ����Ļ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delretopicok
	Call Master_Se()
	Dim BbsID,FindBoard,Rs
	BbsID = Request.Form("bbsid")
	If Request.Form("posttime") = "" or ( Not isNumeric(Request.Form("posttime")) ) Then 
		Error2 "���ڱ���Ϊ���֡�"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			Set Rs=Team.execute("Select ID From "&Isforum&"Forum Where forumid="& BbsID)
			While Not RS.EOF
				If IsSqlDataBase=1 Then
					team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff(d,Posttime, " & SqlNowString & ") > " & Request.Form("posttime")&" And Topicid="&RS(0) )
				Else
					team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff('d',Posttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" And Topicid="& RS(0) )
				End If
				Rs.MoveNext
			Wend
		Else
			If IsSqlDataBase=1 Then
				team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff(d,Posttime, " & SqlNowString & ") > " & Request.Form("posttime")&" ")
			Else
				team.Execute( "Delete From ["&Isforum&""&Request.Form("Reforumname")&"] Where Datediff('d',Posttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" ")
			End If
		End If
		SuccessMsg " �Ѿ��������� [ "&Request.Form("Reforumname")&" ]���� "&request("posttime")&" ����ǰ�Ļ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delliketopicok
	Call Master_Se()
	Dim BbsID,FindBoard,BoardName
	BbsID = Request.Form("bbsid")
	If Request.Form("topic") = "" Then 
		Error2 "��û�������ַ���"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
		End If
		team.Execute( "Delete From ["&Isforum&"forum] Where Topic Like  '%"&HtmlEncode(Trim(Request.Form("topic")))&"%'  "& FindBoard )
		SuccessMsg " �Ѿ�������������� "&Request.Form("topic")&"  ������ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delusertopicok
	Call Master_Se()
	Dim BbsID,FindBoard,BoardName
	BbsID = Request.Form("bbsid")
	If Request.Form("postname") = "" Then 
		Error2 "��û�������û�����"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
			BoardName = " �ڰ�� <a href=../BoardList.asp?ID="& BbsID &" "
		End If
		team.Execute( "Delete From ["&Isforum&"forum] Where UserName = '"&HtmlEncode(Trim(Request.Form("postname")))&"'  "& FindBoard )
		SuccessMsg " �Ѿ��� "&Request.Form("postname")&"  "&BoardName&" ���������ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub delforumok
	Call Master_Se()
	Dim BbsID,FindBoard
	BbsID = Request.Form("bbsid")
	If Request.Form("posttime") = "" or ( Not isNumeric(Request.Form("posttime")) ) Then 
		Error2 "���ڱ���Ϊ���֡�"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
		End If
		If IsSqlDataBase=1 Then
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff(d,Lasttime, " & SqlNowString & ") > " & Request.Form("posttime")&" "& FindBoard )
		Else
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff('d',Lasttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" "& FindBoard )
		End If
		SuccessMsg " �Ѿ���"&request("posttime")&"��û�и��¹�������ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub deltopicok
	Call Master_Se()
	Dim BbsID,FindBoard
	BbsID = Request.Form("bbsid")
	If Request.Form("posttime") = "" or ( Not isNumeric(Request.Form("posttime")) ) Then 
		Error2 "���ڱ���Ϊ���֡�"
	Else
		If ( Not BbsID = "" ) or isNumeric(BbsID) Then 
			FindBoard = " and forumid= "& BbsID 
		End If
		If IsSqlDataBase=1 Then
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff(d,Posttime, " & SqlNowString & ") > " & Request.Form("posttime")&" "& FindBoard  )
		Else
			team.Execute( "Delete From ["&Isforum&"forum] Where Datediff('d',Posttime, " & SqlNowString & " ) > "& Request.Form("posttime")&" "& FindBoard  )
		End If
		SuccessMsg " �Ѿ���"&request("posttime")&"����ǰ������ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_Manage.asp>��ݹ���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Manage.asp> " 
	End If
End Sub

Sub Main()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2" >
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><BR>
      <ul>
        <li>������̳�رյ�����½������в��������ڲ�����ɺ� <a href=Admin_Update.asp><B>������̳ͳ��</B></a> ��Ϣ��</li>
      </ul>
      <ul>
        <li>�˲����ǲ�����ת�ģ������Ƽ��� <a href=Admin_MDBS.asp><B>���ݺ����ݿ�</B></a> �Ժ��������ȷ������ÿһ�����衣</li>
      </ul>
      <ul>
        <li>�˲�����ʱ�佫��������ݿ��С���죬���ݿ�Խ��ʱ��Խ����</li>
      </ul>
      <ul>
        <li> ����̳������̳����<B><%=team.execute("Select count(id)from bbsconfig")(0)%></B> ��
          ��������<B><%=team.execute("Select count(id)from forum")(0)%></font></B> ��
          ��ǰ����������<B><%=team.execute("Select count(id)from ["&team.Club_Class(11)&"]")(0)%></B>���������̳��ҳ��ͳ�ƺ͵�ǰ��ͳ�������� <a href=Admin_Update.asp><B>������̳ͳ��</B> ��</li>
      </ul></td>
  </tr>
</table>
<BR>
<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr>
    <td class="a1" colspan="3">����ɾ������</td>
  </tr>
  <form method="post" action="?Action=deltopicok">
    <tr class="a4">
      <td width="40%"> ɾ�� <INPUT size="3" name="posttime" value="180"> ����ǰ������</td>
      <td width="40%"><select name="bbsid">
          <option value="">������̳</option>
          <%ForumList_Sel(0)%>
        </select></td>
      <td width="20%"><input type="submit" value=" ȷ �� "></td>
  </tr>
</form>
<form method="post" action="?Action=delforumok">
    <tr class="a3">
      <td> ɾ��<INPUT size="3" name="posttime" value="180"> ��û�и��µ�����</td>
      <td><select name="bbsid">
          <option value="">������̳</option>
          <%ForumList_Sel(0)%>
        </select></td>
      <td><input type="submit" value=" ȷ �� "></td>
  </tr>
   </form>
  <form method="post" action="?Action=delusertopicok">
  <tr class="a4">
	<td>ɾ�� <input size="10" name="postname"> ������������� </td>
    <td> <select name="bbsid">
			<option value="">������̳</option>
			<%ForumList_Sel(0)%>
			</select>
	</td>  
    <td><input type="submit" value=" ȷ �� "></td>
  </tr>
  </form>
  <form method="post" action="?Action=delliketopicok">
  <tr class="a3">
	<td>ɾ������������� <input size="10" name="topic"> ����������</td>
    <td><select name="bbsid">
			<option value="">������̳</option>
			<%ForumList_Sel(0)%>
			</select>
	</td>
    <td><input type="submit" value=" ȷ �� "></td>
   </tr>
</form>
</table>
<BR/>

<table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
  <tr  class="a1">
    <td colspan="3"> ����ɾ������ </td>
  </tr>
  <form method="post" action="?Action=delretopicok">
    <tr class="a4">
      <td width="40%"> ɾ�� <INPUT size="3" name="posttime" value="180"> ����ǰ�Ļ���</td>
      <td width="40%"><select name="bbsid">
						<option value="">������̳</option>
						 <%ForumList_Sel(0)%>
						 </select> - <select name="Reforumname">
										<option value="ReForum">��ѡ�������</option>
<%	Dim Value,i,Rs1
	Set Rs1 = Team.Execute(" Select id,TableName From TableList ")
	If Not Rs1.Eof Then
		Value = Rs1.GetRows(-1)
	End If
	Rs1.Close:Set Rs1=Nothing
	If IsArray(Value) Then
		For i=0 To Ubound(Value,2)
			Echo "<option value="&Value(1,i)&">"&Value(1,i)&"</option>"
		Next
	End If
	%></select>
      </td>
      <td width="20%"><input type="submit" value=" ȷ �� "></td>
    </tr>
  </form>
  <form method="post" action="?Action=deluserretopicok">
  <tr Class="a3">
	<td>ɾ�� <input size="10" name="postname"> ��������л��� </td> 
    <td>	<select name="bbsid">
				<option value="">������̳</option>
				<%ForumList_Sel(0)%>
			</select> -
			<select name="Reforumname">
				<option value="ReForum">��ѡ�������</option>
				<%
			If IsArray(Value) Then
				For i=0 To Ubound(Value,2)
					Echo "<option value="&Value(1,i)&">"&Value(1,i)&"</option>"
				Next
			End If	%>
		</select>
    </td>  
    <td><input type="submit" value=" ȷ �� ">
</tr>
</form>
</table>
<BR>
  <center>
    <input type="submit" name="submit" value="�� ��" onclick="{if(confirm('��ȷ��Ҫɾ����̳ô?')){return true;}return false;}">
  </center>
</form>
<form method="post" action="?Action=uniteok">
  <table cellspacing="1" cellpadding="3" width="90%" border="0" class="a2" align="center">
    <tr class="a1">
      <td colspan="3">�ƶ���̳ - ��ָ����̳�����Ӱ�������ɸѡת��Ŀ����̳��ͬʱ����Դ��̳</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">Դ��̳:</td>
      <td bgcolor="#FFFFFF" width="60%"  align="left">��
        <select name="source">
          <option value="">�� ��ѡ��</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">Ŀ����̳:</td>
      <td bgcolor="#FFFFFF" width="60%"  align="left">��
        <select name="target">
          <option value="">�� ��ѡ��</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">�����趨:</td>
      <td bgcolor="#FFFFFF" width="60%" align="left">�����ƶ�
        <input size="2" name="posttime" value="0">
        ��ǰ������</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">�û��趨:</td>
      <td bgcolor="#FFFFFF" width="60%"  align="left">�����ƶ�
        <input size="8" name="postname">
        ���������</td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="submit" value="�� ��">
  </center>
</form>
<form method="post" action="?Action=Forumsmerge">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="3">�ϲ���̳ - Դ��̳������ȫ��ת��Ŀ����̳��ͬʱɾ��Դ��̳</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">Դ��̳:</td>
      <td bgcolor="#FFFFFF" width="60%" align="left">��
        <select name="source">
          <option value="">�� ��ѡ��</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">Ŀ����̳:</td>
      <td bgcolor="#FFFFFF" width="60%" align="left">��
        <select name="target">
          <option value="">�� ��ѡ��</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="submit" value="�� ��">
  </center>
</form>
<br>
<%
End Sub

Sub DelForums%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?Action=ServerDelForum&ID=<%=request("ID")%>">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan=5>TEAM's ��ʾ</td>
    </tr>
    <tr align="center">
      <td bgcolor="#FFFFFF"><br>
        <br>
        <br>
        ���������ɻָ�����ȷ��Ҫɾ������̳������������Ӻ͸�����?<br>
        ע��: ɾ����̳����������û��������ͻ���<br>
        <br>
        <br>
        <br>
        <input type="submit" name="forumsubmit" value=" ȷ �� ">
        &nbsp;
        <input type="button" value=" ȡ �� " onClick="history.go(-1);"></td>
    </tr>
  </table>
</form>
<br>
<%
End Sub

Sub ForumList_Sel(V)
	Dim SQL,ii,RS,W
	Set Rs=Team.Execute("Select ID,BBSname,Followid From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		W="���� "
		If V = 0 Then W="�� "
		Response.Write "<option value="&RS(0)&""
		Response.Write ">"&String(ii,"��")&""&W&""&RS(1)&"</option>"
		ii=ii+1
		ForumList_Sel RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub
%>
