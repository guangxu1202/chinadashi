<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Public boards
Dim Admin_Class,Uid
Call Master_Us()
Uid = Cid(Request("uid"))
Header()
Admin_Class=",8,"
Call Master_Se()
Select Case Request("action")
	Case "medals"
		Call medals
	Case "medalsok"
		Call medalsok
	Case "announcements"
		Call announcements
	Case "newsannouncements"
		Call newsannouncements
	Case "announcementsok"
		Call announcementsok
	Case "forumlinks"
		Call forumlinks
	Case "forumlinksok"
		Call forumlinksok
	Case "adv"
		Call adv
	Case "advadd"
		Call advadd
	Case "advaddok"
		Call advaddok
	Case "advok"
		Call advok
	Case "advedit"
		Call advedit
	Case "onlinelist"
		Call onlinelist
	Case "onlinelistok"
		Call onlinelistok
End Select

Sub onlinelistok
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"OnlineTypes] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"OnlineTypes] set Sorts="&Cid(Request.Form("sorts"&i+1))&",OnlineName='"&Replace(Request.Form("titles"&i+1),"'","")&"',Onlineimg='"&Replace(Request.Form("urls"&i+1),"'","")&"' Where ID="&newid(i))
		Next
		if Request.Form("newMembers")<>"" and Request.Form("newurl")<>"" Then
			Dim mTitle
			mTitle = ""
			mTitle = Replace(Request.Form("newMembers"),"'","")
			If InStr(Trim(mTitle),"�ο�/δ��½")>0 Then mTitle ="�ο�"
			team.execute("insert into ["&Isforum&"OnlineTypes] (Sorts,OnlineName,Onlineimg) values ("&Cid(Request.Form("newsorts"))&",'"&mTitle&"','"&Replace(Request.Form("newurl"),"'","")&"' ) ")
		End if
	End If
	Application.Contents.RemoveAll()
	team.SaveLog ("����ͼ���������")
	SuccessMsg " ����ͼ��������ɣ���ȴ�ϵͳ�Զ����ص� <a href=Admin_Change.asp?action=onlinelist>�����б��� </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=onlinelist>�� "
End Sub

Sub onlinelist%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>�����������Զ�����ҳ�������б�ҳ��ʾ�����߻�Ա���鼰ͼ����ֻ�������б��ܴ�ʱ��Ч��
      </ul>
      <ul>
        <li>����δ�����ʾ���û����Ա������ʾ�������б���
      </ul>
      <ul>
        <li>�û���ͼ��������дͼƬ�ļ�����������ӦͼƬ�ļ��ϴ��� Skins/�������Ӧ���Ŀ¼�С�
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=onlinelistok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
	  <td><input type="checkbox" name="chkall" onClick="checkall(this.form)"> ɾ</td>
      <td>��ʾ˳��</td>
      <td>��ͷ��</td>
      <td>�û���ͼ��</td>
    </tr>
	<%
	Dim Rs,Imgs,i,Styleurl
	Set Rs = team.execute("Select Styleurl From ["&Isforum&"Style] Where ID= "& INT(team.Forum_setting(18)))
	If Not Rs.Eof Then
		Styleurl = Rs(0)
	Else
		Styleurl = "skins/teams"
	End If
	Rs.close:Set Rs=Nothing
	Set Rs = team.execute("Select ID,Sorts,OnlineName,Onlineimg From ["&isforum&"OnlineTypes] Order By Sorts Asc")
	If Rs.Eof Then
		SuccessMsg " δ�ҵ����ݱ���ȷ�����ݿ��Ѿ�������"
	End if
	i=0
	Do While Not Rs.Eof
		i=i+1
			If Rs(3)<>"" Then 
				Imgs = "<img src=../"&Styleurl&"/"&RS(3)&" align=""absmiddle"">"
			Else
				Imgs = ""
			End if
			Echo " <tr align=""center""> <td bgcolor=""#FFFFFF""><input Name=""newid"" type=""hidden"" value="&RS(0)&"> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&"></td>	"
			Echo " <td bgcolor=""#F8F8F8""><input type=""text"" size=""3"" name=""sorts"&i&""" value="&Rs(1)&"></td>"
			Echo " <td bgcolor=""#FFFFFF""><input type=""text"" size=""15"" name=""titles"&i&""" value="&Rs(2)&"></td>"
			Echo " <td bgcolor=""#F8F8F8"" align=""left""><input type=""text"" size=""40"" name=""urls"&i&""" value="&Rs(3)&"> "
			Echo " "&Imgs&"</td></tr>"
		Rs.moveNext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
    <tr align="center" class="a1">
	  <td> ���� </td>
      <td>��ʾ˳��</td>
      <td>��ͷ��</td>
      <td>�û���ͼ��</td>
    </tr>
    <tr align="center" class="a3">
	  <td> &nbsp; </td>
      <td> <input type="text" size="3" name="newsorts" value="0"> </td>
      <td><select name="newMembers" style="width:100%">
		<option value=""> ��ѡ���û��� </option>
		<%
		Dim Gs
		Set Gs = team.execute("Select ID,GroupName,Members From "&IsForum&"UserGroup Where MemberRank<=1 Order By ID ASC")
		Do While Not Gs.Eof
			If Gs(2)="�α�" Then
				Echo " <option value="""&Gs(1)&"""> "&Gs(1)&" </option> "
			Else
				Echo " <option value="""&Gs(2)&"""> "&Gs(2)&" </option> "
			End if
			Gs.MoveNext
		Loop
		Gs.Close:Set Gs=Nothing
		%>
		</select></td>
      <td align="left"> <input type="text" size="40" name="newurl" value=""> </td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="onlinesubmit" value="�� ��">
  </center>
</form>
</td>
</tr>
<br>
<br>
<%
End Sub

Sub advedit
	Dim Rs
	If Uid = "" Or Not IsNumeric(Uid) Then
		SuccessMsg " �������� "
	Else
		Set Rs = team.execute("Select Titles,Types,Boards,StarTime,StopTime,bodys From ["&Isforum&"AdvList] Where ID="& uid)
		If Rs.Eof Then
			SuccessMsg " �������� "
		Else	
		%>
<br>
<Script>
function findobj(n, d) {
	var p, i, x;
	if(!d) d = document;
	if((p = n.indexOf("?"))>0 && parent.frames.length) {
		d = parent.frames[n.substring(p + 1)].document;
		n = n.substring(0, p);
	}
	if(x != d[n] && d.all) x = d.all[n];
	for(i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
	for(i = 0; !x && d.layers && i < d.layers.length; i++) x = findobj(n, d.layers[i].document);
	if(!x && document.getElementById) x = document.getElementById(n);
	return x;
}
</Script>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" name="settings" action="?action=advaddok&edit=1">
  <input type="hidden" name="uid" value="<%=UID%>">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�༭��� - <%=RS(1)%></td>
    </tr>
    <tr>
      <td width="50%" bgcolor="#F8F8F8" ><b>������(����):</b><br>
        <span class="a3">ע��: ������ֻΪʶ����ϲ�ͬ�����Ŀ֮�ã������ڹ������ʾ</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="advtitle" value="<%=RS(0)%>">
      </td>
    </tr>
    <tr>
      <td width="50%" bgcolor="#F8F8F8" valign="top"><b>���Ͷ�ŷ�Χ(��ѡ):</b><br>
        <span class="a3">���ñ����Ͷ�ŵ�ҳ�����̳��Χ�����԰�ס CTRL ��ѡ��ѡ��ȫ����Ϊ������ѡ����Ͷ�ŵķ�Χ</span></td>
      <td bgcolor="#FFFFFF"><select name="advtargets" size="10" multiple="multiple">
          <%
			Dim IsOk 
			If Instr(Rs(2),",")>0 or IsNumEric(Rs(2)) Then
				boards = Split(Rs(2),",")
			Else
				IsOk = Rs(2)
			End if
			Response.Write "<option value=""all"" "
			If IsOk ="all" Or Isok="index" Then Response.Write "selected=""selected"" "
			Response.Write " >&nbsp;&nbsp;> ȫ��</option> " 
			'Response.Write "<option value=""index"" "
			'If IsOk ="index" Then Response.Write "selected=""selected"" "
			'Response.Write " >&nbsp;&nbsp;> ��ҳ</option> " 
			Call BBsList(0)
			%>
        </select>
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��Чʱ��:</b><br>
        <span class="a3">���ù���������ʱ�䣬��ʽ yyyy-mm-dd������Ϊ�����ƽ���ʱ��</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="startime" value="<%=RS(3)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ʱ��:</b><br>
        <span class="a3">���ù���������ʱ�䣬��ʽ yyyy-mm-dd������Ϊ�����ƽ���ʱ��</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="stoptime" value="<%=RS(4)%>">
      </td>
    </tr>
    <tr>
      <td width="50%" bgcolor="#F8F8F8" ><b>��� html ����:</b><br>
        <span class="a3">��ֱ��������Ҫչ�ֵĹ��� html ����</span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="advcode" cols="40" style="height:70;overflow-y:visible;"><%=RS(5)%></textarea>
      </td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="advsubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%		End If
	End if
End Sub
Sub advaddok
	Dim Bodys,textsize,imagewidth,imageheight,imagealt
	If Request.Form("advtitle")&""="" Then  SuccessMsg "��������⡣"
	If Request("edit") = 1 Then
		If Uid = "" Or Not IsNumeric(Uid) Then
			SuccessMsg " �������� "
		Else
			team.execute("Update ["&Isforum&"AdvList] set Titles='"&Replace(Request.Form("advtitle"),"'","")&"',bodys='"&Replace(Request.Form("advcode"),"'","")&"',Boards='"&Replace(Request.form("advtargets")," ","")&"',StarTime='"&Replace(Request.form("startime")," ","")&"',StopTime='"&Replace(Request.form("stoptime")," ","")&"' Where ID="& UID )
		End if
	Else
		Select Case Request.Form("advnewstyle")
			Case "code"
				If Request.Form("advcode")&""="" Then  
					SuccessMsg "���������ݡ�"
				Else
					Bodys = Request.Form("advcode")
				End If
			Case "text"
				If Request.Form("textlink")&""="" or Request.Form("texttitle")&""="" Then  
					SuccessMsg "������������ݡ�"
				End if
				If Request.Form("textsize") <> "" Then
					textsize = " Style=""font-size:"& HtmlEncode(Request.Form("textsize")) &"""  "
				End if
				Bodys = "<a href="""& HtmlEncode(Request.Form("textlink")) &""" target=""_blank"" "&textsize&"> "& HtmlEncode(Request.Form("texttitle")) &" </a>"
			Case "image"
				If Request.Form("imageurl")&""="" or Request.Form("imagelink")&""="" Then  
					SuccessMsg "������������ݡ�"
				End if
				If Request.Form("imagewidth")<>"" and Isnumeric(Request.Form("imagewidth")) Then
					imagewidth = " width="""& Request.Form("imagewidth")&""""
				End if
				If Request.Form("imageheight")<>"" and Isnumeric(Request.Form("imageheight")) Then
					imageheight = " height="""&Request.Form("imageheight")&""""
				End if
				if Request.Form("imagealt")<>"" Then
					imagealt = " alt="""& Request.Form("imagealt")&""" " 
				End if
				Bodys = "<a href="""& HtmlEncode(Request.Form("imagelink")) &""" target=""_blank""><img src="""& HtmlEncode(Request.Form("imageurl")) &""" Border=""0"" align=""absmiddle"" "&imagewidth&" "&imageheight&" "& imagealt&"> </a>"
			Case "flash"
				If Request.Form("flashurl")&""="" Then 
					SuccessMsg "������FLASH��ַ��"
				End If
				If Request.Form("flashwidth")="" or Not IsNumeric(Request.Form("flashwidth")) Then
					SuccessMsg " ��ȱ���Ϊ���� ��"
				ElseIf Request.Form("flashheight")="" or Not IsNumeric(Request.Form("flashheight")) Then
					SuccessMsg " �߶ȱ���Ϊ���� ��"
				Else
					Bodys = "<embed width="""&Request.Form("flashwidth") &""" height="""&Request.Form("flashheight") &""" src="""& HtmlEncode(Request.Form("flashurl")) &""" type=""application/x-shockwave-flash""></embed>"
				End If
		End Select
		team.execute("insert into ["&Isforum&"AdvList] (Dois,Sorts,Titles,Types,StarTime,StopTime,bodys,Boards) values (1,0,'"&Replace(Request.Form("advtitle"),"'","")&"','"&Replace(Request.Form("types"),"'","")&"','"&Replace(Request.Form("startime"),"'","")&"','"&Replace(Request.Form("stoptime"),"'","")&"','"&Replace(Bodys,"'","")&"','"&Replace(Request.form("advtargets")," ","")&"') ")
	End if
	Cache.DelCache("ForumAdvsLoad")
	team.SaveLog ("����������")
	SuccessMsg " ����������  ����ȴ�ϵͳ�Զ����ص� <a href=Admin_Change.asp?action=adv>�������</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=adv>�� " 
End Sub

Sub advadd
	Dim tmp,tmp1,tmp2,tmp3
	Select Case Request("type")
		Case "headerbanner"
			tmp = "ͷ����������ʾ����̳ҳ�����Ϸ���ͨ��ʹ�� 468x60 ͼƬ�� Flash ����ʽ����ǰҳ���ж��ͷ��������ʱ��ϵͳ�����ѡȡ����֮һ��ʾ��"
			tmp1 = "�����ܹ���ҳ��򿪵ĵ�һʱ�佫�������չ��������Ŀ��λ�ã���˳�Ϊ����ҳ�м�λ��ߡ����ʺϽ�����ҵ������Ʒ���ƹ�Ĺ������֮һ��"
			tmp2 = "ͷ��������"
			tmp3 = 1
		Case "footerbanner"
			tmp = "β����������ʾ����̳ҳ�����·���ͨ��ʹ�� 468x60 �������ߴ�ͼƬ��Flash ����ʽ����ǰҳ���ж��β��������ʱ��ϵͳ�����ѡȡ����֮һ��ʾ�� "
			tmp1 = "��ҳ��ͷ�����в���ȣ�ҳ��β����չ�ֻ�����Խϵͣ�ͨ��������������ߵķ��У�ͬʱ�ֻ����ܹ��������жԹ�����ݸ���Ȥ�����ڣ�����ʺ����Զ��º͵��ƹ㡣"
			tmp2 = "β��������"
			tmp3 = 2
		Case "text"
			tmp = "ҳ�����ֹ���Ա�����ʽ����ʾ����ҳ�������б��������������ҳ������Ϸ���ͨ��ʹ�����ֵ���ʽ��Ҳ��ʹ��СͼƬ�� Flash����ǰҳ���ж�����ֹ��ʱ��ϵͳ���Ա�����ʽ�����趨����ʾ˳��ȫ��չ�֣�ͬʱ�ܹ��Ա�������� 3~5 �ķ�Χ�ڶ�̬�Ų������Զ�ʵ����ѵĹ������Ч����"
			tmp1 = "���ڴ�����ͨ����������ʽչ�֣��������ڵĽϿ��ϵ�ҳ��λ�ã�ʹ�ô������Ϊ�˷����߱ض�������֮һ��ͬһҳ����Գ��ֶ��ʮ�������ֹ������ԣ�Ҳ����������һ��ƽ�񻯵��Լ۱Ƚϸߵ��ƹ㷽ʽ��ͬʱ����������̳����������͹���֮�á�"
			tmp2 = "ҳ�����ֹ��"
			tmp3 = 3
		Case "thread"
			tmp = "���ڹ����ʾ�����ӱ�����Ϸ���ͨ��ʹ�����ֵ���ʽ����ǰҳ���ж�����ڹ��ʱ��ϵͳ����г�ȡ��ÿҳ������ȵ���Ŀ���������ʾ��"
			tmp1 = "������������̳����ĵ���ɲ��֣�λ�����������Ϸ������ڹ�棬������û������������ʱ��Ȼ�ı����ܣ�����������ŵ����ԣ��ʺ����ض����ݵ���Ч�ƹ㣬Ҳ��������̳����������͹���֮�á��������ö������ڹ����ʵ�ֹ�����ݵĲ��컯���Ӷ�������������ߵ�ע������"
			tmp2 = "���ڹ��"
			tmp3 = 4
		Case "float"
			tmp = "Ư�����չ����ҳ�����½ǣ���ҳ�����ʱ���������ƶ��Ա���ԭ����λ�ã�ͨ��ʹ��СͼƬ�� Flash ����ʽ����ǰҳ���ж��Ư�����ʱ��ϵͳ�����ѡȡ����֮һ��ʾ��"
			tmp1 = " Ư������ǽ���ǿ����ҵ�ƹ����Ч�ֶΣ�����ҳ���еĸ����ԣ�ʹ����̶���ͼƬ��������ȣ������ױ���ע������Ϊ��ˣ�����ǿ���ԵĹ�עҲ�������¶Դ˹�����ݲ�����Ȥ�ķ����ߵķ��С���ע�ⲻҪ�������ͼƬ�� Flash ��Ư��������ʽ��ʾ������Ӱ��ҳ���Ķ���"
			tmp2 = "Ư�����"
			tmp3 = 5
		Case "couplebanner"
			tmp = "��������Գ�����ͼƬ����ʽ��ʾ��ҳ�涥�����࣬����һ��������ͨ��ʹ�ÿ�С�ߴ�ĳ�����ͼƬ�� Flash ����ʽ���������һ��ֻ��ʹ������Լ��������ȵ������ʹ�ã���ʹ�ó��� 90% ���ϵİٷֱ�Լ���������ʱ�����ܻ�Ӱ������ߵ���������������������������С�� 800 ����ʱ���Զ�����ʾ�����档��ǰҳ���ж���������ʱ��ϵͳ�����ѡȡ����֮һ��ʾ��"
			tmp1 = "�����������ֻչ���ڸ߷ֱ���(1024x768 �����)��Ļ�����ֻ࣬ռ��ҳ��Ŀհ�������˲������·����߷��У��ܹ����õ�ͻ���ƹ����ݡ������ڶԷֱ��ʺ�������ȵ�����Ҫ��ʹ�ù������ڱ����޷��ﵽ 100%��"
			tmp2 = "�������"
			tmp3 = 6
		Case "affbanner"
			tmp = "����λ����Գ�����ͼƬ����ʽ��ʾ��ҳ����࣬ͨ��ʹ�ÿ�С�ߴ�ĳ�����ͼƬ�� Flash ����ʽ������λ���һ��ռ�ù�������ࡣ��ǰҳ���ж������λ���ʱ��ϵͳ�����ѡȡ����֮һ��ʾ��"
			tmp1 = "����λ�������ֻչ���ڹ�����������һ����Ƭ���ԣ��������ڼĴ��ڹ��������Ե������һ��Ӱ�졣"
			tmp2 = "����λ���"
			tmp3 = 7

		Case "threadleft"
			tmp = "�������ӹ���Գ�����ͼƬ����ʽ��ʾ��ҳ����࣬ͨ��ʹ�ÿ�С�ߴ�ĳ�����ͼƬ�� Flash ����ʽ����������λ���һ��ռ�������������Ҳࡣ����ʱȽϴ�,�����ʵ������ÿ��Դ����ܴ������."
			tmp1 = "��������λ�������ֻչ��������������������һ����Ƭ���ԣ��������ڼĴ����������������Ե������һ��Ӱ�졣"
			tmp2 = "��������λ���"
			tmp3 = 8
	End Select
%>
<br>
<Script>
function findobj(n, d) {
	var p, i, x;
	if(!d) d = document;
	if((p = n.indexOf("?"))>0 && parent.frames.length) {
		d = parent.frames[n.substring(p + 1)].document;
		n = n.substring(0, p);
	}
	if(x != d[n] && d.all) x = d.all[n];
	for(i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
	for(i = 0; !x && d.layers && i < d.layers.length; i++) x = findobj(n, d.layers[i].document);
	if(!x && document.getElementById) x = document.getElementById(n);
	return x;
}
</Script>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>չ�ַ�ʽ: <%=tmp%>
      </ul>
      <ul>
        <li>��ֵ����: <%=tmp1%>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" name="settings" action="?action=advaddok">
  <input type="hidden" name="types" value="<%=tmp3%>">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">��ӹ�� - <%=tmp2%></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������(����):</b><br>
        <span class="a3">ע��: ������ֻΪʶ����ϲ�ͬ�����Ŀ֮�ã������ڹ������ʾ</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="advtitle" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>���Ͷ�ŷ�Χ(��ѡ):</b><br>
        <span class="a3">���ñ����Ͷ�ŵ�ҳ�����̳��Χ�����԰�ס CTRL ��ѡ��ѡ��ȫ����Ϊ������ѡ����Ͷ�ŵķ�Χ</span></td>
      <td bgcolor="#FFFFFF"><select name="advtargets" size="10" multiple="multiple">
          <option value="all" selected="selected">&nbsp;> ȫ��</option>
          <% BBsList(0) %>
        </select></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��Чʱ��:</b><br>
        <span class="a3">���ù���������ʱ�䣬��ʽ yyyy-mm-dd������Ϊ�����ƽ���ʱ��</span></td>
      <td bgcolor="#FFFFFF"> <input type="text" size="30" name="startime" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ʱ��:</b><br>
        <span class="a3">���ù���������ʱ�䣬��ʽ yyyy-mm-dd������Ϊ�����ƽ���ʱ��</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="stoptime" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>ѡ����ģʽ:</b><br>
        <span class="a3">��ѡ������Ĺ��չ�ַ�ʽ�����Է���Ĳ�������档</span></td>
      <td bgcolor="#FFFFFF"><select name="advnewstyle" onChange="var styles;var key;styles=new Array('code','text','image','flash'); for(key in styles) {var obj=findobj('style_'+styles[key]); obj.style.display=styles[key]==this.options[this.selectedIndex].value?'':'none';}">
          <option value="code"> ����</option>
          <option value="text"> ����</option>
          <option value="image"> ͼƬ</option>
          <option value="flash"> Flash</option>
        </select></td>
    </tr>
  </table>
  <div id="style_code" style=""><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">Html ����</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" valign="top"><b>��� html ����:</b><br>
          <span class="a3">��ֱ��������Ҫչ�ֵĹ��� html ����</span></td>
        <td bgcolor="#FFFFFF"><textarea rows="5" name="advcode" cols="40" style="height:70;overflow-y:visible;"></textarea></td>
      </tr>
    </table>
  </div>
  <div id="style_text" style="display: none"><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">���ֹ��</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>��������(����):</b><br>
          <span class="a3">���������ֹ�����ʾ����</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="texttitle" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>��������(����):</b><br>
          <span class="a3">���������ֹ��ָ��� URL ���ӵ�ַ,����http://��ͷ.</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="textlink" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>���ִ�С(ѡ��):</b><br>
          <span class="a3">���������ֹ���������ʾ���壬��ʹ�� pt��px��em Ϊ��λ</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="textsize" value="">
        </td>
      </tr>
    </table>
  </div>
  <div id="style_image" style="display:none"><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">ͼƬ���</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>ͼƬ��ַ(����):</b><br>
          <span class="a3">������ͼƬ����ͼƬ���õ�ַ</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imageurl" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>ͼƬ����(����):</b><br>
          <span class="a3">������ͼƬ���ָ��� URL ���ӵ�ַ</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imagelink" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>ͼƬ���(ѡ��):</b><br>
          <span class="a3">������ͼƬ���Ŀ�ȣ���λΪ����</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imagewidth" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>ͼƬ�߶�(ѡ��):</b><br>
          <span class="a3">������ͼƬ���ĸ߶ȣ���λΪ����</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imageheight" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>ͼƬ�滻����(ѡ��):</b><br>
          <span class="a3">������ͼƬ���������ͣ������Ϣ</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imagealt" value="">
        </td>
      </tr>
    </table>
  </div>
  <div id="style_flash" style="display: none"><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">Flash ���</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>Flash ��ַ(����):</b><br>
          <span class="a3">������ Flash ���ĵ��õ�ַ</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="flashurl" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>Flash ���(����):</b><br>
          <span class="a3">������ Flash ���Ŀ�ȣ���λΪ����</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="flashwidth" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>Flash �߶�(����):</b><br>
          <span class="a3">������ Flash ���ĸ߶ȣ���λΪ����</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="flashheight" value="">
        </td>
      </tr>
    </table>
  </div>
  <br>
  <br>
  <center>
    <input type="submit" name="advsubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub advok
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"AdvList] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"AdvList] set Dois="&Cid(Request.Form("availablenew"&i+1))&",Sorts="&CID(Request.Form("displayordernew"&i+1))&",Titles='"&Replace(Request.Form("titlenew"&i+1),"'","")&"' Where ID="&newid(i))
		Next
	End if
	Cache.DelCache("ForumAdvsLoad")
	team.SaveLog ("����������")
	SuccessMsg " ����������  ����ȴ�ϵͳ�Զ����ص� <a href=Admin_Change.asp?action=adv>�������</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=adv>�� " 
End Sub

Sub adv %>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>�������;���������ڵ�λ�á�</li>
      </ul></td>
  </tr>
</table>
<BR>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr>
    <td colspan="2" class="a1">��ӹ��</td>
  </tr>
  <tr>
    <td colspan="2" class="a4">
	  <input type="button" value="ͷ��������" onClick="window.location='?action=advadd&type=headerbanner';">
      &nbsp;
      <input type="button" value="β��������" onClick="window.location='?action=advadd&type=footerbanner';">
      &nbsp;
      <input type="button" value="ҳ�����ֹ��" onClick="window.location='?action=advadd&type=text';">
      &nbsp;
      <input type="button" value="���ڹ��" onClick="window.location='?action=advadd&type=thread';">
      &nbsp;
      <input type="button" value="Ư�����" onClick="window.location='?action=advadd&type=float';">
      &nbsp;
      <input type="button" value="�������" onClick="window.location='?action=advadd&type=couplebanner';">
	  &nbsp;
      <input type="button" value="����λ���" onClick="window.location='?action=advadd&type=affbanner';">
	  &nbsp;
	  <input type="button" value="��������λ���" onClick="window.location='?action=advadd&type=threadleft';"></td>
	  
	  </td>
  </tr>
</table>
<BR>
<form method="post" action="?action=advok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
      <td width="48"><input type="checkbox" name="chkall" class="a1" onClick="checkall(this.form,'delete')">
        ɾ?</td>
      <td width="5%">����</td>
      <td width="8%">��ʾ˳��</td>
      <td width="15%">����</td>
      <td width="20%">����</td>
      <td width="15%">��ʼʱ��</td>
      <td width="15%">��ֹʱ��</td>
      <td width="15%">Ͷ�ŷ�Χ</td>
      <td width="6%">�༭</td>
    </tr>
    <%
	dim Rs,i,tmp
	i = 0
	Set Rs=team.execute("Select ID,Dois,Sorts,Titles,Types,StarTime,StopTime,Boards From ["&Isforum&"AdvList] Order By Sorts Desc")
	Do While Not Rs.Eof
		i = i+1
		Select Case RS(4)
			Case 1
				tmp = "ͷ��������"
			Case 2
				tmp = "β��������"
			Case 3
				tmp = "ҳ�����ֹ��"
			Case 4
				tmp = "���ڹ��"
			Case 5
				tmp = "Ư�����"
			Case 6
				tmp = "�������"
			Case 7
				tmp = "���������"
			Case 8
				tmp = "�������������"
		End Select				
	%>
    <tr align="center" class="a4">
      <input type="hidden" name="newid" value="<%=RS(0)%>">
      <td><input type="checkbox" name="deleteid" value="<%=RS(0)%>"></td>
      <td><input type="checkbox" name="availablenew<%=i%>" value="1" <%If Rs(1)=1 Then%>checked<%End if%>></td>
      <td><input type="text" size="2" name="displayordernew<%=i%>" value="<%=RS(2)%>"></td>
      <td><input type="text" size="15" name="titlenew<%=i%>" value="<%=RS(3)%>"></td>
      <td><%=tmp%></td>
      <td><%=IIF(RS(5)&""="","������",RS(5))%></td>
      <td><%=IIF(RS(6)&""="","������",RS(6))%></td>
      <td>
	  <%if Rs(7) = "all" Then 
			Echo "ȫ��" 
		Else 
			Echo RS(7) 
		End if %></td>
      <td><a href="?action=advedit&uid=<%=RS(0)%>">[����]</a></td>
    </tr>
    <%   Rs.Movenext
	Loop
	Rs.Close:Set RS=Nothing
	%>
  </table>
  <br>
  <center>
    <input type="submit" name="forumlinksubmit" value="�� ��">
  </center>
</form>
<%
End Sub

Sub forumlinksok
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"Link] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"Link] set Name='"&Replace(Request.Form("name"&i+1),"'","")&"',Url='"&Replace(Request.Form("url"&i+1),"'","")&"',Intro='"&Replace(Request.Form("note"&i+1),"'","")&"',SetTops="&Cid(Request.Form("displayorder"&i+1))&",logo='"&Replace(Request.Form("logo"&i+1),"'","")&"' Where ID="&newid(i))
		Next
		If Request.Form("newname")<>"" and Request.Form("newurl")<>"" Then
			team.execute("insert into ["&Isforum&"Link] (Name,Url,Intro,SetTops,logo) values ('"&Replace(Request.Form("newname"),"'","")&"','"&Replace(Request.Form("newurl"),"'","")&"','"&Replace(Request.Form("newnote"),"'","")&"',"&CID(Request.Form("newdisplayorder"))&",'"&Replace(Request.Form("newlogo"),"'","")&"') ")
		End if
	End If
	Cache.DelCache("Superlink")
	team.SaveLog ("������������")
	SuccessMsg " ��������������� ����ȴ�ϵͳ�Զ����ص� <a href=Admin_Change.asp?action=forumlinks>��������</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=forumlinks>�� "
End Sub

Sub forumlinks %>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>�������������ҳ��ʾ������̳��������и���ɾ�����ɡ�</li>
      </ul>
      <ul>
        <li>δ��д����˵������Ŀ���Խ�������ʾ��</li>
      </ul>
      <ul>
        <li>δ��дlogo ��ַ����Ŀ��������������ʾ��</li>
      </ul>
      <ul>
        <li>��̳ URL���� http:// ��ʼ����Ȼ�����������޷����ʵ������</li>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=forumlinksok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="6">�������ӱ༭</td>
    </tr>
    <tr align="center" class="a3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)">
        ɾ?</td>
      <td>��ʾ˳��</td>
      <td>��̳����</td>
      <td>��̳ URL</td>
      <td>����˵��</td>
      <td>logo ��ַ(��ѡ)</td>
    </tr>
    <%Dim Rs,i
	i=0
	Set Rs=team.execute("Select ID,Name,Url,Intro,SetTops,logo From ["&Isforum&"Link] Order By SetTops Desc")
	Do While Not Rs.Eof
		i= i+1
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td bgcolor="#F8F8F8"><Input Name="newid" type="hidden" value="<%=RS(0)%>">
        <input type="checkbox" name="deleteid" value="<%=RS(0)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="3" name="displayorder<%=i%>" value="<%=RS(4)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="15" name="name<%=i%>" value="<%=RS(1)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="15" name="url<%=i%>" value="<%=RS(2)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="15" name="note<%=i%>" value="<%=RS(3)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="15" name="logo<%=i%>" value="<%=RS(5)%>"></td>
    </tr>
    <%
		Rs.Movenext
	Loop
	Rs.close:Set Rs=Nothing
	%>
    <tr>
      <td colspan="6" class="a4" height="5"></td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td>����:</td>
      <td><input type="text" size="3"	name="newdisplayorder"></td>
      <td><input type="text" size="15" name="newname"></td>
      <td><input type="text" size="15" name="newurl"></td>
      <td><input type="text" size="15" name="newnote"></td>
      <td><input type="text" size="15" name="newlogo"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="forumlinksubmit" value="�� ��">
  </center>
</form>
<%
End Sub


Sub newsannouncements
	Dim newsubject,newcss,newendtime,newmessage
	Newsubject = HtmlEncode(Trim(Request.Form("newsubject")))
	Newmessage = team.checkStr(Trim(Request.Form("newmessage")))
	If Newsubject &""="" Then 
		SuccessMsg "������ⲻ��Ϊ�ա�"
	ElseIf Newmessage &""="" Then 
		SuccessMsg "�������ݲ���Ϊ�ա�"	
	Else
		If Trim(Request.Form("newendtime"))<>"" Then
			If Not Isdate(Trim(Request.Form("newendtime"))) Then
				SuccessMsg "����ʱ��ĸ�ʽ����ȷ���������ʵ������ڸ�ʽ��"
			End If
		End if
		If Request("edit") = 1 Then
			team.execute(" Update ["&Isforum&"Affiche] Set Affichetitle='"&Newsubject&"',Affichecontent='"&Newmessage&"',Afficheman='"&TK_UserName&"',Afficheinfo='"&Replace(Trim(Request.Form("newcss")),"'","")&"',Lifetime='"&Trim(Request.Form("newendtime"))&"',Affichetime="&SqlNowString&" Where ID="&UID)
			Cache.DelCache("BBsAffiche")
			SuccessMsg "����༭��ɣ���ȴ�ϵͳ�Զ����ص� <a href=Admin_Change.asp?action=announcements>��̳����</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=announcements>��"	
		Else
			team.execute("insert into ["&Isforum&"Affiche] (Affichetitle,Affichecontent,Afficheman,Afficheinfo,Lifetime,Affichetime) values ('"&Newsubject&"','"&Newmessage&"','"&TK_UserName&"','"&Replace(Trim(Request.Form("newcss")),"'","")&"','"&Trim(Request.Form("newendtime"))&"',"&SqlNowString&") ")
			Cache.DelCache("BBsAffiche")
			SuccessMsg "�µĹ��淢����ɣ���ȴ�ϵͳ�Զ����ص� <a href=Admin_Change.asp?action=announcements>��̳����</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=announcements>��"	
		End If
	End If
	team.SaveLog ("��������")
End Sub

Sub announcementsok
	Dim ho
	If request.form("deleteid") = "" Then
		SuccessMsg " ��ѡ����Ҫɾ���Ĺ��� "
	Else
		for each ho in request.form("deleteid")
			team.execute("Delete from ["&Isforum&"Affiche] Where ID="&ho)
		next
	End If
	Cache.DelCache("BBsAffiche")
	team.SaveLog ("����ɾ��")
	SuccessMsg " ����ɾ����ɣ���ȴ�ϵͳ�Զ����ص� <a href=Admin_Change.asp?action=announcements>��̳����</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=announcements>��"
End sub

Sub  announcements
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>�������ı��⣬���ɶԹ�����б༭��
        <li>CSSЧ�� ��font-weight: bold; �������֡�color: #FF0000; ������ɫ ��
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=announcementsok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="7">��̳����༭</td>
    </tr>
    <tr align="center" class="a3">
      <td width="48"><input type="checkbox" name="chkall" onClick="checkall(this.form)">
        ɾ?</td>
      <td>����</td>
      <td>����</td>
      <td>����</td>
      <td>��ʼʱ��</td>
      <td>��ֹʱ��</td>
    </tr>
    <%
	Dim Rs
	Set Rs=team.execute("Select ID,Affichetitle,Affichecontent,Afficheman,Afficheinfo,Lifetime,Affichetime From ["&Isforum&"Affiche] order By Id Asc")
	Do While Not Rs.Eof
	%>
    <tr align="center">
      <td bgcolor="#F8F8F8"><input type="checkbox" name="deleteid" value="<%=RS(0)%>"></td>
      <td bgcolor="#FFFFFF"><a href="./Profile.asp?username=<%=RS(3)%>" target="_blank"><%=RS(3)%></a></td>
      <td bgcolor="#F8F8F8"><a href="?action=announcements&uid=<%=rs(0)%>&edit=1" title="����༭�˹���"><span Style="<%=Rs(4)%>"><%=RS(1)%></span></a></td>
      <td bgcolor="#FFFFFF"><a href="?action=announcements&uid=<%=rs(0)%>&edit=1" title="����༭�˹���"><%=CutStr(RS(2),15)%></a></td>
      <td bgcolor="#F8F8F8"><%=RS(6)%></td>
      <td bgcolor="#FFFFFF"><% If RS(5)&"" = "" Then:Echo "������": Else:Echo RS(4):End If%></td>
    </tr>
    <%	Rs.MoveNext
	Loop
	Rs.close:set Rs=nothing%>
  </table>
  <br>
  <center>
    <input type="submit" name="announcesubmit" value="�� ��">
  </center>
</form>
<br>
<%
If request("edit")=1 Then
	Dim Rs1
	If UID="" Or Not IsNumeric(UID) Then
		SuccessMsg " ��������"
	Else
		Set Rs1=team.execute("Select ID,Affichetitle,Affichecontent,Afficheman,Afficheinfo,Lifetime,Affichetime From ["&Isforum&"Affiche] Where ID="& UID)
		If Rs1.eof Then 
			SuccessMsg " ��������"
		Else
			Echo " <form method=""post"" action=""?action=newsannouncements&edit=1&uid="&UID&"""> "
		End If
	End if
Else
	Echo "<form method=""post"" action=""?action=newsannouncements"">"
End If
If request("edit")=1 Then
%>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td colspan="2">�༭��̳����</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>����:</b></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newsubject" Value = "<%=RS1(1)%>"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>������ɫ:</b><BR>
      <span class="a3">֧��ʹ��CSSЧ��</span></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newcss" Value = "<%=RS1(4)%>"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>����ʱ��:</b><br>
      ��ʽ: yyyy-mm-dd</td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newendtime" Value = "<%=RS1(5)%>">
      ����Ϊ������</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>����:</b><br>
      ��������֧��UBB����<BR>
      UBB����ʹ����鿴<a href="../Help.asp?page=mise#1"> <B>UBBָ��</B> </a>
    <td width="60%" bgcolor="#FFFFFF"><textarea name="newmessage" cols="60" rows="10"><%=Server.htmlEncode(RS1(2))%></textarea></td>
  </tr>
</table>
<%Else%>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td colspan="2">�����̳����</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>����:</b></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newsubject"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>������ɫ:</b><BR>
      <span class="a3">֧��ʹ��CSSЧ��</span></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newcss"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>����ʱ��:</b><br>
      ��ʽ: yyyy-mm-dd</td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newendtime">
      ����Ϊ������</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>����:</b><br>
      ��������֧��UBB����<BR>
      UBB����ʹ����鿴<a href="../Help.asp?page=mise#1"> <B>UBBָ��</B> </a>
    <td width="60%" bgcolor="#FFFFFF"><textarea name="newmessage" cols="60" rows="10"></textarea></td>
  </tr>
</table>
<%End if%>
<br>
<center>
<input type="submit" name="addsubmit" value="�� ��">
</form>
<br>
<br>
<%
End Sub

Sub medalsok
	Dim MedalName,MedalSet,Medalimg
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"Medals] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"Medals] set MedalName='"&Request.Form("MedalName"&i+1)&"',Medalimg='"&Request.Form("Medalimg"&i+1)&"',MedalSet="&Cid(Request.Form("MedalSet"&i+1))&" Where ID="&newid(i))
		Next
		If Request.Form("newname")<>"" and Request.Form("newimage")<>"" Then
			team.execute("insert into ["&Isforum&"Medals] (MedalName,Medalimg,MedalSet) values ('"&Request.Form("newname")&"','"&Request.Form("newimage")&"',"&CID(Request.Form("availablenew"))&" ) ")
		End if
	End If
	team.SaveLog ("ѫ������")
	SuccessMsg " ѫ��������� �� "
End Sub

Sub medals	
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>�������������ÿ��԰䷢���û���ѫ����Ϣ��ѫ��ͼƬ������дͼƬ�ļ�����������ӦͼƬ�ļ��ϴ��� ../images/plus Ŀ¼�С�
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=medalsok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="5">ѫ�±༭</td>
    </tr>
    <tr align="center" class="a3">
      <td><input type="checkbox" name="chkall" class="a4" onClick="checkall(this.form, 'delete')">
        ɾ?</td>
      <td>����</td>
      <td>ͼƬ��ַ</td>
      <td>ѫ��ͼƬ</td>
      <td>����</td>
    </tr>
    <%
	Dim Rs,i
	i=0
	Set Rs=team.execute("Select ID,MedalName,Medalimg,MedalSet From ["&Isforum&"Medals] Order By ID asc")
	Do While Not Rs.Eof
		i = i+1
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td bgcolor="#F8F8F8" width="48"><input type="checkbox" name="deleteid" value="<%=RS(0)%>">
        <Input Name="newid" type="hidden" value="<%=RS(0)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="MedalName<%=i%>" value="<%=RS(1)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="30" name="Medalimg<%=i%>" value="<%=RS(2)%>"></td>
      <td bgcolor="#FFFFFF"><img src="../images/plus/<%=RS(2)%>" align="absmiddle"> </td>
      <td bgcolor="#F8F8F8"><input type="checkbox" name="MedalSet<%=i%>" value="1" <%If Rs(3)=1 then%>checked<%end if%>></td>
    </tr>
    <%
		Rs.MoveNext
	Loop
	Rs.close:Set Rs=Nothing
	%>
    <tr>
      <td colspan="5" class="a4" height="2"></td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td>����:</td>
      <td><input type="text" size="30" name="newname"></td>
      <td><input type="text" size="30" name="newimage"></td>
      <td>&nbsp;</td>
      <td><input type="checkbox" name="availablenew" value="1"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="medalsubmit" value="�� ��">
  </center>
</form>
<%
End Sub
Sub BBsList(V)
	Dim SQL,ii,RS,i
	Set Rs=Team.Execute("Select ID,BBSname,Followid From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		If RS(2)=0 Then 
			Echo "<optgroup label="""&Rs(1)&""">"
		Else
			Echo "<option value="&RS(0)&" " 
			If Isarray(boards) Then
				for i=0 to ubound(boards)
					If RS(0) = int(boards(i)) Then Echo "selected=""selected"" " 
				next
			end if
			Echo " >"&String(ii,"��") & RS(1)&"</option>"
		End if
		ii=ii+1
		BBsList RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub

Sub Menus()
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Name,url,followid,SortNum,Newtype From "&IsForum&"Menu Where Followid=0 Order By SortNum")
	If Rs.Eof Then
		Echo "<BR><ul><center> Ŀǰû������κβ˵� </center></ul> "
	End if
	Do While Not RS.Eof
		Echo "<tr class=""a4"" align=""center""><td width=""10""> <Input Name=UID value="&RS(0)&" type=hidden> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&"></td><td width=""100""> ����  <input type=text name=SortNum Value="&RS(4)&" Size=""1""> </td><td width=""50%"" align=""left""> �� <a target=_blank href=../"&RS(2)&"><b>"&RS(1)&"</b></a> </td><td> "
		If Rs(5)=1 Then	
			Echo "ǰ̨�˵�"
		Else
			Echo "��̨�˵�"
		End if
		Echo " </td><td> <a href=""?action=menuadd&fid="&RS(0)&"&Mid="&Rs(5)&""" title=""��ӱ�������¼��˵�"">[���]</a> <a href=""?action=menuadd&uid="&RS(0)&"&edit=1&Mid="&Rs(5)&""" title=""�༭���˵�����"">[�༭]</a> </td></tr>"
		Call Menus_1(Rs(0))
		Echo " "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub

Sub Menus_1(a)
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Name,url,followid,SortNum,Newtype From "&IsForum&"Menu Where Followid="&a&" Order By SortNum")
	Do While Not RS.Eof
		Echo "<tr class=""a4"" align=""center""><td width=""10""> <Input Name=UID value="&RS(0)&" type=hidden> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&"></td><td width=""100""> ����  <input type=text name=SortNum Value="&RS(4)&" Size=""1""> </td><td width=""50%"" align=""left"">���� ��<a target=_blank href=../"&RS(2)&"><b>"&RS(1)&"</b></a> </td><td> "
		If Rs(5)=1 Then	
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

footer()
%>
