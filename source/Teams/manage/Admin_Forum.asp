<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim ii,ID
Dim Admin_Class
Call Master_Us()
Header()
ii=0:ID=Request("id")
Admin_Class=",2,"
Call Master_Se()
team.SaveLog ("��̳���� [�������༭��� ] �޸�")
Select Case Request("Action")
	Case "Manages"
		Call Manages	'������
	Case "FindForum"
		ID= Request.Form("ForumID")
		Call Manages	'���Ұ��
	Case "Forumadd"
		Call Forumadd	'�༭���
	Case "ForumSort"	'�������
			Dim Myid,MySortNum
			Myid=Split(Request.Form("UID"),",")
			MySortNum=Split(Request.Form("SortNum"),",")
			For U=0 To Ubound(Myid)
				team.Execute("Update "&IsForum&"Bbsconfig set SortNum="&MySortNum(U)&" where ID="&Myid(U))
			Next
			Cache.DelCache("BoardLists")
			SuccessMsg("�������!")	
	Case "ForumAddok"	
		Dim fup
		fup=ReQuest.Form("fup")
		If Request.Form("newforum")="" Then Error2 "������Ʋ���Ϊ��!"
		Select Case Request("add")
			Case "Forum_0"
				team.Execute("insert into "&IsForum&"Bbsconfig(Followid,bbsname,Board_Last,Board_Setting,today,toltopic,tolrestore,hide,Board_Model,SortNum) values (0,'"&Replace(Request.Form("newforum"),"'","")&"','��������$@$ - $@$"&Now&"','0$$$0$$$0$$$0$$$1$$$1$$$1$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0',0,0,0,0,0,0) ")
				Cache.DelCache("BoardLists")
				SuccessMsg("һ������ ["&Request.Form("newforum")&"] ��ӳɹ�<BR><a href=Admin_Forum.asp>��ת�����˵�������ϸ����������</a>!")
			Case "Forum_1"
				team.Execute("insert into "&IsForum&"Bbsconfig(Followid,bbsname,Board_Last,Board_Setting,today,toltopic,tolrestore,hide,Board_Model,SortNum) values ("&Request.Form("fup")&",'"&Replace(Request.Form("newforum"),"'","")&"','��������$@$ - $@$"&Now&"','0$$$0$$$0$$$0$$$1$$$0$$$1$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0',0,0,0,0,0,0) ")
				Cache.DelCache("BoardLists")
				SuccessMsg("���� ["&Request.Form("newforum")&"] ��ӳɹ�<BR><a href=Admin_Forum.asp>��ת�����˵�������ϸ����������</a>!")
		End Select
	Case "Forumeditok"
		Dim My_Board_Setting,U,My_ExtCredit,ExtCredits
		If Request.Form("BbsName")="" Then Error2 "�������Ʋ���Ϊ��!"
		My_Board_Setting=""
		My_ExtCredit = ""
		ExtCredits= Split(team.Club_Class(21),"|")
		For U=0 to Ubound(ExtCredits)
			If U=0 Then
				My_ExtCredit=Replace(Request.Form("ExtCredits0_0"),",","")&","&Replace(Request.Form("ExtCredits0_1"),",","")&","&Replace(Request.Form("ExtCredits0_2"),",","")
			Else				
				My_ExtCredit=My_ExtCredit & "|"&Replace(Request.Form("ExtCredits"&U&"_0"),",","")&","&Replace(Request.Form("ExtCredits"&U&"_1"),",","")&","&Replace(Request.Form("ExtCredits"&U&"_2"),",","")
			End If
		Next
		For U=0 to 19
			If U=0 Then
				My_Board_Setting=Replace(Request.Form("Board_Setting(0)"),"$$$","")
			ElseIf U=14 Then				
				My_Board_Setting=My_Board_Setting & "$$$"&My_ExtCredit
			Else
				My_Board_Setting=My_Board_Setting & "$$$"&Replace(Request.Form("Board_Setting("&U&")"),"$$$","")
			End If
		Next
		team.Execute("Update "&IsForum&"Bbsconfig Set bbsname='"&HTMLEncode(Request.Form("BbsName"))&"',Readme='"&Replace(Trim(Request.Form("Readme")),"'","''")&"',Icon='"&Replace(Trim(Request.Form("Icon")),"'","''")&"',Board_Key='"&Replace(Trim(Request.Form("Board_Key")),"'","''")&"',Hide="&HTMLEncode(Trim(Request.Form("Hide")))&",Pass='"&HTMLEncode(Trim(Request.Form("Pass")))&"',Followid="&Cid(Request.Form("fupnew"))&",Board_URL='"&HtmlEncode(Trim(Request.Form("Board_URL")))&"',Board_Setting='"&My_Board_Setting&"',Lookperm='"&Replace(Request.Form("lookperm")," ","")&",',Postperm='"&Replace(Request.Form("postperm")," ","")&",',Downperm='"&Replace(Request.Form("downperm")," ","")&",',upperm='"&Replace(Request.Form("upperm")," ","")&",' Where ID="&ID)
		Cache.DelCache("ForumsBoards_"&ID)
		Cache.DelCache("ThreadBoards_"&ID)
		Cache.DelCache("SaveThreadBoards_"&ID)
		Cache.DelCache("Boards_"&ID)
		Cache.DelCache("BoardLists")
		SuccessMsg("���� ["&Request.Form("BbsName")&"] �༭�ɹ�����ȴ�ϵͳ�Զ����ص� <a href=Admin_Forum.asp>�༭���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Forum.asp>��")
	Case "SetModerators"
		Call SetModerators
	Case "ModelSet_0"
		If ID="" or (Not isNumeric(ID)) Then 
			SuccessMsg " ID��������! "
		Else
			team.execute("Update ["&Isforum&"BbsConfig] Set Board_Model = 1 Where Id="&ID&" or Followid="&ID)
			Cache.DelCache("BoardLists")
			SuccessMsg "�Ѿ������������з�ʽ�޸�Ϊ���ģʽ����ȴ�ϵͳ�Զ����ص� <a href=Admin_Forum.asp>�༭���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Forum.asp>��"
		End If
	Case "ModelSet_1"
		If ID="" or (Not isNumeric(ID)) Then 
			SuccessMsg " ID��������! "
		Else
			team.execute("Update ["&Isforum&"BbsConfig] Set Board_Model = 0 Where Id="&ID&" or Followid="&ID)
			Cache.DelCache("BoardLists")
			SuccessMsg "�Ѿ������������з�ʽ�޸�Ϊ��׼ģʽ����ȴ�ϵͳ�Զ����ص� <a href=Admin_Forum.asp>�༭���</a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_Forum.asp>��"
		End If
	Case "DelForum"
		Call DelForums
	Case "ServerDelForum"
		Dim Rs
		If ID="" or (Not isNumeric(ID)) Then 
			SuccessMsg " ID��������! "
		Else
			team.Execute("Delete From "&IsForum&"Bbsconfig Where ID="&ID)
			Set Rs = team.execute("Select ID,ReList From ["&IsForum&"Forum] Where forumid=" & ID)
			Do While Not Rs.Eof
				team.Execute("Delete From ["&IsForum & RS(1) &"] Where topicid="& RS(0) )
				Rs.MoveNext
			Loop
			Rs.close:Set Rs=Nothing
			team.Execute("Delete From ["&IsForum&"Forum] Where forumid="&ID)
			Cache.DelCache("BoardLists")
			Cache.DelCache("ForumsBoards_"&ID)
			Cache.DelCache("Boards_"&ID)
			SuccessMsg("ɾ����̳�ɹ�<BR><a href=Admin_Forum.asp>��ת�����˵�������������</a>��ȴ�3���Ӻ�ϵͳ�Զ�ת�����˵����档<meta http-equiv=refresh content=3;url=Admin_Forum.asp>")
		End if
	Case Else
		Call Main()
End Select

Sub Main()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2" >
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
   <td><BR><ul>
        <li>����Ҫ�����ύһ�����ࡣ</li>
      </ul>
      <ul>
        <li>ÿ���������Ĺ����ܺ����������¼����湦�ܣ����������Ƽ���Ҫ����������</li>
      </ul>
      <ul>
        <li>�����Զ��ڡ���ʾ˳���������̳��������ÿ������������ 0 ��ʼ��</li>
      </ul></td>
  </tr>
</table><BR>
<form method="post" action="?Action=FindForum">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="5">������̳</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="15%">����:</td>
      <td bgcolor="#FFFFFF" width="70%"><input type="text" name="ForumID" value="����ҵ���̳ID" size="40" OnFocus="this.value = ''"></td>
      <td bgcolor="#F8F8F8" width="15%"><input type="submit" name="forumsubmit" value="�� ��"></td>
    </tr>
  </table>
</form>
<form method="Post" action="?Action=ForumAddok&add=Forum_0">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="3">�����һ������</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="15%">����:</td>
      <td bgcolor="#FFFFFF" width="70%"><input type="text" name="newforum" value="�·�������" size="40"></td>
      <td bgcolor="#F8F8F8" width="15%"><input type="submit" name="catsubmit" value="�� ��"></td>
    </tr>
  </table>
</form>
<form method="post" action="?Action=ForumSort">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td>�༭��̳</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="100%" valign=top height=200>
	  <table cellspacing="1" cellpadding="1" width="98%" align="center" class="a2">
	  <tr class="a4"><td>
			<% ForumList(0) %>
			</td></tr>
        </table></td>
    </tr>
  </table><br><center><input type="submit" name="detailsubmit" value="�� ��"></form><br>
<%
End Sub

Sub Manages
		Dim Board_Setting,RS
		Dim B_Lookperm,B_Postperm,B_DownPerm,B_Upperm
		If ID="" or (Not isNumeric(ID)) Then SuccessMsg " ID����ֻ��������! "
		Set Rs=team.Execute("Select bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Readme,Board_Key,Board_URL,Lookperm,Postperm,DownPerm,Upperm From "&IsForum&"Bbsconfig Where ID="&ID)
		If RS.Eof or Rs.Bof Then
			SuccessMsg("ID��������!")
		Else
			Board_Setting = Split(RS("Board_Setting"),"$$$")
		%>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?Action=Forumeditok">
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="detailsubmit" value="submit">
	<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td>������ʾ</td>
    </tr>
    <tr bgcolor="#F8F8F8">
      <td><br>
        <ul>
          ��������û�м̳��ԣ������Ե�ǰ��̳��Ч��������¼�����̳����Ӱ�졣
        </ul></td>
    </tr>
  </table>
  <br>
  <br>
  <a name="��̳��ϸ���� - <%=RS(0)%>"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">��̳��ϸ���� - <%=RS(0)%></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��ʾ��̳:</b><br>
        <span class="a3">ѡ�񡰷񡱽���ʱ����̳���ز���ʾ������̳�����Խ����������û��Կ�ͨ��ֱ���ṩ���� id �� URL ���ʵ�����̳��������ص���һ����飬��ô�����ڵ��¼���齫���������һ�����ء�</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Hide" value="0" <%If RS("Hide")=0 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Hide" value="1" <%If RS("Hide")=1 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�ϼ���̳:</b><br>
        <span class="a3">����̳���ϼ���̳�����</span></td>
      <td bgcolor="#FFFFFF">
	  <select name="fupnew">
			<option value="0">&nbsp;>>һ����̳</option>
			<% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��񷽰�:</b><br>
        <span class="a3">�����߽��뱾��̳��ʹ�õķ�񷽰�</span></td>
      <td bgcolor="#FFFFFF">
	  <select name="Board_Setting(0)">
		<option value="<%=Int(team.Forum_setting(18))%>" SELECTED>������̳Ĭ��ģ��</option>
      <%
		Dim RS1,SytyleID
		Set Rs1=team.Execute( "Select StyleName,ID From ["&IsForum&"Style] Order By ID Asc" )
		Do While Not RS1.Eof
			SytyleID = SytyleID &  "<option value="&RS1(1)&"" 
			If Int(Rs1(1)) = Int(Board_Setting(0)) Then SytyleID = SytyleID & " SELECTED"
			SytyleID = SytyleID &">"&RS1(0)&"</option>"
			Rs1.Movenext
		Loop
		RS1.CLOSE:Set RS1=Nothing
		Response.Write SytyleID
	'����     ����         ���� ����  ͼ��  Ȩ��     ����   ����     ת���ַ
	'bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Readme,Board_Key,Board_URL
%>
        </select></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��̳ת�� URL</b><br>
        <span class="a3">�������ת�� URL(���� http://www.team5.cn)���û����������̳������ת�������õ� URL��һ���趨���޷�������̳ҳ�棬��ȷ���Ƿ���Ҫʹ�ô˹��ܣ�����Ϊ������ת�� URL����վ�����URL��ַ�����HTTP://</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Board_URL" value="<%=RS("Board_URL")%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��̳����:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="BbsName" value="<%=RS("BbsName")%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��̳ͼ��:</b><br>
        <span class="a3">��̳���ƺͼ������Сͼ�꣬����д��Ի���Ե�ַ</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Icon" value="<%=RS("Icon")%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>��̳���:</b><br>
        <span class="a3">����ʾ����̳���Ƶ����棬�ṩ�Ա���̳�ļ��������֧��Ubb���� </span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="Readme" cols="30" style="height:70;overflow-y:visible;"><%=ReplaceStr(RS("Readme"),"<BR>",VbCrlf)%></textarea></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>����̳����:</b><br>
        <span class="a3">��ʾ�������б�ҳ�ĵ�ǰ��̳����֧�� Ubb ���룬����Ϊ����ʾ</span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="Board_Key" cols="30" style="height:70;overflow-y:visible;"><%=ReplaceStr(RS("Board_Key"),"<BR>",VbCrlf)%></textarea></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��������޸ı���̳����:</b><br>
        <span class="a3">�����Ƿ������������Ͱ���ͨ��ϵͳ�����޸ı������</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(1)" value="0" <%If Board_Setting(1)=0 Then%>checked<%End If%>>
        ����������޸� 
		<input type="radio" name="Board_Setting(1)" value="1" <%If Board_Setting(1)=1 Then%>checked<%End If%>>��������޸� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>����Ȩ������:</b><br>
        <span class="a3">������ͨ����ѡ������ <B>�����̳���</B>�� ���Ƹ��û���鿴�������ݼ����ӱ����б��Ȩ�ޡ����������ø��û����ڲ鿴�����ӱ��������£����޷��鿴�����������ϸ���ݡ�</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(8)" value="0" <%If Board_Setting(8)=0 Then%>checked<%End If%>>
        ȫ������
		<input type="radio" name="Board_Setting(8)" value="1" <%If Board_Setting(8)=1 Then%>checked<%End If%>>ֻ�ڲ鿴��������ʱ��������
		</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top">
	  <B>�����οͷ���Ȩ��:</B> 
	  <BR>��������οͷ���������Ҫ���� <BR>
	     1. ��Ȩ������ ��<a href="Admin_Group.asp?Action=Editmodel_1&ID=28#�������">��������ο���Ȩ��</a>�� <BR>
		 2.  <B><a href="#��̳Ȩ��">���»������</a></B> �����οͷ���Ȩ�� <br>
        <span class="a3">  </span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(9)" value="1" <%If CID(Board_Setting(9))=1 Then%>checked<%End If%>>
        ��
		<input type="radio" name="Board_Setting(9)" value="0" <%If CID(Board_Setting(9))=0 Then%>checked<%End If%>>��
		</td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="�� ��"><br>
  <br>
  <a name="����ѡ��"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">����ѡ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�������:</b><br>
        <span class="a3">ѡ���ǡ���ʹ�û��ڱ��淢������Ӵ����������Ա���ͨ�������ʾ�������򿪴˹��ܺ����������û������趨��Щ�鷢���ɲ�����ˣ�Ҳ�����ڹ��������趨��Щ�������˱��˵�����</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(2)" value="0" <%If Board_Setting(2)=0 Then%>checked<%End If%>>
        ��<br>
        <input type="radio" name="Board_Setting(2)" value="1" <%If Board_Setting(2)=1 Then%>checked<%End If%>>
        ���������<br>
        <input type="radio" name="Board_Setting(2)" value="2" <%If Board_Setting(2)=2 Then%>checked<%End If%>>
        �����������»ظ� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>���ƻظ�����:</b><br>
        <span class="a3">ѡ���ǡ���ʹ�û��ڱ����޷�����ظ����⡣</span></td>
      <td bgcolor="#FFFFFF">
	    <input type="radio" name="Board_Setting(5)" value="1" <%If CID(Board_Setting(5))=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(5)" value="0" <%If CID(Board_Setting(5))=0 Then%>checked<%End If%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��������ļ�:</b><br>
        <span class="a3">�Ƿ������û����ڱ��淢�������������Լ��� �ļ� �С�ע��: һ�����ⱻ���� �ļ��������ݻᱻ���������۵�ǰ��̳�����ʲô����Ȩ���趨</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(4)" value="1" <%If Board_Setting(4)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(4)" value="0" <%If Board_Setting(4)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ʹ�� UBB ����:</b><br>
        <span class="a3">UBB ������һ�ּ򻯺Ͱ�ȫ��ҳ���ʽ���룬�� <a href="../Help.asp?page=mise#1" target="_blank">�������鿴����̳�ṩ��UBB����</a></span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(6)" value="1" <%If Board_Setting(6)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(6)" value="0" <%If Board_Setting(6)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�������ݸ�����:</b><br>
        <span class="a3">ѡ���ǡ�����������������������ĸ����ִ���ʹ�÷������޷�����ԭʼ���ݡ�ע��: �����ܻ���΢���ط���������</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(7)" value="1" <%If Board_Setting(7)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(7)" value="0" <%If Board_Setting(7)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="�� ��"><br>
  <br>
  <a name="��������"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">��������</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�Զ��巢�������ӻ���:</b><br>
        <span class="a3">���ñ���̳�Ƿ�ʹ�ö����ķ���������ֹ���ѡ���ǡ����û��ڱ���̳������ʱ�����ֽ�����������������������������������������������ֵ��ѡ�񡰷񡱣����ֽ���ȫ��̳Ĭ���趨�Ĺ�������</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(12)" value="1" <%If Board_Setting(12)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(12)" value="0" <%If Board_Setting(12)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�Զ��巢�ظ����ӻ���:</b><br>
        <span class="a3">���ñ���̳�Ƿ�ʹ�ö����ķ��»ظ����ֹ���ѡ���ǡ����û��ڱ���̳���ظ�ʱ�����ֽ�����������������������������������������������ֵ��ѡ�񡰷񡱣����ֽ���ȫ��̳Ĭ���趨�Ĺ�������</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(13)" value="1" <%If Board_Setting(13)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(13)" value="0" <%If Board_Setting(13)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�Զ��徫�������ӻ���:</b><br>
        <span class="a3">���ñ���̳�Ƿ�ʹ�ö����ķ��»ظ����ֹ���ѡ���ǡ����û����ӱ����뾫��ʱ�����ֽ�����������������������������������������������ֵ��ѡ�񡰷񡱣����ֽ���ȫ��̳Ĭ���趨�Ĺ�������</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(3)" value="1" <%If Board_Setting(3)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(3)" value="0" <%If Board_Setting(3)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td colspan="2" bgcolor="#F8F8F8"><table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
          <tr align="center" class="a1">
            <td>���ִ���</td>
            <td>��������</td>
            <td>������(+)</td>
            <td>�ظ�(+)</td>
			<td>�Ӿ���(+)</td>
          </tr>
          <% 
	Dim ExtCredits,U,ExtSort,Board_Up,My_ExtSort,MY_ExtCredits,Mydisabled
	ExtCredits= Split(team.Club_Class(21),"|")
	MY_ExtCredits=Split(Board_Setting(14),"|")
	For U=0 to Ubound(ExtCredits)
		ExtSort=Split(ExtCredits(U),",")
		My_ExtSort=Split(MY_ExtCredits(U),",")
		If ExtSort(3)=1 then
			Mydisabled =""
		Else
			Mydisabled = "disabled"
		End If
	%>
          <tr align="center" <%=Mydisabled%>>
            <td bgcolor="#F8F8F8">ExtCredits<%=U+1%></td>
            <td bgcolor="#FFFFFF"><%=ExtSort(0)%></td>
            <td bgcolor="#F8F8F8"><input type="text" size="2" name="ExtCredits<%=U%>_0" value="<%=My_ExtSort(0)%>"></td>
            <td bgcolor="#FFFFFF"><input type="text" size="2" name="ExtCredits<%=U%>_1" value="<%=My_ExtSort(1)%>"></td>
			<td bgcolor="#FFFFFF"><input type="text" size="2" name="ExtCredits<%=U%>_2" value="<%=My_ExtSort(2)%>"></td>
          </tr>
          <%
	Next
	%>
        </table></td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="�� ��"><br>
  <br>
  <a name="�������"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�������</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����������:</b><br>
        <span class="a3">�����Ƿ��ڱ���̳����������๦�ܣ�����Ҫͬʱ�趨��Ӧ�ķ���ѡ��������ñ�����</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(15)" value="1" <%If Board_Setting(15)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(15)" value="0" <%If Board_Setting(15)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����������:</b><br>
        <span class="a3">���ѡ���ǡ������߷�������ʱ������ѡ�������Ӧ�������ܷ��������ܱ��롰����������ࡱ��ſ�ʹ�� </span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(16)" value="1" <%If Board_Setting(16)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(16)" value="0" <%If Board_Setting(16)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����������:</b><br>
        <span class="a3">���ѡ���ǡ����û��������ڱ���̳�а��ղ�ͬ�����������⡣ע��: �����ܱ��롰����������ࡱ��ſ�ʹ�ò�����ط���������</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(17)" value="1" <%If Board_Setting(17)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(17)" value="0" <%If Board_Setting(17)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>���ǰ׺:</b><br>
        <span class="a3">�����Ƿ��������б��У����ѷ��������ǰ����������ʾ��ע��: �����ܱ��롰����������ࡱ��ſ�ʹ��</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(18)" value="1" <%If Board_Setting(18)=1 Then%>checked<%End If%>>
        ��
        <input type="radio" name="Board_Setting(18)" value="0" <%If Board_Setting(18)=0 Then%>checked<%End If%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>���ѡ��:</b><br>
        <span class="a3">����д�����ڱ���̳��ʹ�õ����ѡ��û�������������ʱ�����ɰ���ѡ�е�������������ÿ�����һ�� ��ע��: �����ܱ��롰����������ࡱ��ſ�ʹ�á�</span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="Board_Setting(19)" cols="30" style="height:70;overflow-y:visible;"><%=Board_Setting(19)%></textarea></td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="�� ��"><br>
  <br>
  <a name="��̳Ȩ��"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">��̳Ȩ�� - ȫ��ѡ����Ĭ������</td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>��������:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Pass" value="<%=RS("Pass")%>"></td>
    </tr>
    <tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>�����̳���:</b><br>
        <span>Ĭ��Ϊȫ�����������̳����Ȩ�޵��û���<br>
        <input type="checkbox" name="chkall1" onClick="checkall(this.form, 'lookperm', 'chkall1')">
        ȫѡ</span></td>
      <td bgcolor="#FFFFFF">
	  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
	  <tr>
		<% 
		Dim Gs,Value,i,m
		Set Gs = team.execute("Select ID,MemberRank,GroupName From "&IsForum&"UserGroup Where Not (ID=7 or ID=6 or ID=5) Order By GroupRank Desc")
		If Gs.Eof or Gs.Bof Then
			SuccessMsg " �û�Ȩ�ޱ�������,���ֶ������±�! "
		Else
			Value = Gs.GetRows(-1)
		End If
		Gs.Close:Set Gs=Nothing

		'bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Readme,Board_Key,Board_URL,Lookperm,Postperm,DownPerm,Upperm
		If Instr(RS(9),",")>0 Then B_Lookperm = Split(RS(9),",")
		If Instr(RS(10),",")>0 Then B_Postperm = Split(RS(10),",")
		If Instr(RS(11),",")>0 Then B_DownPerm = Split(RS(11),",")
		If Instr(RS(12),",")>0 Then B_Upperm = Split(RS(12),",")
		If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				response.write "<td><input type=""checkbox"" name=""lookperm"" class=""radio"" value="&Replace(Value(0,i)," ","")&" "
				If Isarray(B_Lookperm) Then
					for m = 0 to Ubound(B_Lookperm)-1
						If Cid(Trim(B_Lookperm(m))) = int(Value(0,i)) Then response.write "checked"
					next
				end if
				response.write "  >"&Value(2,i)&"</td> "
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table>
		</td>
    </tr><tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>���»������:</b><br>
        <span>Ĭ��Ϊ���ο���������з���Ȩ�޵��û���<br>
        <input type="checkbox" name="chkall2" onClick="checkall(this.form, 'postperm', 'chkall2')">
        ȫѡ</span></td>
      <td bgcolor="#FFFFFF">
	  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr><%
        If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				Echo "<td><input type=""checkbox"" name=""postperm"" class=""radio"" value="&Value(0,i)&" "
				If Isarray(B_Postperm) Then
					for m = 0 to Ubound(B_Postperm)-1
						If Cid(Trim(B_Postperm(m))) = cid(Value(0,i)) Then Echo "checked"
					next
				end if
				Echo "  >"&Value(2,i)&"</td> "
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table></td>
    </tr>
	<tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>����/�鿴�������:</b><br>
        <span>Ĭ��Ϊȫ����������/�鿴����Ȩ�޵��û���<br>
        <input type="checkbox" name="chkall4" onClick="checkall(this.form, 'downperm', 'chkall4')">
        ȫѡ</span></td>
      <td bgcolor="#FFFFFF">
		  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr><%
        If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				Echo "<td><input type=""checkbox"" name=""downperm"" class=""radio"" value="&Value(0,i)&" "
				If Isarray(B_DownPerm) Then
					for m = 0 to Ubound(B_DownPerm)-1
						If Cid(Trim(B_DownPerm(m))) = Cid(Value(0,i)) Then Echo "checked"
					next
				end if
				Echo "  >"&Value(2,i)&"</td> "				
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table></td>
    </tr>
	<tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>�ϴ��������:</b><br>
        <span>Ĭ��Ϊ���ο���������ϴ�����Ȩ�޵��û���<br>
        <input type="checkbox" name="chkall5" onClick="checkall(this.form, 'upperm', 'chkall5')">
        ȫѡ</span></td>
      <td bgcolor="#FFFFFF">
		  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr><%
        If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				Echo "<td><input type=""checkbox"" name=""upperm"" class=""radio"" value="&Value(0,i)&" "
				If Isarray(B_Upperm) Then
					for m = 0 to Ubound(B_Upperm)
						If Cid(Trim(B_Upperm(m))) = CID(Value(0,i)) Then Echo "checked"
					next
				end if
				Echo "  >"&Value(2,i)&"</td> "	
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table></td>
    </tr>
	<tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
  </table>
  <br>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="�� ��">
</form>
<br>
<%	
	End If
	Rs.Close:Set Rs=Nothing
End Sub


Sub Forumadd%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr bgcolor="#F8F8F8">
    <td><br>
      <ul>
        ��ֻ������˰���Ժ�ſ��Զ԰�������ϸ�����á�
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?Action=ForumAddok&add=Forum_1">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan=5>�������̳</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="15%">����:</td>
      <td bgcolor="#FFFFFF" width="28%"><input type="text" name="newforum" value="����̳����" size="20"></td>
      <td bgcolor="#F8F8F8" width="15%">�ϼ���̳:</td>
      <td bgcolor="#FFFFFF" width="27%"><select name="fup">
		  <option value="0">&nbsp;>>һ����̳</option>
          <% ForumList_Sel(0) %>
        </select></td>
      <td bgcolor="#F8F8F8" width="15%"><input type="submit" name="forumsubmit" value="�� ��"></td>
    </tr>
  </table>
</form>
<br>
<%
End Sub
Sub SetModerators
	Dim Rs3,ho
	Dim newmoderator,newdisplayorder,Rs1
	If Request("UpModers")=1 Then
		Newmoderator = HTMLEncode(Request.Form("newmoderator"))
		Newdisplayorder = HTMLEncode(Request.Form("newdisplayorder"))
		for each ho in request.form("isdelete")
			Team.execute("Delete from ["&isforum&"Moderators] Where id="&ho)
		next
		If Request.form("isdelete")="" Then
			If Newmoderator="" or Newdisplayorder="" Then Error2("��������Ϊ��!")
			If team.execute("Select * from ["&isforum&"User] where UserName='"&Newmoderator&"' ").Eof Then
				Error2("ָ���û������ڣ��뷵�ء�")
			Else
				Set Rs3 = team.execute("Select UserGroupID from ["&isforum&"User] where UserName='"&Newmoderator&"' ")
				If Not RS3.Eof Then
					If Not (CID(Rs3(0)) = 1 Or CID(Rs3(0)) = 2) Then
						If team.execute("Select ManageUser from "&isforum&"Moderators where ManageUser='"&Newmoderator&"' and BoardID="&ID).Eof Then
							team.execute("insert into "&isforum&"Moderators (BoardID,ManageUser,Issort) values ("&ID&",'"&Newmoderator&"',"&Newdisplayorder&") ")
							team.execute("Update ["&isforum&"User] Set UserGroupID=3,Members='����',Levelname='����||||||16||0' where UserName='"&Newmoderator&"' ")
						Else
							error2 " �˰����Ѿ�����! "
						End If
					End If 
				End If
				RS3.Close:Set Rs3=nothing
			End If
		End If
		Cache.DelCache("ManageUsers")
		SuccessMsg("�������óɹ�!")	
	Else%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?Action=SetModerators&UpModers=1&ID=<%=request("ID")%>">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan=4>TEAM's - �༭���� - <%=team.execute("Select bbsname from "&isforum&"Bbsconfig where id="&id )(0)%></td>
    </tr>
    <tr align="center" class=a3>
      <td> <input type="checkbox" name="chkall" class="a4" onClick="checkall(this.form)">ɾ?</td>
      <td>�û���</td>
      <td>��ʾ˳��</td>
    </tr>
    <%  
		Dim Rs
		Set Rs=team.Execute("Select id,ManageUser,Issort from "&isforum&"Moderators Where BoardID="& ID)
		Do While Not Rs.Eof
		%>
    <tr align="center" class="a4">
      <td><input type="checkbox" name="isdelete" value="<%=rs(0)%>"></td>
      <td><%=rs(1)%></td>
      <td><%=rs(2)%></td>
    </tr>
    <% Rs.MoveNext
		Loop
		Rs.Close:Set Rs=Nothing
		%>
    <tr align="center" class="a3">
      <td>����:</td>
      <td><input type='text' name="newmoderator" size="20"></td>
      <td><input type="text" name="newdisplayorder" size="2" value="0"></td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="forumsubmit" value=" �� �� ">
  &nbsp;
</form>
</center>
<br>
<%
	End If
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
Sub UniForum()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<br>
<br>
<br>
<br>
<form method="post" action="?Action=Forumsmerge">
  <table cellspacing="1" cellpadding="4" width="85%" align="center" class="a2">
    <tr class="a1">
      <td colspan="3">�ϲ���̳ - Դ��̳������ȫ��ת��Ŀ����̳��ͬʱɾ��Դ��̳</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">Դ��̳:</td>
      <td bgcolor="#FFFFFF" width="60%"><select name="source">
          <option value="">�� ��ѡ��</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">Ŀ����̳:</td>
      <td bgcolor="#FFFFFF" width="60%"><select name="target">
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
		If Request("Action") = "Forumadd" Then
			If RS(0) = int(Request("ID")) Then Response.Write " selected "
		End If
		If Request("Action") = "Manages" Then
			If RS(0) = int(Request("RootID")) Then Response.Write " selected "
		End If
		Response.Write ">"&String(ii,"&nbsp;")&""&W&""&RS(1)&"</option>"
		ii=ii+1
		ForumList_Sel RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub

Dim ManageUsers,Moderuser
Sub ForumList(V)
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Hide,BbsName,SortNum,Board_Model From "&IsForum&"Bbsconfig Where Followid=0 Order By SortNum")
	Do While Not RS.Eof
		ManageUsers = team.GroupManages()
		Moderuser = ""
		If isarray(ManageUsers) Then
			for	u=0 to Ubound(ManageUsers,2)
				If ManageUsers(2,u) = Rs(0) Then
					Moderuser = Moderuser & ManageUsers(1,u) & " "
				End If
			Next
		End If
		Select Case RS(1)
			Case 1
				T="����"
			Case Else
				T="����"
		End Select
		If RS(4)=1 then
			sty = "<a href=?Action=ModelSet_1&ID="&RS(0)&" title=""���ת����ʾģʽ"">���ģʽ</a>"
		Else
			sty = "<a href=?Action=ModelSet_0&id="&rs(0)&" title=""���ת����ʾģʽ"">����ģʽ</a>"
		End If
		Echo "<ul><li><a target=_blank href=../Default.asp?rootid="&RS(0)&"><b>"&RS(2)&"</b></a> - <span class=a4></a> - ��ʾ˳��: <input type=text name=SortNum Value="&RS(3)&" Size=""1""><Input Name=UID value="&RS(0)&" type=hidden> - <a href=""?Action=Forumadd&ID="&RS(0)&""" title=""��ӱ��������̳���¼���̳"">[���]</a> <a href=""?Action=Manages&ID="&RS(0)&""" title=""�༭����̳����"">[�༭]</a> <a href=""?Action=DelForum&ID="&RS(0)&""" title=""ɾ������̳��������������"">[ɾ��]</a> - [״̬: <b>"&T&"</b>]</a> - [��ʾģʽ: "&sty&" ] - [<a href=""?Action=SetModerators&ID="&RS(0)&""" title=""�༭����̳����"">���� "&Moderuser&"</a>]</span>"
		Call ForumList_1(Rs(0))
		Echo " </li></ul> "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub

Sub ForumList_1(a)
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Hide,BbsName,SortNum,Board_Model,Followid From "&IsForum&"Bbsconfig Where Followid="&a&" Order By SortNum")
	Do While Not RS.Eof
		ManageUsers = team.GroupManages()
		Moderuser = ""
		If isarray(ManageUsers) Then
			for	u=0 to Ubound(ManageUsers,2)
				If ManageUsers(2,u) = Rs(0) Then
					Moderuser = Moderuser & ManageUsers(1,u) & " "
				End If
			Next
		End If
		Select Case RS(1)
			Case 1
				T="����"
			Case Else
				T="����"
		End Select
		Echo "<ul><li>"&String(ii*2,"��")& S &"<a target=_blank href=../Forums.asp?fid="&RS(0)&"><b>"&RS(2)&"</b></a> - <span class=a4></a> - ��ʾ˳��: <input type=text name=SortNum Value="&RS(3)&" Size=""1""><Input Name=UID value="&RS(0)&" type=hidden> - <a href=""?Action=Forumadd&ID="&RS(0)&""" title=""��ӱ��������̳���¼���̳"">[���]</a> <a href=""?Action=Manages&ID="&RS(0)&"&RootID="&RS(5)&""" title=""�༭����̳����"">[�༭]</a> <a href=""?Action=DelForum&ID="&RS(0)&""" title=""ɾ������̳��������������"">[ɾ��]</a> - [״̬: <b>"&T&"</b>]</a> - [<a href=""?Action=SetModerators&ID="&RS(0)&""" title=""�༭����̳����"">���� "&Moderuser&"</a>]</span>"
		Call ForumList_1(Rs(0))
		Echo " </li></ul> "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing

End Sub

%>
