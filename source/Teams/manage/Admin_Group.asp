<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim ii,ID
Dim Admin_Class
Call Master_Us()
Header()
ii=0:
ID=CID(Request("ID"))
Admin_Class=",5,"
Call Master_Se()
team.SaveLog ("�����뼶�� [������������ ���û��� ] ")
Select Case Request("Action")
	Case "Editmodel_1"
		Call Editmodel_1
	Case "Editmodel_2"
		Call Editmodel_2
	Case "IsuserGroup"
		Call IsuserGroup
	Case "EditUserGroup"
		Call EditUserGroup
	Case "EditUserManages"
		Call EditUserManages
	Case "ManagesMember"
		Call ManagesMember
	Case Else
		Call Main()
End Select

Sub Main()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
  <tr class="a1">
    <td>TEAM's��ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>TEAM's! �������������Ա�����������������Լ������˹���Ȩ�޵������飬������Ա�����⣬���������������ϸ���ù���Ȩ�ޡ�
      </ul>
      <ul>
        <li>�����Զ�������鷽��Ϊ��
          <ul>
            <li>����<a href="Admin_Group.asp?Action=IsuserGroup"><b>�û�������</b></a>������һ���µ������飻
            <li>�༭�������飬�������������ĳ�ֹ���Ȩ��(����Ա���������������)��ͬʱ�༭�����������Ŀ���ã�
            <li>����<a href="Admin_Group.asp"><b>����������</b></a>���༭����Ĺ���Ȩ�ޡ�
          </ul>
      </ul>
      <ul>
        <li>ɾ���Զ��������ķ��������¶��֣�
          <ul>
            <li>�༭����������ã�ȡ������Ȩ�޹�����
            <li>����<a href="Admin_Group.asp?Action=IsuserGroup"><b>�û�������</b></a>���༭��ȡ������Ȩ�޹�������ֱ��ɾ���������顣
          </ul>
      </ul></td>
  </tr>
</table>
<br />
<table cellspacing="1" cellpadding="4" width="90%" align="center" class=a2>
<tr class="a1" align="center">
  <td>����</td>
  <td>����</td>
  <td>������</td>
  <td>��������</td>
  <td>����Ȩ��</td>
</tr>
<%
Dim Rs
Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,UserColor,UserImg,Rank from "&IsForum&"UserGroup Where GroupTips = 3 and MemberRank = -1 or (GroupRank = 1 or GroupRank = 2 or GroupRank = 3)")
Do While Not Rs.Eof
%>
<tr align="center">
  <td class="a3"><%=Rs(2)%></td>
  <td class="a4">����</td>
  <td class="a3"><%=Rs(1)%></td>
  <td class="a4"><a href="Admin_Group.asp?Action=Editmodel_1&ID=<%=Rs(0)%>">[�༭]</a></td>
  <td class="a3"><a href="Admin_Group.asp?Action=Editmodel_2&ID=<%=Rs(0)%>">[�༭]</a></td>
</tr>
<%
	Rs.MoveNext
Loop
Rs.Close:Set Rs=Nothing
Response.Write "</table>"
End Sub


Sub Editmodel_1 
	Dim Rs,Group_Set_Class
	Set Rs=team.Execute("Select ID,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,UserColor,UserImg,Rank from "&IsForum&"UserGroup Where ID="&ID)
	If Rs.Eof Then 
		SuccessMsg "��������! "
	Else
		Group_Set_Class = Split(Rs(4),"|")
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<form method="post" action="?Action=EditUserGroup">
 <input type="hidden" name="myid" value="<%=ID%>">
  <a name="�༭�û���"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�༭�û���</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�û���ͷ��:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="GroupName" value="<%=Replace(RS(1),"&nbsp;","")%>"></td>
    </tr>
  </table>
  <br>
  <%If Request("OnGroup")="yes" Then%>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�û����������</td>
    </tr>
    <tr class="a4">
      <td><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr class="a4">
		  <input type="hidden" name="oldgroup" value="<%=Rs(3)%>">
            <%
			Dim Gs,u
			Set Gs = team.execute("Select Max(ID) as ID,Members From ["&IsForum&"UserGroup] Where GroupTips=3 or GroupTips=2 or GroupTips=1 Group By Members")
			If Gs.Eof or Gs.Bof Then
				SuccessMsg " �û�Ȩ�ޱ�������,���ֶ������±�! "
			Else
				u=0
				Do While Not Gs.Eof 
					u = u+1
					Echo "<td> "
					Response.write "<input type=""radio"" name=""mygroups"" value="""&Gs(0)&""" "
					If int(Rs(3)) =  int(Gs(0)) Then Response.write " checked "
					Response.write "> "&Gs(1)&" </td>"
					If U= 5 Then 
						Echo "</tr><tr class=""a4"">"
						U=0
					End If
					Gs.MoveNext
				Loop
			End If
			Gs.Close:Set Gs=Nothing%>
        </table></td>
    </tr>
  </table>
  <%End if%>
  <br>
  <a name="����Ȩ��"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">����Ȩ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>���������̳:</b><br>
        <span class="a4">ѡ�񡰷񡱽����׽�ֹ�û�������̳���κ�ҳ��</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(0)" value="1" <%if CID(Group_Set_Class(0))=1 Then%>checked<%end if%>>
		�� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(0)" value="0" <%if CID(Group_Set_Class(0))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�Ķ�Ȩ��:</b><br>
        <span class="a4">�����û�������ӻ򸽼���Ȩ�޼��𣬷�Χ 0��255��0 Ϊ��ֹ�û�����κ����ӻ򸽼������û����Ķ�Ȩ��С�����ӵ��Ķ�Ȩ�����(Ĭ��ʱΪ 1)ʱ���û��������Ķ������ӻ����ظø���</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(1)" value="<%=CID(Group_Set_Class(1))%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����鿴�û�����:</b><br>
        <span class="a4">�����Ƿ�����鿴�����û���������Ϣ</span></td>
      <td bgcolor="#FFFFFF">
	  <input type="radio" name="Group_Set_Class(2)" value="1" <%if CID(Group_Set_Class(2))=1 Then%>checked<%end if%>>�� &nbsp; &nbsp;
      <input type="radio" name="Group_Set_Class(2)" value="0" <%if CID(Group_Set_Class(2))=0 Then%>checked<%end if%>>��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�������ת��:</b><br>
        <span class="a4">�����Ƿ������û��������н��Լ��Ľ��׻���ת�ø������û���ע��: ����������ѡ�������ý��׻��ּ����˿��������ʺź�ſ�ʹ��</span></td>
      <td bgcolor="#FFFFFF">
	    <input type="radio" name="Group_Set_Class(3)" value="1" <%if CID(Group_Set_Class(3))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(3)" value="0" <%if CID(Group_Set_Class(3))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ʹ������:</b><br>
        <span class="a4">�����Ƿ�����ͨ�����ݿ�������������Ͷ���Ϣ������ע��: ����������ʱ��ȫ���������ǳ��ķѷ�������Դ��������</span></td>
      <td bgcolor="#FFFFFF">
	    <input type="radio" name="Group_Set_Class(4)" value="0" <%if CID(Group_Set_Class(4))=0 Then%>checked<%end if%>>
        ��������<br>
        <input type="radio" name="Group_Set_Class(4)" value="1" <%if CID(Group_Set_Class(4))=1 Then%>checked<%end if%>>
        ֻ������������<br>
        <input type="radio" name="Group_Set_Class(4)" value="2" <%if CID(Group_Set_Class(4))=2 Then%>checked<%end if%>>
        ����ȫ������</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ʹ��ͷ��:</b><br>
        <span class="a4">�����Ƿ�����ʹ���Զ���ͷ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(5)" value="0" <%if CID(Group_Set_Class(5))=0 Then%>checked<%end if%>>
        ��ֹʹ��ͷ��<br>
        <input type="radio" name="Group_Set_Class(5)" value="1" <%if CID(Group_Set_Class(5))=1 Then%>checked<%end if%>>
        ����ʹ����̳ͷ��<br>
        <input type="radio" name="Group_Set_Class(5)" value="2" <%if CID(Group_Set_Class(5))=2 Then%>checked<%end if%>>
        ����ʹ����̳ͷ����ϴ�ͷ��<br>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������û�����:</b><br>
        <span class="a4">�����Ƿ�����������û������ӽ������ֲ���</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(10)" value="1" <%if CID(Group_Set_Class(10))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(10)" value="0" <%if CID(Group_Set_Class(10))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�������ɶ���Ϣ֪ͨ����:</b><br>
        <span class="a4">�����û��ڶ��������ֻ�������ʱ�Ƿ�ǿ���������ɺ�֪ͨ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(6)" value="0" <%if CID(Group_Set_Class(6))=0 Then%>checked<%end if%>>
        ��ǿ��<br>
        <input type="radio" name="Group_Set_Class(6)" value="1" <%if CID(Group_Set_Class(6))=1 Then%>checked<%end if%>>
        ǿ����������<br>
        <input type="radio" name="Group_Set_Class(6)" value="2" <%if CID(Group_Set_Class(6))=2 Then%>checked<%end if%>>
        ǿ��֪ͨ����<br>
        <input type="radio" name="Group_Set_Class(6)" value="3" <%if CID(Group_Set_Class(6))=3 Then%>checked<%end if%>>
        ǿ���������ɺ�֪ͨ����</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ʹ�� �ļ�:</b><br>
        <span class="a4">�����Ƿ���������¼�����˵� �ļ� �У��Ӷ����������</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(7)" value="1" <%if CID(Group_Set_Class(7))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(7)" value="0" <%if CID(Group_Set_Class(7))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������ͶƱ:</b><br>
        <span class="a4">�����Ƿ������û�����ͶƱ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(8)" value="1" <%if CID(Group_Set_Class(8))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(8)" value="0" <%if CID(Group_Set_Class(8))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������:</b><br>
        <span class="a4">�����Ƿ������û�������֯�����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(9)" value="1" <%if CID(Group_Set_Class(9))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(9)" value="0" <%if CID(Group_Set_Class(9))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������������:</b><br>
        <span class="a4">�����Ƿ������û������������������</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(20)" value="1" <%if CID(Group_Set_Class(20))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(20)" value="0" <%if CID(Group_Set_Class(20))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����Զ���ͷ��:</b><br>
        <span class="a4">�����Ƿ������û������Լ���ͷ�����ֲ�����������ʾ</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(11)" value="1" <%if CID(Group_Set_Class(11))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(11)" value="0" <%if CID(Group_Set_Class(11))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����Ϣ�ռ�������:</b><br>
        <span class="a4">�����û�����Ϣ���ɱ������Ϣ��Ŀ��0 Ϊ��ֹʹ�ö���Ϣ</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(12)" value="<%=Group_Set_Class(12)%>"></td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="�� ��">
  </center>
  <br>
  <a name="�������"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�������</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����»���:</b><br>
        <span class="a4">�����Ƿ������»��⡣ע��: ֻ�е��û����Ķ�Ȩ�޸��� 0 ʱ�����ܷ��»���</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(13)" value="1" <%if CID(Group_Set_Class(13))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(13)" value="0" <%if CID(Group_Set_Class(13))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������ظ�:</b><br>
        <span class="a4">�����Ƿ�������ظ���ע��: ֻ�е��û����Ķ�Ȩ�޸��� 0 ʱ�����ܷ���ظ�</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(14)" value="1" <%if CID(Group_Set_Class(14))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(14)" value="0" <%if CID(Group_Set_Class(14))=0 Then%>checked<%end if%>>
        ��</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�������ͶƱ:</b><br>
        <span class="a4">�����Ƿ����������̳��ͶƱ</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(15)" value="1" <%if CID(Group_Set_Class(15))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(15)" value="0" <%if CID(Group_Set_Class(15))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
    <td width="60%" bgcolor="#F8F8F8" ><b>����ֱ�ӷ���:</b><br>
        <span class="a4">��ѡ��ֻ����̳������Ϊ��Ҫ�������ʱ�������ã�����ѡ���ǡ�����������û�������˶�ֱ�ӷ���������</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(16)" value="0" <%if CID(Group_Set_Class(16))=0 Then%>checked<%end if%>>
        ȫ����Ҫ���<br>
        <input type="radio" name="Group_Set_Class(16)" value="1" <%if CID(Group_Set_Class(16))=1 Then%>checked<%end if%>>
        ���»ظ�����Ҫ���<br>
        <input type="radio" name="Group_Set_Class(16)" value="2" <%if CID(Group_Set_Class(16))=2 Then%>checked<%end if%>>
        �������ⲻ��Ҫ���<br>
        <input type="radio" name="Group_Set_Class(16)" value="3" <%if CID(Group_Set_Class(16))=3 Then%>checked<%end if%>>
        ȫ������Ҫ���</td>
    </tr><tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����������:</b><br>
        <span class="a4">�Ƿ������û�������������ͻظ���ֻҪ�û������̳�����û�����ʹ�������������ܡ�����������ͬ���οͷ������û���Ҫ��¼��ſ�ʹ�ã������͹���Ա���Բ鿴��ʵ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(17)" value="1" <%if CID(Group_Set_Class(17))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(17)" value="0" <%if CID(Group_Set_Class(17))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������������Ȩ��:</b><br>
        <span class="a4">�����Ƿ���������������Ҫָ���Ķ�Ȩ�޲ſ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(18)" value="1" <%if CID(Group_Set_Class(18))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(18)" value="0" <%if CID(Group_Set_Class(18))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ʹ�������ɫ:</b><br>
        <span class="a4">�����Ƿ�����������ʱ������ѡ�����ӵı�����ɫ��</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(19)" value="1" <%if CID(Group_Set_Class(19))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(19)" value="0" <%if CID(Group_Set_Class(19))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ǩ����ʹ�� UBB ����:</b><br>
        <span class="a4">�����Ƿ�����û�ǩ���е� UBB ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(21)" value="1" <%if CID(Group_Set_Class(21))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(21)" value="0" <%if CID(Group_Set_Class(21))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ǩ����ʹ�� [img] ����:</b><br>
        <span class="a4">�����Ƿ�����û�ǩ���е� [img] ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(22)" value="1" <%if CID(Group_Set_Class(22))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(22)" value="0" <%if CID(Group_Set_Class(22))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>���ǩ������:</b> <br>
        <span class="a4">�����û�ǩ������ֽ�����0 Ϊ�������û�ʹ��ǩ��</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(23)" value="<%=Group_Set_Class(23)%>">
      </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="�� ��">
  </center>
  <br>
  <a name="�������"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�������</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��������/�鿴����:</b><br>
        <span class="a4">�����Ƿ�������û����������Ȩ�޵���̳�����ػ�鿴����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(24)" value="1" <%if CID(Group_Set_Class(24))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(24)" value="0" <%if CID(Group_Set_Class(24))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����������:</b><br>
        <span class="a4">�����Ƿ������ϴ���������̳�С���Ҫ����ѡ�������ϴ���������Ч����ο� <B>����ѡ��</B> - <B>��������</B> </span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(25)" value="1" <%if CID(Group_Set_Class(25))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(25)" value="0" <%if CID(Group_Set_Class(25))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>ÿ���ϴ���������:</b><br>
        <span class="a4">�����û�ÿ���ϴ�����ʱ����ͬʱ�ϴ��ĸ���������</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(26)" value="<%=Group_Set_Class(26)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��󸽼��ߴ�(KB):</b><br>
        <span class="a4">���ø�������ֽ�������Ҫ ����ѡ�� �������Ч����������ֻ��С�ڻ���� ����ѡ�� ������������������ ����ѡ�� ��������ִ�С�<BR> Ŀǰ�Ļ������ò���Ϊ��<%=team.Forum_setting(71)%></span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(27)" value="<%=Group_Set_Class(27)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>ÿ���ϴ�������������:</b><br>
        <span class="a4">�����û�ÿ 24 Сʱ�����ϴ��ĸ����ܸ�����0 Ϊ�����ơ�</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(28)" value="<%=Group_Set_Class(28)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����������:</b><br>
        <span class="a4">���������ϴ��ĸ�����չ���������չ��֮���ð�Ƕ��� "," �ָ����Ϊ������</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(29)" value="<%=Group_Set_Class(29)%>">
      </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="�� ��">
  </center>
</form>
<%
	End If
	Rs.Close:Set Rs=Nothing
End Sub
	
Sub EditUserGroup
	Dim myid,Saves,u
	myid = Request.form("myid")
	If myid = "" or (Not IsNumeric(myid)) Then  
		SuccessMsg "��������! "
	Else
		for u=0 to 29
			if Saves ="" Then
				Saves = Request.form("Group_Set_Class(0)")
			else
				Saves = Saves & "|"& Request.form("Group_Set_Class("&u&")")
			End if
		next
		team.execute("Update ["&IsForum&"UserGroup] set IsBrowse = '"&Saves&"',GroupName='"&request.form("GroupName")&"' Where ID="& MyID)
		If Cid(Request.Form("oldgroup")) <> Cid(Request.Form("mygroups")) Then
			team.execute("Update ["&IsForum&"UserGroup] set GroupRank="&Cid(Request.Form("mygroups"))&"  Where ID="& MyID)
		End If
		Application.Contents.RemoveAll()
		SuccessMsg " �û����Ȩ�޸��³ɹ�!"
	End If
End Sub

Sub Editmodel_2
	Dim Rs,Manage_Set_Class
	Set Rs=team.Execute("Select Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,UserColor,UserImg,Rank from "&IsForum&"UserGroup Where ID="&ID)
	If Rs.Eof Then 
		SuccessMsg "��������! "
	Else
		Manage_Set_Class = Split(Rs(5),"|")
	%><br><br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?Action=EditUserManages">
  <input type="hidden" name="myid" value="<%=ID%>">
  <a name="�༭�����Ա�� - <%=RS(1)%>"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">�༭�����Ա�� - <%=RS(1)%></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����༭����:</b><br>
        <span class="a4">�����Ƿ�����༭����Χ�ڵ�����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(0)" value="1" <%if CID(Manage_Set_Class(0))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(0)" value="0" <%if CID(Manage_Set_Class(0))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����ö�����:</b><br>
        <span class="a4">���������ö�����ļ������ö�����ȫ��̳�ö��������ö� ���ڰ���ö���</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(1)" value="0" <%if CID(Manage_Set_Class(1))=0 Then%>checked<%end if%>>
        �������ö�<br>
        <input type="radio" name="Manage_Set_Class(1)" value="1" <%if CID(Manage_Set_Class(1))=1 Then%>checked<%end if%>>
        �������ö�<br>
        <input type="radio" name="Manage_Set_Class(1)" value="2" <%if CID(Manage_Set_Class(1))=2 Then%>checked<%end if%>>
        �������ö�/���ö�<br>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����������:</b><br>
        <span class="a4">�����Ƿ���������û���������ӣ�ֻ����̳������Ҫ���ʱ��Ч</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(2)" value="1" <%if CID(Manage_Set_Class(2))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(2)" value="0" <%if CID(Manage_Set_Class(2))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ɾ������:</b><br>
        <span class="a4">�����Ƿ�����ɾ������Χ�ڵ�����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(3)" value="1" <%if CID(Manage_Set_Class(3))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(3)" value="0" <%if CID(Manage_Set_Class(3))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����ƶ�����:</b><br>
        <span class="a4">�����Ƿ������ƶ�����Χ�ڵ�����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(4)" value="1" <%if CID(Manage_Set_Class(4))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(4)" value="0" <%if CID(Manage_Set_Class(4))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>������ǰ����:</b><br>
        <span class="a4">�����Ƿ�������ǰ����Χ�ڵ�����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(5)" value="1" <%if CID(Manage_Set_Class(5))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(5)" value="0" <%if CID(Manage_Set_Class(5))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��������/�����������:</b><br>
        <span class="a4">�����Ƿ���������/�����������Χ�ڵ�����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(6)" value="1" <%if CID(Manage_Set_Class(6))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(6)" value="0" <%if CID(Manage_Set_Class(6))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����ر�/������:</b><br>
        <span class="a4">�����Ƿ���������/�����������Χ�ڵ�����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(7)" value="1" <%if CID(Manage_Set_Class(7))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(7)" value="0" <%if CID(Manage_Set_Class(7))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����������/�Ƴ�����:</b><br>
        <span class="a4">�����Ƿ���������Χ�ڵ����Ӽ���/�Ƴ�������</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(8)" value="1" <%if CID(Manage_Set_Class(8))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(8)" value="0" <%if CID(Manage_Set_Class(8))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����������������/�Ƴ�ר��:</b><br>
        <span class="a4">�����Ƿ���������Χ�ڵ����Ӽ���/�Ƴ�ר��</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(12)" value="1" <%if CID(Manage_Set_Class(12))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(12)" value="0" <%if CID(Manage_Set_Class(12))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>

    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>����鿴 IP:</b><br>
        <span class="a4">�����Ƿ�����鿴�û� IP</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(10)" value="1" <%if CID(Manage_Set_Class(10))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(10)" value="0" <%if CID(Manage_Set_Class(10))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����ֹ IP:</b><br>
        <span class="a4">�����Ƿ�������ӻ��޸Ľ�ֹ IP ����</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(11)" value="1" <%if CID(Manage_Set_Class(11))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(11)" value="0" <%if CID(Manage_Set_Class(11))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����޸İ�����:</b><br>
        <span class="a4">ֻ�й�������û��ſ����д�Ȩ��</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(9)" value="1" <%if CID(Manage_Set_Class(9))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(9)" value="0" <%if CID(Manage_Set_Class(9))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>�����ֹ�û�:</b><br>
        <span class="a4">�����Ƿ������ֹ�û����������</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(13)" value="1" <%if CID(Manage_Set_Class(13))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(13)" value="0" <%if CID(Manage_Set_Class(13))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>��������û�:</b><br>
        <span class="a4">�����Ƿ����������ע���û���ֻ����̳������Ҫ�˹�������û�ʱ��Ч</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(14)" value="1" <%if CID(Manage_Set_Class(14))=1 Then%>checked<%end if%>>
        �� &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(14)" value="0" <%if CID(Manage_Set_Class(14))=0 Then%>checked<%end if%>>
        �� </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="groupsubmit" value="�� ��">
  <center>
</form>
<%
	End If
End Sub
	
Sub EditUserManages
	Dim myid,Saves,u
	myid = Request.form("myid")
	If myid = "" or (Not IsNumeric(myid)) Then  
		SuccessMsg "��������! "
	Else
		for u=0 to 14
			if Saves ="" Then
				Saves = Request.form("Manage_Set_Class(0)")
			else
				Saves = Saves & "|"& Request.form("Manage_Set_Class("&u&")")
			End if
		next
		team.execute("Update UserGroup set IsManage = '"&Saves&"' Where ID="& MyID)
		Application.Contents.RemoveAll()
		SuccessMsg " ������Ȩ�޸��³ɹ�!"
	End If
End Sub
	
Sub  IsuserGroup	
	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr bgcolor="#F8F8F8">
    <td><br>
      <ul>
        <li>TEAM��̳�û����Ϊϵͳ�顢������ͻ�Ա�飬��Ա���Ի���ȷ������Ȩ�ޣ���ϵͳ�������������Ϊ�趨����������̳ϵͳ���иı䡣
      </ul>
      <ul>
        <li>ϵͳ�����������趨����Ҫָ�����֣�TEAM Ԥ���˴���̳����Ա���ο͵ȵ� 8 ��ϵͳͷ�Σ���������û���Ҫ�ڱ༭��Աʱ������롣
      </ul>
      <ul>
        <li>����û���ָ�����û��飬��ôɾ�����û���𽫵����û��޷�������̳����Ҫ�ֶ��������ø��û����ڵ��û��顣
      </ul>
      <ul>
        <li>����޸����û�������ƣ���ô���������е���̳�û����¸�ֵ���˲��������Ĵ�����ϵͳ��Դ��
      </ul>	  
	  </td>
  </tr>
</table>
<br>
<form method="post" action="?Action=ManagesMember&master=0">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="7">��Ա�û���</td>
    </tr>
	<a name="��Ա�û���"></a>
    <tr class="a4" align="center">
      <td width="48"><input type="checkbox" name="chkall" class="radio" onClick="checkall(this.form)">
        ɾ?</td>
      <td>��ͷ��</td>
      <td>���ִ���</td>
      <td>������</td>
      <td>������ɫ</td>
      <td>��ͷ��</td>
      <td>�༭</td>
    </tr>
	<%
	Dim Rs
	Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,Rank,UserColor,Userimg from "&IsForum&"UserGroup Where GroupTips=1 Order By Memberrank Desc")
	Do While Not Rs.Eof
	%><input type="hidden" name="upid" value="<%=RS(0)%>">
    <tr align="center">
      <td bgcolor="#F8F8F8">
		<input type="checkbox" name="myid" value="<%=RS(0)%>"  class="radio">
	  </td>
      <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName" value="<%=Replace(RS(2)," ","")%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="6" name="Memberrank" value="<%=RS(3)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank" value="<%=RS(7)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor" value="<%=RS(8)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg" value="<%=RS(9)%>"></td>
      <td bgcolor="#FFFFFF" nowrap><a href="?Action=Editmodel_1&ID=<%=RS(0)%>">[����]</a></td>
    </tr>
	<%	Rs.Movenext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
    <tr>
      <td colspan="7" class="a4" height="2"></td>
    </tr>
    <tr align="center" bgcolor="#F8F8F8">
      <td>����:</td>
      <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName1" value=""></td>
      <td bgcolor="#F8F8F8"><input type="text" size="6" name="Memberrank1" value=""></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank1" value=""></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor1" value=""></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg1" value=""></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="groupsubmit" value="�� ��">
  &nbsp;
</form>
<br>
<br>
<form method="post" action="?Action=ManagesMember&master=1">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="6">�����û���</td>
    </tr>
	<a name="�����û���"></a>
    <tr class="a4" align="center">
      <td width="48"><input type="checkbox" name="chkall" class="a4" onClick="checkall(this.form)">
        ɾ?</td>
      <td nowrap>��ͷ��</td>
      <td nowrap>������</td>
      <td nowrap>������ɫ</td>
      <td nowrap>��ͷ��</td>
      <td nowrap>�༭</td>
    </tr>
	<%
	Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,Rank,UserColor,Userimg from "&IsForum&"UserGroup Where GroupTips=2 Order By ID Desc")
	Do While Not Rs.Eof
	%><input type="hidden" name="gupid" value="<%=RS(0)%>">
    <tr align="center">
      <td bgcolor="#F8F8F8">
		<input type="checkbox" name="myid" value="<%=RS(0)%>"  class="radio">
	  </td>
	  <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName" value="<%=Replace(RS(2)," ","")%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank" value="<%=RS(7)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor" value="<%=RS(8)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg" value="<%=RS(9)%>"></td>
      <td bgcolor="#FFFFFF" nowrap><a href="?Action=Editmodel_1&ID=<%=RS(0)%>&OnGroup=yes">[����]</a></td>
    </tr>
	<%  Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
	<tr>
      <td colspan="6" class="a4" height="2"></td>
    </tr>
    <tr align="center" bgcolor="#F8F8F8">
      <td>����:</td>
      <td><input type="text" size="20" name="GroupName1"></td>
      <td><input type="text" size="2" name="Rank1"></td>
      <td><input type="text" size="30" name="UserColor1"></td>
      <td><input type="text" size="20" name="Userimg1"></td>
	  <td>&nbsp;</td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="groupsubmit" value="�� ��">
  </center>
</form>
<br>
<form method="post" action="?Action=ManagesMember&master=2">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="48">ϵͳ�û���</td>
    </tr>
	<a name="ϵͳ�û���"></a>
    <tr class="a4" align="center">
	  <td width="2"></td>
	  <td>ϵͳͷ��</td>
      <td>��ͷ��</td>
      <td>������</td>
      <td>������ɫ</td>
      <td>��ͷ��</td>
      <td>�༭</td>
    </tr>
	<%
	Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,Rank,UserColor,Userimg from "&IsForum&"UserGroup Where MemberRank = -1 Order By GroupRank Desc")
	Do While Not Rs.Eof
	%>
    <tr align="center">
	  <td bgcolor="#FFFFFF"><input type="hidden" name="myid" value="<%=RS(0)%>"></td>
	  <td bgcolor="#F8F8F8"><%=RS(1)%></td>
      <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName" value="<%=Replace(RS(2),"&nbsp;","")%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank" value="<%=RS(7)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor" value="<%=RS(8)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg" value="<%=RS(9)%>"></td>
      <td bgcolor="#FFFFFF" nowrap><a href="?Action=Editmodel_1&ID=<%=RS(0)%>">[����]</a></td>
    </tr>
	<%  Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
  </table>
  <br>
  <center>
    <input type="submit" name="groupsubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
End Sub
Sub ManagesMember
	Dim Myid,GroupName,Rank,Memberrank,UserColor,Userimg,tmp,u
	Dim ho,Rs
	tmp = ""
	GroupName = ""
	If request.form("myid") = "" Then
		Myid=Split(Request.Form("upid"),",")
		GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
		Rank = Split(Replace(Request.Form("Rank")," ",""),",")
		UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
		Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
		Memberrank = Split(Replace(Request.Form("Memberrank")," ",""),",")
		For U=0 To Ubound(Myid)
			Set Rs = team.Execute("Select GroupName,Members From ["&IsForum&"UserGroup] Where ID="&Myid(U))
			If Not (Rs.Eof And Rs.Bof) Then
				If Trim(Rs(0)) <> Trim(GroupName(U)) Then
					team.Execute("Update ["&IsForum&"User] set Members='"&Rs(1)&"',Levelname='"&GroupName(U)&"||"&UserColor(U)&"||"&Userimg(U)&"||"&Cid(Rank(U))&"||0' where UserGroupID="&Myid(U))
				End If
			End If
		Next
	End If

	Select Case Request("master")
		Case "0"
			for each ho in request.form("myid")
				Team.execute("Delete from "&isforum&"UserGroup Where ID="&ho)
			next
			If request.form("myid") = "" Then
				If Request.Form("upid") <> "" Then
					If Instr(Request.Form("upid"),",")>0 Then
						
						Myid=Split(Request.Form("upid"),",")
						GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
						Rank = Split(Replace(Request.Form("Rank")," ",""),",")
						UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
						Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
						Memberrank = Split(Replace(Request.Form("Memberrank")," ",""),",")
						For U=0 To Ubound(Myid)
							team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&GroupName(U)&"',Memberrank="&Cid(Memberrank(U))&",Rank="&Cid(Rank(U))&",UserColor='"&UserColor(U)&"',Userimg='"&Userimg(U)&"' where ID="&Myid(U))
						Next
					Else
						team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&HtmlEncode(Request.Form("GroupName"))&"',Rank="&Cid(Request.Form("Rank"))&",UserColor='"&Request.Form("UserColor")&"',Userimg='"&Request.Form("Userimg")&"' where ID="&Request.Form("upid"))
					End If
				End if
				if Not (Request.Form("GroupName1")="" and Request.Form("Rank1")="") Then
					team.Execute("insert into "&IsForum&"UserGroup (GroupName,GroupTips,Rank,UserColor,Userimg,GroupRank,MemberRank,Members,IsBrowse,IsManage) values ('"&HtmlEncode(Request.Form("GroupName1"))&"',1,"&Cid(Request.Form("Rank1"))&",'"&Request.Form("UserColor1")&"','"&Request.Form("Userimg1")&"',0,"&CID(Request.Form("Memberrank1"))&",'ע���û�','1|1|1|0|0|0|1|0|0|0|0|0|20|1|1|1|0|0|0|0|0|0|0|0|0|0|1|100|50|','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
				End If
			End If
			tmp = tmp & " ��Ա�û��� �޸����óɹ�! "			
		Case "1"
			for each ho in request.form("myid")
				Team.execute("Delete from "&isforum&"UserGroup Where ID="&ho)
			next
			If request.form("myid") = "" Then
				If Request.Form("gupid") <> "" Then
					If Instr(Request.Form("gupid"),",")>0 Then
						Myid=Split(Request.Form("gupid"),",")
						GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
						Rank = Split(Replace(Request.Form("Rank")," ",""),",")
						UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
						Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
						For U=0 To Ubound(Myid)
							team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&GroupName(U)&"',Rank="&Cid(Rank(U))&",UserColor='"&UserColor(U)&"',Userimg='"&Userimg(U)&"' where ID="&Myid(U))
						Next
					Else
						team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&HtmlEncode(Request.Form("GroupName"))&"',Rank="&Cid(Request.Form("Rank"))&",UserColor='"&Request.Form("UserColor")&"',Userimg='"&Request.Form("Userimg")&"' where ID="&Request.Form("gupid"))
					End If
				End if
				if Not (Request.Form("GroupName1")="" and Request.Form("Rank1")="") Then
					team.Execute("insert into "&IsForum&"UserGroup (GroupName,GroupTips,Rank,UserColor,Userimg,GroupRank,MemberRank,Members,IsBrowse,IsManage) values ('"&HtmlEncode(Request.Form("GroupName1"))&"',2,"&Cid(Request.Form("Rank1"))&",'"&Request.Form("UserColor1")&"','"&Request.Form("Userimg1")&"',4,-1,'"&HtmlEncode(Request.Form("GroupName1"))&"','1|255|1|1|1|2|0|1|1|1|1|1|50|1|1|1|0|1|0|0|1|1|1|50|1|1|10|1024|200|','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
				End If
			End If
			tmp = tmp & " �����û��� �޸����óɹ�! "
		Case "2"
			Myid=Split(Request.Form("myid"),",")
			GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
			Rank = Split(Replace(Request.Form("Rank")," ",""),",")
			UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
			Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
			For U=0 To Ubound(Myid)
				team.Execute("Update "&IsForum&"UserGroup set GroupName='"&GroupName(U)&"',Rank="&Rank(U)&",UserColor='"&UserColor(U)&"',Userimg='"&Userimg(U)&"' where ID="&Myid(U))
			Next
			tmp = tmp & " ϵͳ�û��� �޸����óɹ�! "
		Case Else
			tmp = tmp & "��������! "
	End Select
	Application.Contents.RemoveAll()
	SuccessMsg tmp &" <br />��лʹ��TEAM��̳ϵͳ,�Ժ�ϵͳ���Զ���������ҳ��! <meta http-equiv=refresh content=3;url=Admin_Group.asp?Action=IsuserGroup#ϵͳ�û���>"
End Sub
%>
