<!--#include file="../Conn.asp"-->
<!--#include file="Const.asp"-->
<SCRIPT src="../js/Command.js"></SCRIPT>
<script>
function admin_Size(num,objname)
{
    var obj=document.getElementById(objname)
    if (parseInt(obj.rows)+num>=3) {
        obj.rows = parseInt(obj.rows) + num;    
    }
    if (num>0)
    {
        obj.width="90%";
    }
}
</script>
<%
Dim ID
Call Master_Us()
Header()
Dim Admin_Class
Admin_Class=",1,"
Call Master_Se()
Select Case Request("action")
	Case "Settingok"
		Call Settingok
	Case "upreg"
		team.Execute("Update ["&isforum&"Clubconfig] Set AgreeMent='"&Replace(Trim(Request.Form("myinfos")),"'","")&"'")
		Cache.DelCache("Club_Class")
		SuccessMsg("ע��Э��༭�ɹ�!")
	Case Else
		Call Main()
End Select


Sub Main()
	Dim InstalledObjects(10)
	'ˮӡ
	InstalledObjects(0) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
	InstalledObjects(1) = "Persits.Jpeg"				'AspJpeg
	InstalledObjects(2) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
	InstalledObjects(3) = "sjCatSoft.Thumbnail"			'sjCatSoft.Thumbnail V2.6
	'�ϴ�
	InstalledObjects(4) = "Adodb.Stream"				'Adodb.Stream
	InstalledObjects(5) = "Persits.Upload"				'Aspupload3.0
	InstalledObjects(6) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
	InstalledObjects(7) = "Scripting.FileSystemObject"	'FSO
	'�ʼ�
	InstalledObjects(8) = "JMail.Message"
	InstalledObjects(9) = "CDONTS.NewMail"
%>
<form method="Post" action="?action=Settingok">
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2" >
<tr class="a1"><td>������ʾ</td></tr>
<tr class="a3"><td>
<br><ul><li>ѡ���Լ��»��ߵ�б������ʾʱ��˵����ѡ���ϵͳЧ�ʡ�������������Դ�����й�(���Ч�ʡ��򽵵�Ч��)�������������������������е�����
</ul></td></tr></table>
<br>

<a name="��������"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
<td colspan="2">��������</td>
</tr>
<tr>
	<td width="60%"  bgcolor="#F8F8F8"><b>��̳����:</b><br><span class="a3">��̳���ƣ�����ʾ�ڵ������ͱ�����</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(1)" value="<%=team.Club_Class(1)%>"></td>
</tr>
<tr><% team.Club_Class(2) = "" %>
	<td width="60%"  bgcolor="#F8F8F8"><b>��̳URL:</b><br></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(2)" value="<%
	If team.Club_Class(2)="" Then 
		Response.Write "http://"&Request.ServerVariables("server_name")&""&replace(Request.ServerVariables("script_name"),ManagePath&"Admincp.asp","")&"" 
	Else 
		Response.Write team.Club_Class(2)
	End If 
	%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8"><b>��վ����:</b><br><span class="a3">��վ���ƣ�����ʾ��ҳ��ײ�����ϵ��ʽ��</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(3)" value="<%=team.Club_Class(3)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��վ����:</b><br><span class="a3">��վ��ʼ���е����ڣ��ɵ��ò�����</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(29)" value="<%=team.Club_Class(29)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��վ URL:</b><br><span class="a3">��վ URL������Ϊ������ʾ��ҳ��ײ�</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(4)" value="<%=team.Club_Class(4)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��վ����:</b><br><span class="a3">��վ�����ĺ��룬��ʾ��ҳβ���½�</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(59)" value="<%=team.Forum_setting(59)%>"></td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����������ҳ:</b><br><span class="a3">����������ҳ��Ҫ�ռ�֧��INDEX.ASPĬ����ҳ����</span></td>
	<td bgcolor="#FFFFFF"><input type="radio" class="radio" name="Forum_setting(111)" value="1" <%If CID(team.Forum_setting(111))=1 Then%>checked<%End If%>> �� &nbsp; &nbsp; <input type="radio" class="radio" name="Forum_setting(111)" value="0" <%If CID(team.Forum_setting(111))=0 Then%>checked<%End If%>> ��</td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��̳�ر�:</b><br><span class="a3">��ʱ����̳�رգ��������޷����ʣ�����Ӱ�����Ա����</span></td>
	<td bgcolor="#FFFFFF"><input type="radio" class="radio" name="Forum_setting(2)" value="1" <%If team.Forum_setting(2)=1 Then%>checked<%End If%>> �� &nbsp; &nbsp; <input type="radio" class="radio" name="Forum_setting(2)" value="0" <%If team.Forum_setting(2)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>��̳�رյ�ԭ��:</b><br><span class="a3">��̳�ر�ʱ���ֵ���ʾ��Ϣ</span></td>
	<td bgcolor="#FFFFFF"><textarea rows="5" name="Forum_setting(3)" cols="60"><%=team.Forum_setting(3)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center><br>

<br><a name="ע������ʿ���"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">ע������ʿ���</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�������û�ע��:</b><br><span class="a3">ѡ�񡰷񡱽���ֹ�ο�ע���Ϊ��Ա������Ӱ���ȥ��ע��Ļ�Ա��ʹ��</span></td><td bgcolor="#FFFFFF"><input type="radio" class="radio" name="Forum_setting(4)" value="1" <%If team.Forum_setting(4)=1 Then%>checked<%End If%>> �� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(4)" value="0" <%If team.Forum_setting(4)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>��ֹ���û�ע��˵��:</b><br><span class="a3">����̳��ֹ�ο�ע���Ϊ��Աʱ,����������ʾ����!</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Forum_setting(5)" cols="60"><%=team.Forum_setting(5)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����ͬһ Email ע�᲻ͬ�û�:</b><br><span class="a3">ѡ�񡰷񡱽�ֻ����һ�� Email ��ַֻ��ע��һ���û���</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(6)" value="1" <%If team.Forum_setting(6)=1 Then%>checked<%End If%>> �� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(6)" value="0" <%If team.Forum_setting(6)=0 Then%>checked<%End If%>> ��
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Email���֧��:</b><br><span class="a3">��ѡ����ȷ���ʼ��������, ֻ�����֧�ֵ������,�ſ������û�����Email��֤. ����������Email��֤����!</span><br>
	<%If IsObjInstalled(InstalledObjects(8)) Then%>��JMail.Message ��<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(9)) Then%>�� CDONTS.NewMail ��<br><%End If%>
	<%If Not IsObjInstalled(InstalledObjects(9)) and Not IsObjInstalled(InstalledObjects(8)) Then%><font color=red>ע��: ���ķ�������֧�ַ����ʼ�����!</font><%End If%>
	</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(1)" value="0" <%If team.Forum_setting(1)=0 Then%>checked<%End If%>> �� <br>
	<input type="radio" class="radio" name="Forum_setting(1)" value="1" <%If team.Forum_setting(1)=1 Then%>checked<%End If%>> JMail.Message	<br>
	<input type="radio" class="radio" name="Forum_setting(1)" value="2" <%If team.Forum_setting(1)=2 Then%>checked<%End If%>> CDONTS.NewMail	<br></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>SMTP Server��ַ:</b><br><span class="a3">�ʼ��������ĵ�ַ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(58)" value="<%=team.Forum_setting(58)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�ʼ���������¼��:</b><br><span class="a3">��¼�ʼ����������û���</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(41)" value="<%=team.Forum_setting(41)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�ʼ���������¼����:</b><br><span class="a3">��¼�ʼ�������������</span></td><td bgcolor="#FFFFFF"><input type="password" size="30" name="Forum_setting(55)" value="<%=team.Forum_setting(55)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�ʼ������˵�ַ:</b><br><span class="a3">��ʾ���ʼ��ķ����˵�ַ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(57)" value="<%=team.Forum_setting(57)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���û�ע����֤:</b><br><span class="a3">ѡ���ޡ��û���ֱ��ע��ɹ���ѡ��Email ��֤�������û�ע�� Email ����һ����֤�ʼ���ȷ���������Ч�ԣ�ѡ���˹���ˡ����ɹ���Ա�˹����ȷ���Ƿ��������û�ע��</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(7)" value="0" <%If team.Forum_setting(7)=0 Then%>checked<%End If%>> �� <br>
	<input type="radio" class="radio" name="Forum_setting(7)" value="1" <%If team.Forum_setting(7)=1 Then%>checked<%End If%>> ע��ɹ�����Email	<br>
	<input type="radio" class="radio" name="Forum_setting(7)" value="2" <%If team.Forum_setting(7)=2 Then%>checked<%End If%>> Email ��֤	<br>
	<input type="radio" class="radio" name="Forum_setting(7)" value="3" <%If team.Forum_setting(7)=3 Then%>checked<%End If%>> �˹����	<br>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>ע��Email������:</b><br><span class="a3">�û�ע���ʼ�������</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(23)" cols="60"  style="height:70;overflow-y:visible;"><%=team.Club_Class(23)%></textarea>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>IP ע��������(��):</b><br><span class="a3">ͬһ IP �ڱ�ʱ�����ڽ�ֻ��ע��һ���ʺţ����ƶ����޸ĺ����ע���û���Ч��0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(10)" value="<%=team.Forum_setting(10)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��½�û����Դ�������:</b><br><span class="a3">���û���½��������ʱ�� ���Գ�������Ĵ������ƣ���������ΪĬ��ֵ 5 ���ڲ����� 10 ��Χ��ȡֵ��0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(54)" value="<%=team.Forum_setting(54)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>�û���Ϣ�����ؼ���:</b><br><span class="a3">�û������û���Ϣ(���û������ǳơ��Զ���ͷ�ε�)���޷�ʹ����Щ�ؼ��֡�ÿ���ؼ���һ�У���ʹ��ͨ��� "*" �� "*����*"(��������)</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(25)" cols="60"><%=team.Club_Class(25)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���ּ�ϰ����(����):</b><br><span class="a3">��ע���û��ڱ������ڽ��޷���������Ӱ�����͹���Ա��0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(14)" value="<%=team.Forum_setting(14)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���ͻ�ӭ����Ϣ:</b><br><span class="a3">ѡ���ǡ����Զ�����ע���û�����һ����ӭ����Ϣ</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(15)" value="1" <%If team.Forum_setting(15)=1 Then%>checked<%End If%>> �� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(15)" value="0" <%If team.Forum_setting(15)=0 Then%>checked<%End If%>> ��
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>��ӭ����Ϣ����:</b><br><span class="a3">ϵͳ���͵Ļ�ӭ����Ϣ������</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Forum_setting(16)" cols="60" style="height:70;overflow-y:visible;"><%=team.Forum_setting(16)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�οͲ鿴����Ȩ��:</b><br><span class="a3">���������û�����ڲ鿴Ȩ�޵�ʱ�򣬿���ͨ����ѡ�������οͲ鿴��������</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(110)" value="0" <%If CID(team.Forum_setting(110))=0 Then%>checked<%End If%>> ������ <BR> 
	<input type="radio" class="radio" name="Forum_setting(110)" value="1" <%If CID(team.Forum_setting(110))=1 Then%>checked<%End If%>> ��������
	<BR>  
	<input type="radio" class="radio" name="Forum_setting(110)" value="2" <%If CID(team.Forum_setting(110))=2 Then%>checked<%End If%>> ��ȫ����
	</td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ע�����Э��:</b><br><span class="a3">���û�ע��ʱ��ʾ���Э��</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(17)" value="1" <%If team.Forum_setting(17)=1 Then%>checked<%End If%>> �� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(17)" value="0" <%If team.Forum_setting(17)=0 Then%>checked<%End If%>> ��
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>���Э������:</b></td><td bgcolor="#FFFFFF"><a href="#ע��Э��"><span class="a3">����鿴ע�����Э�����ϸ����</span></a></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="��������ʾ��ʽ"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">��������ʾ��ʽ</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Ĭ����̳���:</b><br><span class="a3">��̳Ĭ�ϵĽ������οͺ�ʹ��Ĭ�Ϸ��Ļ�Ա���Դ˷����ʾ</span></td><td bgcolor="#FFFFFF"><select name="Forum_setting(18)">
<%	Dim SQL,RS,SytyleID,MyCheck
	Set Rs=team.Execute("Select ID,StyleName From ["&IsForum&"Style] order by StyleHid Asc")
	Do While Not RS.Eof
		MyCheck = ""
		If Rs(0) = Int(team.Forum_setting(18)) Then 
			MyCheck = " selected=""selected"""
		End if
		Echo  "<option value="""&RS(0)&""""&MyCheck&">"&rs(1)&"</option>"
		Rs.MoveNext
	Loop
	RS.CLOSE:Set RS=Nothing
%>
	</select></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ÿҳ��ʾ������:</b><br><span class="a3">�����б���ÿҳ��ʾ������Ŀ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(19)" value="<%=team.Forum_setting(19)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������ʾ����:</b><br><span class="a3">�����б���ÿ��������ʾ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(88)" value="<%=team.Forum_setting(88)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ÿҳ��ʾ����:</b><br><span class="a3">�����б���ÿҳ��ʾ������Ŀ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(20)" value="<%=team.Forum_setting(20)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ÿҳ��ʾ��Ա��:</b><br><span class="a3">��Ա�б���ÿҳ��ʾ��Ա��Ŀ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(21)" value="<%=team.Forum_setting(21)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�����б����ҳ��:</b><br><span class="a3">�����б����û����Է��ĵ������ҳ������������ΪĬ��ֵ 1000�����ڲ����� 2500 ��Χ��ȡֵ��0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(25)" value="<%=team.Forum_setting(25)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���Ż����������:</b><br><span class="a3">����һ���������Ļ��⽫��ʾΪ���Ż���</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(22)" value="<%=team.Forum_setting(22)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����������ֵ:</b><br><span class="a3">�������ڴﵽ�˷�ֵ(��Ϊ N)ʱ��N ��������ʾΪ 1 ��������N ��������ʾΪ 1 ��̫����Ĭ��ֵΪ 3������Ϊ 0 ��ȡ������ܣ�ʼ����������ʾ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(23)" value="<%=team.Forum_setting(23)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ���������̳����:</b><br><span class="a3">��������̳�б����������У���ʾ��������ʹ�����̳�����б���������������Ϊ 30 ���ڣ�0 Ϊ�رմ˹���</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(24)" value="<%=team.Forum_setting(24)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������ʾ��ʽ:</b><br><span class="a3">��ҳ��̳�б��а�����ʾ��ʽ</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(26)" value="1" <%If team.Forum_setting(26)=1 Then%>checked<%End If%>> ƽ����ʾ  &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(26)" value="0" <%If team.Forum_setting(26)=0 Then%>checked<%End If%>> �����˵�</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ҳ��ʾ��̳���¼�����̳:</b><br><span class="a3">��ҳ��̳�б�������̳�����·���ʾ�¼�����̳���ֺ�����(������ڵĻ�)��ע��: �����ܲ���������̳�������Ȩ�޵������ֻҪ���ڼ��ᱻ��ʾ����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(27)" value="1" <%If team.Forum_setting(27)=1 Then%>checked<%End If%>> ��&nbsp; &nbsp;  
	<input type="radio" class="radio" name="Forum_setting(27)" value="0" <%If team.Forum_setting(27)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���汾��������:</b><br><span class="a3">����̳�����ü��汾�����ʱ��,���ڴ����������Զ�����һ��.</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(32)" value="<%=team.Forum_setting(32)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ��������˵�:</b><br><span class="a3">�����Ƿ�����̳������ʾ���õ���̳��������˵����û�����ͨ���˲˵��л���ͬ����̳���</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(28)" value="1" <%If team.Forum_setting(28)=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(28)" value="0" <%If team.Forum_setting(28)=0 Then%>checked<%End If%>> ��</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ�Զ��������˵�:</b><br><span class="a3">�����Ƿ�����̳������ʾ���õ���̳�Զ��������˵����û�����ͨ���˲˵��л���ͬ����̳������</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(92)" value="1" <%If team.Forum_setting(92)=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(92)" value="0" <%If team.Forum_setting(92)=0 Then%>checked<%End If%>> ��</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ��������:</b><br><span class="a3">�����Ƿ�����̳��ҳ��ʾ<B>��������</B>״̬����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(36)" value="1" <%If team.Forum_setting(36)=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(36)" value="0" <%If team.Forum_setting(36)=0 Then%>checked<%End If%>> ��</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ����״��:</b><br><span class="a3">�����Ƿ�����̳��ҳ��ʾ<B>����״��</B>״̬����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(40)" value="1" <%If team.Forum_setting(40)=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(40)" value="0" <%If team.Forum_setting(40)=0 Then%>checked<%End If%>> ��</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ��������:</b><br><span class="a3">�����Ƿ�����̳��ҳ��ʾ<B>��������(3��)</B>״̬����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(113)" value="1" <%If CID(team.Forum_setting(113))=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(113)" value="0" <%If CID(team.Forum_setting(113))=0 Then%>checked<%End If%>> ��</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���ٷ���:</b><br><span class="a3">�����̳������ҳ��ײ���ʾ���ٷ�����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(29)" value="1" <%If team.Forum_setting(29)=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(29)" value="0" <%If team.Forum_setting(29)=0 Then%>checked<%End If%>> ��</td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="���������Ż�"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">���������Ż�</td></tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����α��̬:</b><br><span class="a3">��ҳ������Ϊ��̬�ļ�����Ҫ�ռ�֧�֡�������ǰ����ռ���ѯ�ʣ�������TEAMר�������ļ���</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(65)" value="1" <%If CID(team.Forum_setting(65))=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(65)" value="0" <%If CID(team.Forum_setting(65))=0 Then%>checked<%End If%>> ��</td>
</tr




<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���⸽����:</b><br><span class="a3">��ҳ����ͨ�������������ע���ص㣬�����������ý������ڱ�������̳���Ƶĺ��棬����ж���ؼ��֣������� "|"��","(��������) �ȷ��ŷָ��� </span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(66)" value="<%=team.Forum_setting(66)%>">
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Meta Keywords:</b><br><span class="a3">Keywords �������ҳ��ͷ���� Meta ��ǩ�У����ڼ�¼��ҳ��Ĺؼ��֣�����ؼ��ּ����ð�Ƕ��� "," ����</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(30)" value="<%=team.Forum_setting(30)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Meta Description:</b><br><span class="a3">Description ������ҳ��ͷ���� Meta ��ǩ�У����ڼ�¼��ҳ��ĸ�Ҫ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(31)" value="<%=team.Forum_setting(31)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>����ͷ����Ϣ:</b><br><span class="a3">������ &lt;head&gt;&lt;/head&gt; ����������� html ���룬����ʹ�ñ����ã�����������</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(26)" cols="60"  style="height:70;overflow-y:visible;"><%=team.Club_Class(26)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="��̳����"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">��̳����</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>���� RSS</i></u>:</b><br><span class="a3">ѡ���ǡ�����̳�������û�ʹ�� RSS �ͻ�������������µ���̳���Ӹ��¡�ע��: �ڷ���̳�ܶ������£������ܿ��ܻ���ط���������</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(33)" value="1" <%If team.Forum_setting(33)=1 Then%>checked<%End If%>>
	�� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(33)" value="0" <%If team.Forum_setting(33)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>RSS TTL(����)</i></u>:</b><br><span class="a3">TTL(Time to Live) �� RSS 2.0 ��һ�����ԣ����ڿ��ƶ������ݵ��Զ�ˢ��ʱ�䣬ʱ��Խ��������ʵʱ�Ծ�Խ�ߣ�������ط�����������ͨ��������Ϊ 30��180 ��Χ�ڵ���ֵ</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(34)" value="<%=team.Forum_setting(34)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��վ�ű�����ʱ��:</b><br><span class="a3">ʹ��Server.scripttimeout������ASP��������ʹ������̱����Ĭ������Ϊ20�룬���ȶȵ��ڷ������������á�</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(91)" value="<%=team.Forum_setting(91)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ����������Ϣ:</b><br><span class="a3">ѡ���ǡ�����ҳ�Ŵ���ʾ��������ʱ������ݿ��ѯ����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(37)" value="1" <%If team.Forum_setting(37)=1 Then%>checked<%End If%>> �� &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(37)" value="0" <%If team.Forum_setting(37)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������ʾ���ʼ�ף��:</b><br><span class="a3">�����Ƿ�����ҳ��ʾ���չ����յĻ�Ա���������Ƿ����ʼ�ף�������������̳�û������ܴ󣬹����ջ�Ա���б���ܻ�Ӱ����ҳҳ�����ۣ�������������ʼ�ף��Ҳ��ķ�һ����ϵͳ��Դ</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(38)" value="0" <%If team.Forum_setting(38)=0 Then%>checked<%End If%>> ��<br>
	<input type="radio" class="radio" name="Forum_setting(38)" value="1" <%If team.Forum_setting(38)=1 Then%>checked<%End If%>> ������ҳ��ʾ�����ջ�Ա<br>
	<input type="radio" class="radio" name="Forum_setting(38)" value="2" <%If team.Forum_setting(38)=2 Then%>checked<%End If%>> ��������ջ�Ա�����ʼ�ף��<br>
	<input type="radio" class="radio" name="Forum_setting(38)" value="3" <%If team.Forum_setting(38)=3 Then%>checked<%End If%>> ��ʾ�������ʼ�ף��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��ʾ�����û�:</b><br><span class="a3">����ҳ����̳�б�ҳ��ʾ���߻�Ա�б�</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(39)" value="0" <%If team.Forum_setting(39)=0 Then%>checked<%End If%>> ����ʾ<br>
	<input type="radio" class="radio" name="Forum_setting(39)" value="1" <%If team.Forum_setting(39)=1 Then%>checked<%End If%>> ������ҳ��ʾ<br>
	<input type="radio" class="radio" name="Forum_setting(39)" value="2" <%If team.Forum_setting(39)=2 Then%>checked<%End If%>> ���ڷ���̳��ʾ<br>
	<input type="radio" class="radio" name="Forum_setting(39)" value="3" <%If team.Forum_setting(39)=3 Then%>checked<%End If%>> ����ҳ�ͷ���̳��ʾ</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>��ʾ��̳��ת�˵�</i></u>:</b><br><span class="a3">ѡ���ǡ������б�ҳ���²���ʾ�����ת�˵���ע��: ������̳�ܶ�ʱ�������ܻ����ؼ��ط���������</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(42)" value="1" <%If team.Forum_setting(42)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(42)" value="0" <%If team.Forum_setting(42)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�������´��ڴ�:</b><br><span class="a3">ѡ���Ƿ��������б�Ĵ򿪷�ʽ�Ƿ�Ϊ�´��ڴ򿪻����ڱ����ڴ�</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(43)" value="1" <%If team.Forum_setting(43)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(43)" value="0" <%If team.Forum_setting(43)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>ͳ��ϵͳ����ʱ��(����)</i></u>:</b><br><span class="a3">ͳ�����ݻ�����µ�ʱ�䣬��ֵԽ�����ݸ���Ƶ��Խ�ͣ�Խ��Լ��Դ��������ʵʱ�̶�Խ�ͣ���������Ϊ 60 ���ϣ�����ռ�ù���ķ�������Դ��</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(44)" value="<%=team.Forum_setting(44)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>�û�����ʱ�����ʱ��(����)</i></u>:</b><br><span class="a3">TEAM! ��ͳ��ÿ���û��ܹ��͵��µ�����ʱ�䣬�����������趨�����û�����ʱ���ʱ��Ƶ�ʡ���������Ϊ 10�����û�ÿ���� 10 ���Ӹ���һ�μ�¼��������ֵԽС����ͳ��Խ��ȷ����������ԴԽ�󡣽�������Ϊ 5��30 ��Χ�ڣ�0 Ϊ����¼�û�����ʱ��</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(45)" value="<%=team.Forum_setting(45)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>�����¼����ʱ��(��)</i></u>:</b><br><span class="a3">ϵͳ�б��������¼��ʱ�䣬Ĭ��Ϊ 3 ���£������� 3~6 ���µķ�Χ��ȡֵ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(47)" value="<%=team.Forum_setting(47)%>"></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>
<br>
<a name="��ȫ����"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
	<td colspan="2">��ȫ����</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����CC����:</b><br><span class="a3">����CC����������ʹ�ô�����������û��޷�������̳��������ڱ���ʱ�򿪡�</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(106)" value="1" <%If team.Forum_setting(106)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(106)" value="0" <%If team.Forum_setting(106)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������֤��:</b><br><span class="a3">ͼƬ��֤����Ա����ù�ˮ��ˢ�³�����������������ύ��Ϣ����ѡ����Ҫ����֤��Ĳ��������õ��û���¼�����ļ��βų�����֤����֤ҳ�棬����Ϊ0ʱ����رմ˹���</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(48)" value="<%=team.Forum_setting(48)%>">
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������Դ���:</b><br><span class="a3">Ϊ�˷�ֹ�û����ⲿ�ύ����!</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(49)" value="1" <%If team.Forum_setting(49)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(49)" value="0" <%If team.Forum_setting(49)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������ˮԤ��(��):</b><br><span class="a3">���η������С�ڴ�ʱ�䣬�����η��Ͷ���Ϣ���С�ڴ�ʱ��Ķ���������ֹ��0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(50)" value="<%=team.Forum_setting(50)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����ʱ������(��):</b><br><span class="a3">�����������С�ڴ�ʱ�佫����ֹ��0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(51)" value="<%=team.Forum_setting(51)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>60 �������������:</u></i></b><br><span class="a3">��̳ϵͳÿ 60 ��ϵͳ��Ӧ���������������0 Ϊ�����ơ�ע��: ����������������أ���������Ϊ 5������ 5~20 ��Χ��ȡֵ���Ա������Ƶ��������������ݱ���</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(52)" value="<%=team.Forum_setting(52)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����������:</b><br><span class="a3">ÿ��������ȡ�������������������ΪĬ��ֵ 500�����ڲ����� 1500 ��Χ��ȡֵ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(53)" value="<%=team.Forum_setting(53)%>"></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="ʱ��μ���������"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">ʱ��μ���������</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>��̳ʱ�������:</b><br><span class="a3">����ÿ�����ʱ������û��ķ���Ȩ�޼�����Ȩ�ޣ��˹��ܶ�Ӧ�����ʱ������ù��ܡ��򿪴˹��ܣ������ʱ������ù��ܲŻῪ����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(56)" value="0" <%If team.Forum_setting(56)=0 Then%>checked<%End If%>> �ر�<br> 
	<input type="radio" class="radio" name="Forum_setting(56)" value="1" <%If team.Forum_setting(56)=1 Then%>checked<%End If%>> ��ʱ�ر�<br> 
	<input type="radio" class="radio" name="Forum_setting(56)" value="2" <%If team.Forum_setting(56)=2 Then%>checked<%End If%>> ��ʱֻ��<br> 
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>��ֹ����ʱ���:</b><br><span class="a3">ÿ���ʱ������û����ܷ�������Ҫ������Ĺ������ʹ��!</span></td><td bgcolor="#FFFFFF"><table cellspacing="1" cellpadding="1" width="90%"><tr><%
		Dim Openclock
		openclock=Split(team.Forum_setting(0),"*")
		For i= 0 to UBound(openclock)
			%>
			<td><input type="checkbox" name="openclock<%=i%>" value="1" <%If openclock(i)="1" Then %>checked<%End If%>><%=i%>�㿪</td>
			<%
			If (i+1) mod 3 = 0 Then Response.Write "</tr><tr>"
		Next
 %></table></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>�û����ӹ�������:</b><br><span class="a3">�Զ����θ��б���ʾ���û���������������ݣ�ÿ��һ���û�����
	</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(7)" cols="60"><%=team.Club_Class(7)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>������������:</b><br><span class="a3">�Զ��滻�û������г��ֵĸ��ֹؼ��֣�ƥ���ʽ���£���Ҫ���ǵ���=�滻����֣�������ɺ󣬽�ֻ��ʾ���˺�����֣�ÿ��һ��ƥ��Ρ�
	</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(5)" cols="60"><%=team.Club_Class(5)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>�û�IP��������:</b><br><span class="a3">���Ƹ�IP���û���½����̳��ÿ��һ��IP�����Բ���ͨ��� * ����IP�ε����ƣ���192.168.1.*������ʾ��IP 192.168.1.1 - 192.168.1.255������IP����Ȩ�ޡ�</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(6)" cols="60"><%=team.Club_Class(6)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="�û�Ȩ��"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">�û�Ȩ��</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������������:</b><br><span class="a3">���ð���ֻ������������Ͻ����̳��Χ�ڶ����ӽ������֡�������ֻ�԰�����Ч���������ֵ���ͨ�û�����������������Ա���ܴ����ƣ�������������Щ�û�����Ȩ�ޣ������Խ�������ȫ��̳��Χ�ڽ�������</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(61)" value="1" <%If team.Forum_setting(61)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(61)" value="0" <%If team.Forum_setting(61)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�����ظ�����:</b><br><span class="a3">ѡ���ǡ��������û���һ�����ӽ��ж�����֣�Ĭ��Ϊ����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(62)" value="1" <%If team.Forum_setting(62)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(62)" value="0" <%If team.Forum_setting(62)=0 Then%>checked<%End If%>> ��</td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����ʱ������(Сʱ):</b><br><span class="a3">���ӷ���󳬹���ʱ�����������û������ܶԴ������֣������͹���Ա���ܴ����ƣ�0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(60)" value="<%=team.Forum_setting(60)%>"></td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����������𱨸�����:</b><br><span class="a3">�����Աͨ������Ϣ����������Ա���淴ӳ���ӡ�ע��: �����ǰ��̳�����û�����ð�����ͬʱ���趨����Ϊ��ֻ���������������ϵͳ���Զ����������ݷ��͸������������Դ�����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(63)" value="0" <%If team.Forum_setting(63)=0 Then%>checked<%End If%>> ��ֹ�û�����<br> 
	<input type="radio" class="radio" name="Forum_setting(63)" value="1" <%If team.Forum_setting(63)=1 Then%>checked<%End If%>> ���������������<br>
	<input type="radio" class="radio" name="Forum_setting(63)" value="2" <%If team.Forum_setting(63)=2 Then%>checked<%End If%>> ������������ͳ�����������<br>
	<input type="radio" class="radio" name="Forum_setting(63)" value="3" <%If team.Forum_setting(63)=3 Then%>checked<%End If%>> �������й�����Ա����
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���Ӻͱ�����С����(�ֽ�):</b><br><span class="a3">�������Ա��ͨ���������������ơ����ö�����Ӱ�죬0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(64)" value="<%=team.Forum_setting(64)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�����������(�ֽ�):</b><br><span class="a3">�������Ա��ͨ���������������ơ����ö�����Ӱ�죬0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(89)" value="<%=team.Forum_setting(89)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�����������(�ֽ�):</b><br><span class="a3">�������Ա��ͨ���������������ơ����ö�����Ӱ��</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(67)" value="<%=team.Forum_setting(67)%>"></td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ͶƱ���ѡ����:</b><br><span class="a3">�趨����ͶƱ���������ѡ����</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(68)" value="<%=team.Forum_setting(68)%>"></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="��������"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">��������</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��������ʾͼƬ����:</b><br><span class="a3">��������ֱ�ӽ�ͼƬ�򶯻�������ʾ������������Ҫ�����������</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(69)" value="1" <%If team.Forum_setting(69)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(69)" value="0" <%If team.Forum_setting(69)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���ظ�������վĿ¼·��:</b><br><span class="a3">�����˹��ܣ��û������ܲ鿴����������·�����Է�ֹ������һ������</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(93)" value="1" <%If team.Forum_setting(93)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(93)" value="0" <%If team.Forum_setting(93)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ѡ���ϴ����:</b><br><span class="a3">
	��ѡ����ʵ��ϴ���ʽ��<br>
	<%If IsObjInstalled(InstalledObjects(4)) Then%>�� �����  ��<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(5)) Then%>�� Aspupload  ��<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(6)) Then%>�� SA-FileUp  ��<%End If%></span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(70)" value="999" <%If team.Forum_setting(70)="999" Then%>checked<%End If%>> �ر� <BR>
	<input type="radio" class="radio" name="Forum_setting(70)" value="0" <%If team.Forum_setting(70)=0 Then%>checked<%End If%>> ������ϴ��� <BR>
	<input type="radio" class="radio" name="Forum_setting(70)" value="1" <%If team.Forum_setting(70)=1 Then%>checked<%End If%>> Aspupload3.0��� <BR>
	<input type="radio" class="radio" name="Forum_setting(70)" value="2" <%If team.Forum_setting(70)=2 Then%>checked<%End If%>> SA-FileUp 4.0���
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�������ļ��Ĵ�С:</b><br><span class="a3">������̳���еȼ��û�ÿ���ϴ��ĸ�����С���������û����������ÿ��������ϴ���С���ƣ���СΪ�������˴�������Ϊ׼��</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(71)" value="<%=team.Forum_setting(71)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�����û�ͷ�񸽼��Ĵ�С:</b><br><span class="a3">�����û��ϴ���ͷ�񸽼���С��</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(72)" value="<%=team.Forum_setting(72)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ͷ��ߴ�����(���|�߶�):</b><br><span class="a3">Ĭ��120*120</span></td><td bgcolor="#FFFFFF">
	��: <input type="text" size="10" name="Forum_setting(108)" value="<%=team.Forum_setting(108)%>">&nbsp;
	��: <input type="text" size="10" name="Forum_setting(109)" value="<%=team.Forum_setting(109)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�����û��ϴ�����������:</b><br><span class="a3">���ñ���̳�������ϴ��ĸ�����չ���������չ��֮�������� "|" �ָ��: rar|jpg|txt���˴�����Ϊ�ܵ������ϴ����ļ����ͣ��û�����Զ�������ÿ�������ϸ���ࡣϵͳ�Զ�������EXE��ASPΪ��׺���ļ����͡�</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(73)" value="<%=team.Forum_setting(73)%>"></td>
</tr>
<tr class="a1"><td colspan="2">ͼƬˮӡ����</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ѡȡˮӡ���:</b><br><span class="a3">�򿪴˹��ܣ�ϵͳ���Զ�Ϊ�û��ϴ���ͼƬ���ˮӡЧ�����˹�����Ҫ�������֧�֣��ݲ�֧�ֶ��� GIF ��ʽ����ѡ����ʵ�ˮӡ�����<br>
	<%If IsObjInstalled(InstalledObjects(0)) Then%>�� CreatePreviewImage  ��<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(1)) Then%>�� AspJpeg���  ��<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(2)) Then%>�� SA-ImgWriter  ��<%End If%>
	<%If Not (IsObjInstalled(InstalledObjects(0)) or IsObjInstalled(InstalledObjects(1)) or IsObjInstalled(InstalledObjects(2)) ) Then%><font color=red> ϵͳ��֧���κ�ˮӡ��� </font><%End If%>
</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(74)" value="999" <%If team.Forum_setting(74)=999 Then%>checked<%End If%>>
	�ر� <br>
	<input type="radio" class="radio" name="Forum_setting(74)" value="0" <%If team.Forum_setting(74)=0 Then%>checked<%End If%>> CreatePreviewImage��� <br>
	<input type="radio" class="radio" name="Forum_setting(74)" value="1" <%If team.Forum_setting(74)=1 Then%>checked<%End If%>> AspJpeg��� <br>
	<input type="radio" class="radio" name="Forum_setting(74)" value="2" <%If team.Forum_setting(74)=2 Then%>checked<%End If%>> SA-ImgWriter���</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ˮӡЧ�����ÿ���:</b><br><span class="a3">�˹�����Ҫ������������֧�֣�ֻ�����֧�֣���ѡ���˺��ʵ����֧�֣��ſ���ʹ�ñ����ܡ�</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(75)" value="0" <%If team.Forum_setting(75)=0 Then%>checked<%End If%>> �ر�ˮӡЧ�� <br>
	<input type="radio" class="radio" name="Forum_setting(75)" value="1" <%If team.Forum_setting(75)=1 Then%>checked<%End If%>> ͼƬˮӡЧ��<br>
	<input type="radio" class="radio" name="Forum_setting(75)" value="2" <%If team.Forum_setting(75)=2 Then%>checked<%End If%>> ����ˮӡЧ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ͼƬ��ַ:</b><br><span class="a3">Ĭ��ˮӡͼƬλ�� images/uplogo.gif�������滻���ļ����޸����µ�ͼƬ��ַ��ʵ�ֲ�ͬ��ˮӡЧ����</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(76)" value="<%=team.Forum_setting(76)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ˮӡͼƬ��������������(���|�߶�):</b><br><span class="a3">Ĭ��88*31</span></td><td bgcolor="#FFFFFF">
	��: <input type="text" size="10" name="Forum_setting(77)" value="<%=team.Forum_setting(77)%>">&nbsp;
	��: <input type="text" size="10" name="Forum_setting(35)" value="<%=team.Forum_setting(35)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����Ԥ��ͼƬ��С:</b><br><span class="a3"></span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(78)" value="1" <%If team.Forum_setting(78)=1 Then%>checked<%End If%>> �̶� 
	<input type="radio" class="radio" name="Forum_setting(78)" value="0" <%If team.Forum_setting(78)=0 Then%>checked<%End If%>> �ȱ�����С</td></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����ˮӡ:</b><br><span class="a3">��֧�ִ�����</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(79)" value="<%=team.Forum_setting(79)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ˮӡ�����С:</b><br><span class="a3">���ˮӡ���ֵ������С</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(80)" value="<%=team.Forum_setting(80)%>">PX</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ˮӡ������ɫ:</b><br><span class="a3">���ˮӡ���ֵ�������ɫ</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(81)" value="<%=team.Forum_setting(81)%>" id=d_bgcolor> 
	<img border="0" src="images/rect.gif" style="cursor:pointer;background-Color:<%=team.Forum_setting(81)%>;" width="18" id="s_bgcolor" onclick="SelectColor('bgcolor')">
	<Script>
	function SelectColor(what){
		var dEL = document.getElementById("d_"+what);
		var sEL = document.getElementById("s_"+what);
		var arr = showModalDialog("images/selcolor.htm", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
		if (arr) {
			dEL.value=arr;
			sEL.style.backgroundColor=arr;
		}
	}
	</script></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>͸������ɫ����:</b><br><span class="a3">���ˮӡ͸������ɫ</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(107)" value="<%=team.Forum_setting(9)%>" id="d_bgcolor1"> 
	<img border="0" src="images/rect.gif" style="cursor:pointer;background-Color:<%=team.Forum_setting(9)%>;" width="18" id="s_bgcolor" onclick="SelectColor('bgcolor1')">
	<Script>
	function SelectColor(what){
		var dEL = document.getElementById("d_"+what);
		var sEL = document.getElementById("s_"+what);
		var arr = showModalDialog("images/selcolor.htm", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
		if (arr) {
			dEL.value=arr;
			sEL.style.backgroundColor=arr;
		}
	}
	</script></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ˮӡ��������:</b><br><span class="a3">���ˮӡ���ֵ�����</span></td><td bgcolor="#FFFFFF">
	<SELECT name="Forum_setting(82)">
		<option value="����" <%SetColors("����")%>>����</option>
		<option value="����_GB2312" <%SetColors("����_GB2312")%>>����</option>
		<option value="������" <%SetColors("������")%>>������</option>
		<option value="����" <%SetColors("����")%>>����</option>
		<option value="����" <%SetColors("����")%>>����</option>
		<OPTION value="Andale Mono" <%SetColors("Andale Mono")%>>Andale Mono</OPTION> 
		<OPTION value="Arial" <%SetColors("Arial")%>>Arial</OPTION> 
		<OPTION value="Arial Black" <%SetColors("Arial Black")%>>Arial Black</OPTION> 
		<OPTION value="Book Antiqua" <%SetColors("Book Antiqua")%>>Book Antiqua</OPTION>
		<OPTION value="Century Gothic" <%SetColors("Century Gothic")%>>Century Gothic</OPTION> 
		<OPTION value="Comic Sans MS" <%SetColors("Comic Sans MS")%>>Comic Sans MS</OPTION>
		<OPTION value="Courier New" <%SetColors("Courier New")%>>Courier New</OPTION>
		<OPTION value="Georgia" <%SetColors("Georgia")%>>Georgia</OPTION>
		<OPTION value="Impact" <%SetColors("Impact")%>>Impact</OPTION>
		<OPTION value="ahoma" <%SetColors("ahoma")%>>Tahoma</OPTION>
		<OPTION value="Times New Roman" <%SetColors("Times New Roman")%>>Times New Roman</OPTION>
		<OPTION value="Trebuchet MS" <%SetColors("Trebuchet MS")%>>Trebuchet MS</OPTION>
		<OPTION value="Script MT Bold" <%SetColors("Script MT Bold")%>>Script MT Bold</OPTION>
		<OPTION value="Stencil" <%SetColors("Stencil")%>>Stencil</OPTION>
		<OPTION value="Verdana" <%SetColors("Verdana")%>>Verdana</OPTION>
		<OPTION value="Lucida Console" <%SetColors("Lucida Console")%>>Lucida Console</OPTION>
	</SELECT></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ͼƬ�������ˮӡ������:</b><br><span class="a3">���ڴ�ѡ��ˮӡ��ӵ�λ��(�� 5 ��λ�ÿ�ѡ)��</span></td><td bgcolor="#FFFFFF">
	<table border="0" cellspacing="1" cellpadding="4" width="100%" class="a2">
	<tr class="tab4">
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="0" <%If CID(team.Forum_setting(83))=0 Then%>checked<%End If%>> ���� 
	 </td>
	 <td>&nbsp;</td>
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="3" <%If CID(team.Forum_setting(83))=3 Then%>checked<%End If%>> ����
	 </td>
	</tr>
	<tr class="tab4">
	 <td>&nbsp;</td>
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="2" <%If CID(team.Forum_setting(83))=2 Then%>checked<%End If%>> ���� </td> 
	 <td>&nbsp;</td>
	</tr>
	<tr class="tab4">
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="1" <%If CID(team.Forum_setting(83))=1 Then%>checked<%End If%>> ����
	 </td>
	 <td>&nbsp;</td>
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="4" <%If CID(team.Forum_setting(83))=4 Then%>checked<%End If%>> ����
	 </td>
	</tr>
	</table>
</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ˮӡ�ں϶�:</b><br><span class="a3">����ˮӡͼƬ��ԭʼͼƬ���ں϶ȣ���ֵԽ��ˮӡͼƬ͸����Խ�͡���������Ҫ����ˮӡ���ܺ����Ч</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(84)" value="<%=team.Forum_setting(84)%>">��60%����д0.6</td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="JS ����"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">JS ����</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���� JS ����:</b><br><span class="a3">JS(JavaScript)���ý�ʹ�������Խ���̳���������е�����Ƕ�뵽������ͨ��ҳ�У����������������̳���ɻ�֪��̳������µ����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(85)" value="1" <%If team.Forum_setting(85)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(85)" value="0" <%If team.Forum_setting(85)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>JS ���ݻ���ʱ��(��):</b><br><span class="a3">����һЩ������������ȽϺķ���Դ��JS ���ó�����û��漼����ʵ�����ݵĶ��ڸ��£�Ĭ��ֵ 1440 ���� ���������ò����� 600 ����ֵ��0 Ϊ������(���ķ�ϵͳ��Դ)</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(86)" value="<%=team.Forum_setting(86)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>JS ��·����:</b><br><span class="a3">Ϊ�˱���������վ�Ƿ�������̳���ݣ��������ķ������������������������������̳ JS ����·�����б�ֻ�����б��е���������վ������ͨ�� JS ��������̳����Ϣ��ÿ������һ�У���֧��ͨ������������ http:// ���������������ݣ�����Ϊ��������·�����κ���վ���ɵ���</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(28)" cols="60"><%=team.Club_Class(28)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center>

<br><a name="��������"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
	<td colspan="2">��������</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Ĭ��ʱ��:</b><br><span class="a3">����ʱ���� GMT ��ʱ��</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(90)" value="<%=team.Forum_setting(90)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�༭����ʱ������(����):</b><br><span class="a3">�������߷����󳬹���ʱ�����ƽ������ٱ༭���������͹���Ա���ܴ����ƣ�0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(94)" value="<%=team.Forum_setting(94)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�༭���Ӹ��ӱ༭��¼:</b><br><span class="a3">�� 60 ���༭������ӡ������� xxx �� xxxx-xx-xx �༭������������Ա�༭���ܴ�����</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(95)" value="1" <%If team.Forum_setting(95)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(95)" value="0" <%If team.Forum_setting(95)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�򿪹����Զ����Ź���:</b><br><span class="a3">ѡ���ǡ������Զ����Ź��ʵĹ��ܣ�ϵͳ�Զ���ÿ�µĵ�һ�췢�Ź��ʸ��û���������������� <B>�û�����</B> �� <B>���ʹ���</B> ѡ�</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(96)" value="1" <%If team.Forum_setting(96)=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(96)" value="0" <%If team.Forum_setting(96)=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��̳ͷ�������:</b><br><span class="a3">�뽫ͷ��ͼƬ�ϴ�����̳�� images/Upface/���棬��������ԭ���ĸ�ʽ���У���Ĭ��30��ͼƬΪ��1-30��������ӵ����31��ʼ�����ڴ˴���д��ȷ��ͼƬ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(100)" value="<%=team.Forum_setting(100)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Ĭ�Ϸ���ģʽ:</b><br><span class="a3">����Ĭ�ϵķ���ģʽ��</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(98)" value="1" <%If team.Forum_setting(98)=1 Then%>checked<%End If%>> ����������ģʽ 
	<input type="radio" class="radio" name="Forum_setting(98)" value="0" <%If team.Forum_setting(98)=0 Then%>checked<%End If%>> UBBģʽ</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>�����������������ѡ��</b><br><span class="a3">���趨�����û�ִ�в��ֹ������������ʱ��ʾ��ÿ������һ�У������������ʾһ�зָ�����--------�����û���ѡ���趨��Ԥ�õ�����ѡ�����������</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(8)" cols="60"><%=team.Club_Class(8)%></textarea>
	</td>
</tr>
<tr>
	<td colspan="2">CC��Ƶ��������</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����CC��Ƶ���˵�����ID</b><br><span class="a3">CC��Ƶ����Ϊ��վ����̳�������ṩרҵ����Ƶ�������񣨰����ϴ���¼�ơ�����ȣ���ͬʱʹ������ƥ�似��Ϊ�û�����ԴԴ���ϵ���Ƶ�㲥��,��ֻҪ�� <A HREF="http://union.bokecc.com/">http://union.bokecc.com/</A>ע��,Ȼ������ID������д���˴�����.  </span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(114)" value="<%=team.Forum_setting(114)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����CC��Ƶ���˵�չ��ID</b><br><span class="a3"> </span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(115)" value="<%=team.Forum_setting(115)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�Ƿ񿪷�չ��:</b></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(116)" value="1" <%If CID(team.Forum_setting(116))=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(116)" value="0" <%If CID(team.Forum_setting(116))=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td colspan="2">www.fs2you.com�Ĵ󸽼�֧��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>����fs2you���̹���:</b></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(119)" value="1" <%If CID(team.Forum_setting(119))=1 Then%>checked<%End If%>> �� 
	<input type="radio" class="radio" name="Forum_setting(119)" value="0" <%If CID(team.Forum_setting(119))=0 Then%>checked<%End If%>> ��</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>fs2you�ʺ�:</b><br><span class="a3"> fs2you��վ�ṩ��&lt;=1G�ĸ����ϴ�����,����������˴˹���,�Ϳ�������fs2you.com�Ŀռ�������������̳����Ҫ�Ĵ�ߴ總��. ����ǰ������������fs2you���ʺ�. ������<A HREF="http://www.fs2you.com/">http://www.fs2you.com/</A>ע��,Ȼ�������ʺ���д���˴�����.  </span>
	</td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(117)" value="<%=team.Forum_setting(117)%>">
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center><br>
<a name="��������"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
<td colspan="2">��������</td>
</tr>
<tr class="a4"><td colspan="2">
	<table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
	<tr class="a4"><td><BR><Ul>
		<li>��ע: ExtCredits0 Ϊϵͳ��������,����ɾ����</li>
		<li>ExtCredits0 Ϊ��ͨ�û���Ȩ�޵ȼ����ն���!</li> 
		<li>ֻ�п���<B>���׻�������</B>ѡ���̳��Ӧ�Ļ��ֹ��ܣ��繤�ʷ��ţ����͹��ܲſ��Կ���!</li>
		</ul></td></tr>
	</table>
</td></tr>
<tr><td colspan="2" bgcolor="#F8F8F8">
<table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
<tr class="a1"><td colspan="7">��չ��������</td></tr>
<tr align="center" class="a3">
	<td>���ִ���</td>
	<td>��������</td>
	<td>���ֵ�λ</td>
	<td>ע���ʼ����</td>
	<td>���ô˻���</td>
	<td>����������ʾ</td>
</tr>
<%
Dim ExtCredits,ExtSort,MustOpen,MustSort,U,M
ExtCredits= Split(team.Club_Class(21),"|")
For U=0 to Ubound(ExtCredits)
	ExtSort=Split(ExtCredits(U),",")
%>
<tr align="center">
	<td bgcolor="#F8F8F8">ExtCredits<%=U%></td>
	<td bgcolor="#FFFFFF"><input type="text" size="8" name="ExtCredits<%=U%>_0" value="<%=ExtSort(0)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="5" name="ExtCredits<%=U%>_1" value="<%=ExtSort(1)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="3" name="ExtCredits<%=U%>_2" value="<%=ExtSort(2)%>"></td>
	<td bgcolor="#FFFFFF"><input type="checkbox" name="ExtCredits<%=U%>_3" value="1" <%If ExtSort(3)=1 or U<2 Then%>checked<%End If%> onclick="findobj('policy<%=U%>').disabled=!this.checked"></td>
	<td bgcolor="#F8F8F8"><input type="checkbox" name="ExtCredits<%=U%>_4" value="1" <%If ExtSort(4)=1 or U<2 Then%>checked<%End If%>></td>
</tr>
<% Next %>
</table></td></tr>
<tr><td colspan="2" bgcolor="#F8F8F8">
<table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
<tr class="a1"><td colspan="11">��չ������������</td></tr>
<tr align="center" class="a4">
	<td>���ִ���</td>
	<td>������(+)</td>
	<td>�ظ�(+)</td>
	<td>�Ӿ���(+)</td>
	<td>�ϴ�����(+)</td>
	<td>���ظ���(-)</td>
	<td>������Ϣ(-)</td>
	<td>����(-)</td>
	<td>�����ƹ�(+)</td>
	<td>���ֲ�������</td>
</tr>
<%
MustOpen = Split(team.Club_Class(22),"|")
For M=0 to Ubound(MustOpen)
	MustSort=Split(MustOpen(M),",")
%>
<tr align="center" id="policy<%=M%>" <%If Split(ExtCredits(M),",")(3)=0 Then%>disabled<%End If%>>
	<td bgcolor="#F8F8F8">Extcredits<%=M%></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_0" value="<%=MustSort(0)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_1" value="<%=MustSort(1)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_2" value="<%=MustSort(2)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_3" value="<%=MustSort(3)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_4" value="<%=MustSort(4)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_5" value="<%=MustSort(5)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_6" value="<%=MustSort(6)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_7" value="<%=MustSort(7)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_8" value="<%=MustSort(8)%>"></td>
</tr>
<%Next%>
<tr><td colspan="11" class="a4">&nbsp;</td></tr>
<tr>
	<td class="a3" align="center">������(+)</td><td class="a4" colspan="10">���߷����������ӵĻ���������������ⱻɾ�������߻���Ҳ�ᰴ�˱�׼��Ӧ����</td>
</tr>
<tr>
	<td class="a3" align="center">�ظ�(+)</td><td class="a4" colspan="10">���߷��»ظ����ӵĻ�����������ûظ���ɾ�������߻���Ҳ�ᰴ�˱�׼��Ӧ����</td>
</tr>
<tr>
	<td class="a3" align="center">�Ӿ���(+)</td><td class="a4" colspan="10">���ⱻ���뾫��ʱ�������ӵĻ���������������ⱻ�Ƴ����������߻���Ҳ�ᰴ�˱�׼��Ӧ����</td>
</tr>
<tr>
	<td class="a3" align="center">�ϴ�����(+)</td><td class="a4" colspan="10">�û�ÿ�ϴ�һ���������ӵĻ�����������ø�����ɾ���������߻���Ҳ�ᰴ�˱�׼��Ӧ����</td>
</tr>
<tr>
	<td class="a3" align="center">���ظ���(-)</td><td class="a4" colspan="10">�û�ÿ����һ�������۳��Ļ�������ע��: ��������ο������ظ����������Խ����ܱ��ƹ�</td>
</tr>
<tr>
	<td class="a3" align="center">������Ϣ(-)</td><td class="a4" colspan="10">�û�ÿ����һ������Ϣ�۳��Ļ�����</td>
</tr>
<tr>
	<td class="a3" align="center">����(-)</td><td class="a4" colspan="10">�û�ÿ����һ�����������۳��Ļ�����</td>
</tr>
<tr>
	<td class="a3" align="center">�����ƹ�(+)</td><td class="a4" colspan="10">������ͨ���û��ṩ���ƹ�����(�� ForumAdv.asp?Uid=1)������̳���ƹ������õĻ�����</td>
</tr>
<tr>
	<td class="a3" align="center">���ֲ�������</td><td class="a4" colspan="10">���û�������ֵ��ڴ�����ʱ������ֹ�û�ִ�л��ֲ������漰�ۼ�������ֵĲ�������������Ϊ -100�������������ۼ��û��� 10 ����λ�����û��������С�� -100 ʱ����������ִ�С�����������</td>
</tr>
<tr>
	<td class="a3" colspan="11">���ϱ���(+)��Ϊ���ӵĻ�����������(-)��Ϊ���ٵĻ���������Ҳ����ͨ�����ø�ֵ�ķ�ʽ������ֵ����������������������ķ�ΧΪ -99��+99�����Ϊ����Ĳ������û��ֲ��ԣ�ϵͳ����Ҫ��Ƶ���ĸ����û����֣�ͬʱ��ζ�����ĸ����ϵͳ��Դ����������ʵ�������������</td>
</tr></table></td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���׻�������:</b><br><span class="smalltxt">���׻�����һ�ֿ������û�������ת�á��������׵Ļ������ͣ�������ָ��һ�ֻ�����Ϊ���׻��֡������ָ�����׻��֣����û�����ֽ��׹��ܽ�����ʹ�á�ע��: ���׻��ֱ����������õĻ��֣�һ��ȷ���뾡����Ҫ���ģ�����������¼�����׿��ܻ�������⡣</span></td><td bgcolor="#FFFFFF">
		<select name="Forum_setting(99)">
			<option value="0">��</option>
			<%
			for i=1 to 7
				Response.Write "<option value="""&i&""" " 
				If Cid(team.Forum_setting(99)) = i Then Response.Write "selected"
				Response.Write ">extcredits"&i&"</option>"
			Next
			%>
		</select>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>������������:</b><br><span class="smalltxt">ָ��һֱ������������Ҫ�Ļ������ͣ������û������ֹ���ע��: ����ʹ�õĻ������Ա����������õĻ��֣�ʹ�ñ�����ѡ��Ŀ���ǽ���̳�Ľ��׻��������ֻ������ֿ����ý��׻��ֵĹ�����Ը�����</span></td><td bgcolor="#FFFFFF">
		<select name="Forum_setting(46)">
			<option value="0">��</option>
			<%
			for i=1 to 7
				Response.Write "<option value="""&i&""" " 
				If Cid(team.Forum_setting(46)) = i Then Response.Write "selected"
				Response.Write ">extcredits"&i&"</option>"
			Next
			%>
		</select>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���ֽ���˰:</b><br><span class="smalltxt">���ֽ���˰(��ʧ��)Ϊ�û������û��ֽ���ת�á��һ�������ʱ�۳���˰�ʣ���ΧΪ 0��1 ֮��ĸ���������������Ϊ 0.2�����û���ת�� 100 ����λ����ʱ����ʧ���Ļ���Ϊ 20 ����λ��0 Ϊ����ʧ</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(11)" value="<%=team.Forum_setting(11)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>ת��������:</b><br><span class="smalltxt">����ת�˺�Ҫ���û���ӵ�е������С��ֵ�����ô˹��ܣ����������ýϴ��������ƣ�ʹ����С�������ֵ���û��޷�ת�ˣ�Ҳ���Խ������������Ϊ������ʹ��ת�����޶��ڿ���͸֧</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(12)" value="<%=team.Forum_setting(12)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>��������߳���ʱ��(Сʱ):</b><br><span class="smalltxt">���õ����ⱻ���߳���ʱ��ϵͳ���������ⷢ��ʱ������ɳ��۵��ʱ�䡣������ʱ�����ƺ󽫱�Ϊ��ͨ���⣬�Ķ�������֧�����ֹ�������Ҳ�����ٻ����Ӧ���棬��СʱΪ��λ��0 Ϊ������</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(13)" value="<%=team.Forum_setting(13)%>"></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center><br>
<a name="��������"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">��������</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�տ�֧�����˺�:</b><br><span class="smalltxt">��������һ����׹��ܣ�����д��ʵ��Ч��֧�����˺ţ�������ȡ�û����ֽ�һ����׻��ֵ���ؿ�����˺���Ч��ȫ�����󣬽������û�֧�����޷���ȷ��������˻��Զ���ֵ������������Ľ��ס�</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(101)" value="<%=team.Forum_setting(101)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>֧������ȫУ����:</b><br><span class="smalltxt">��������һ����׹��ܣ�����д��ǰ��֧�����˻���ƥ��İ�ȫУ���룬�����Ѿ������˰�ȫ�룬���ڰ�ȫ���ǽ�ֻ��ʾ*�š������������¼֧���������̼ҹ����еġ���ȡ��ȫУ���롱�����û�鿴��ȫ������ݣ����˺���Ч��ȫ�����󣬽������û�֧�����޷���ȷ��������˻��Զ���ֵ������������Ľ��ס�</span></td><td bgcolor="#FFFFFF"><input type="password" size="30" name="Forum_setting(102)" value="<%=team.Forum_setting(102)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���������ID:</b><br><span class="smalltxt">��д��ID��֧�������׹��ܲ���ʹ�á����Ҫ��ȡ��ID������Ҫ��֧������վ������Ʒ���׷���</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(103)" value="<%=team.Forum_setting(103)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>�ֽ�/���ֶһ�����:</b><br><span class="smalltxt">�û�����Ļ�������ʵ����֮��ı���״��������Ϊ����1Ԫ����һ�����Ϊ1����1����֣�1����ң���һ�����Ϊ3����3����ֶһ�1����ҡ�</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(104)" value="<%=team.Forum_setting(104)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>���ֶһ�������:</b><br><span class="smalltxt">���ֶһ���Ҫ���û���ӵ�е������С��ֵ�����ô˹��ܣ����������ýϴ��������ƣ�ʹ����С�������ֵ���û��޷��һ���</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(105)" value="<%=team.Forum_setting(105)%>"></td>
</tr>

<tr class="a2"><td colspan="2"><img src="../images/dian.gif" align="absmiddle">  <a href="Admin_plus.asp?action=buyalipays" target="_blank"> �鿴ϵͳ���׶��� </a></td></tr>

<tr class="a2"><td colspan="2"><img src="../images/dian.gif" align="absmiddle"> �������: <a href="https://www.alipay.com/user/user_register.htm" target="_blank">ע��֧�����ʺ�</a> | <a href="https://www.alipay.com/" target="_blank">��½֧����</a> | <a href="http://help.alipay.com/support/index.htm/" target="_blank">֧�����ͷ�</a></td></tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="�� ��"></center><br>
</form>
<a name="ע��Э��"></a>
<form name="myform" method="post"  action="?action=upreg">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
	<td colspan="2">ע��Э��</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#F8F8F8">
		<textarea rows="10" name="myinfos" cols="60" style="overflow-y:visible;width:100%;"><%=server.htmlencode(team.Club_Class(13))%></textarea>
		<li> <B>[ �ı���֧�� UBB ] </B>
		<li> <B>{$clubname}����Ϊ��̳����</B>
	</td>
</tr>
</table><BR><center><input type="Submit" value="�� ��" name="Submit" />&nbsp;</center>
</form><br><br>
<%
End Sub

Sub SetColors(str)
	if Trim(team.Forum_setting(82))=Trim(""&str&"") Then
		Response.Write "selected"
	end if
End Sub


Sub Settingok	
	Dim ClubSystem,openclock,Co,u
	Dim ExtSort,MustSort,ExtCredits,MustOpen
	openclock=""
	For Co=0 to 23
		If openclock="" Then
			If Request.form("openclock"&Co)="1" Then
				openclock="1"
			Else
				openclock="0"
			End If
		Else
			If Request.form("openclock"&Co)="1" Then
				openclock=openclock&"*1"
			Else
				openclock=openclock&"*0"
			End If
		End If
	Next
	ClubSystem = ClubSystem & openclock &"$$$"
	If CID(Request.Form("Forum_setting(7)"))=1 Then
		If CID(Request.Form("Forum_setting(1)"))=0 Then SuccessMsg "������ѡ��Email���֧�֣��ſ���ʹ�ô˹��ܡ�"
	End If
	For i=1 to 120
		If i=8 Then ClubSystem= ClubSystem & team.Forum_setting(8)
		'If i = 87 Then ClubSystem= ClubSystem & Request.Cookies("userlogininfo")
		ClubSystem= ClubSystem & Replace(Request.Form("Forum_setting("&i&")"),"$$$","")&"$$$"
	Next
	For i=0 to 7
		ExtSort=""
		ExtSort= Request.Form("ExtCredits"&i&"_0")&","&Request.Form("ExtCredits"&i&"_1")&","&Request.Form("ExtCredits"&i&"_2")
		If Request.Form("ExtCredits"&i&"_3")=1 Then
			ExtSort= ExtSort& ",1"
		Else
			ExtSort= ExtSort& ",0"
		End If
		If Request.Form("ExtCredits"&i&"_4")=1 Then
			ExtSort= ExtSort& ",1"
		Else
			ExtSort= ExtSort& ",0"
		End If
		If ExtCredits="" Then
			ExtCredits = ExtCredits & ExtSort
		Else
			ExtCredits = ExtCredits & "|" &  ExtSort
		End If
	Next
	For i=0 to 7
		MustSort=""
		For u=0 to 9
			If MustSort="" Then
				If Request.Form("MustSort"&i&"_"&u&"")="" Then
					MustSort = "0"
				Else
					MustSort = Request.Form("MustSort"&i&"_"&u&"")
				End If
			Else
				If Request.Form("MustSort"&i&"_"&u&"")="" Then
					MustSort = MustSort &",0"
				Else
					MustSort = MustSort &","&Request.Form("MustSort"&i&"_"&u&"")
				End If
			End If
		Next
		If MustOpen="" Then
			MustOpen = MustOpen & MustSort
		Else
			MustOpen = MustOpen &"|"& MustSort
		End If
	Next
	If Len(Trim(Request.Form("Club_Class(8)")))>255 Then
		SuccessMsg "�����������������ѡ��ܶ���255���ַ�"
	End If
	If Len(Trim(Request.Form("Club_Class(28)")))>255 Then
		SuccessMsg "JS ��·���Ʋ��ܶ���255���ַ�"
	End If	
	If Len(Trim(Request.Form("Club_Class(7)")))>255 Then
		SuccessMsg "�û����ӹ������ò��ܶ���255���ַ�"
	End If	
	If Len(Trim(Request.Form("Club_Class(5)")))>255 Then
		SuccessMsg "�����������ò��ܶ���255���ַ�"
	End If	
	If Len(Trim(Request.Form("Club_Class(6)")))>255 Then
		SuccessMsg "�û�IP�������ò��ܶ���255���ַ�"
	End If
	team.Execute("Update "&IsForum&"Clubconfig set Allclass='"&Replace(ClubSystem,"'","")&"',ClubName='"&Replace(Trim(Request.Form("Club_Class(1)")),"'","")&"',Cluburl='"&Replace(Trim(Request.Form("Club_Class(2)")),"'","")&"',Homename='"&Replace(Trim(Request.Form("Club_Class(3)")),"'","")&"',Homeurl='"&Replace(Trim(Request.Form("Club_Class(4)")),"'","")&"',ExtCredits='"&ExtCredits&"',MustOpen='"&Replace(MustOpen,"'","")&"',ClearMail='"&Replace(Trim(Request.Form("Club_Class(23)")),"'","")&"',ClearIP='"&Replace(Trim(Request.Form("Club_Class(24)")),"'","")&"',UserKey='"&Replace(Trim(Request.Form("Club_Class(25)")),"'","")&"',BodyMeta='"&Replace(Trim(Request.Form("Club_Class(26)")),"'","")&"',ClearPost='"&Replace(Trim(Request.Form("Club_Class(27)")),"'","")&"',Badlist='"&Replace(Trim(Request.Form("Club_Class(7)")),"'","")&"',BadWords='"&Replace(Trim(Request.Form("Club_Class(5)")),"'","")&"',Badip='"&Replace(Trim(Request.Form("Club_Class(6)")),"'","")&"',JSUrl='"&Replace(Trim(Request.Form("Club_Class(28)")),"'","")&"',Starday='"&Replace(Trim(Request.Form("Club_Class(29)")),"'","")&"',ManageText='"&Replace(Trim(Request.Form("Club_Class(8)")),"'","")&"'")
	team.Execute("update ["&IsForum&"Style] Set StyleHid=0")
	team.Execute("update ["&IsForum&"Style] Set StyleHid=1 Where ID="& CID(request.Form("Forum_setting(18)")) )
	Cache.DelCache("Club_Class")
	team.SaveLog ("�������ø���")
	SuccessMsg "��̳�������ø��³ɹ�!"
End Sub
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
'�Զ������ļ���,��Ҫ�ƣӣ����֧�֡�
Private Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	'On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & "HTML"))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & "HTML")
		End If
		If Err.Number = 0 Then
			CreatePath = PathValue & "HTML" & "/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function
Call Footer()
%>