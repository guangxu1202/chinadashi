<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<!-- #include file="../inc/MD5.asp" -->
<%
Call Master_Us()
Header()
Dim Admin_Class
Call Master_Se()
Select Case Request("menu")
	Case "adminlogin"
		adminlogin
	Case "leftbody"
		leftbody
	Case "pass"
		pass
	Case "topbanner"
		topbanner
	Case "out"
		session("Admin_Pass")=""
		Session("UserMember")=""
		Error1("������̳��ҳ��,���Ե�...<meta http-equiv=refresh content=3;url=default.asp>")
	Case Else
		Call ManageIndex
End Select

Sub topbanner
	Dim MSCode
	If IsSqlDataBase = 1 Then
		MSCode="SQL"
	Else
		MSCode="ACC"
	End If
%>
<body topmargin="0" rightmargin="0" leftmargin="0">
<table border="0" cellpadding="3" cellspacing="0" width="100%" class="a2">
	<tr class="a1" align="center">
		<td><b>��̳�汾��<%=team.Forum_setting(8)%> - <%=MSCode%></b></td>
		<td><a href="http://www.team5.cn" target="_blank">���ʹٷ���̳</a></td>
		<td><a href="../" target="_blank">վ����ҳ </a></td>
	</tr>
</table>
<%
End Sub


Sub pass
	Dim Admin_Class,datapath,datafile,mylockinfo
	Dim Members,Thread,ldb,lConnStr,lConn
	Members = team.execute("Select count(id) from ["&IsForum&"User]")(0)
	Thread = team.execute("Select count(id) from ["&IsForum&"Forum]")(0)
	team.SaveLog ("��½��̨����")
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="5" topmargin="5">
<table cellpadding="3" cellspacing="1" border="0" width="90%" class="a2">
<tr>
	<td class="a1" colspan="2" height="25"><b>TEAM5.CN �ٷ���̬</b></td>
</tr>
<tr>
	<td class="a4" colspan="2">
    <script src="http://server.team5.cn/GetNews.asp?version=2.0.4&release=<%=team.iBuild%>&bbsname=<%=team.Club_Class(1)%>&members=<%=Members%>&threads=<%=Thread%>&posts=<%=Application(CacheName&"_ConverPostNum")%>&urls=<%=Request.ServerVariables("server_name")%>&tmaster=<%=HtmlEncode(tk_UserName)%>"></script>
	</td>
</tr>
</table>
<br>
<table cellpadding="3" cellspacing="1" border="0" width="90%" class="a2">
<tr><td class="a1"><b>ϵͳ��Ϣ</b></td></tr>
<tr><td class="a4">
<table cellpadding="3" cellspacing="0" border="0" width="100%">
	<tr class="a4">
		<td>��̳������ͳ��: </td><td><%=Thread%></td>
	</tr>
	<tr class="a3">
		<td>��ǰ����������: </td><td>[ <%=team.Club_Class(11)%> ]</td>
	</tr>		
	<tr class="a4">
		<td>��ǰ������ͳ��: </td><td><%=team.execute("Select count(id)from ["&team.Club_Class(11)&"]")(0)%></td>
	</tr>
	<tr class="a3">
		<td>��ǰ�û���ͳ��: </td><td><%=Members%></td>
	</tr>	

	<tr class="a4">
		<td>��ǰ����������: </td><td><%=Request.ServerVariables("server_name")%> 
		/ <%=Request.ServerVariables("LOCAL_ADDR")%></td>
	</tr>
	<tr class="a3">
		<td>��ǰ�ű���������: </td><td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="a4">
		<td>��ǰIIS�汾ϵͳ: </td><td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	</tr>
	<tr class="a3">
		<td>��ǰ���ݿ�·��: </td><td><%=db%></td>
	</tr>
</table>
</td></tr></table>
<BR />
<table cellpadding="5" cellspacing="1" border="0" width="90%" class="a2">
	<tr><td class="a1"align="center"><b>�����ݷ�ʽ</b></td></tr>
	<tr><td class=a4>

	<table cellpadding="3" cellspacing="0" border="0" width="100%">
	<tr class="a3">
		<td>��̳����״̬: </td>
		<td>
		<% If team.Forum_setting(2)=1 Then%><font color="blue">�ر�</font><%Else%><font color="red">��������</font><%End If%>
	</tr>
	<tr class="a4">
		<td>����������̳: </td>
		<td>
		<a href="admin_update.asp?action=UP_clear">ȷ������</a>
	</tr>
	<tr class="a3">
		<td>���ٲ����û�</td>
		<td>
		<form method="post" action="Admin_User.asp?action=finduser">
			<input size="30" name="username"> <input type="submit" value="���̲���">
		</td></form>
	</tr>
	<tr class="a3">
		<td>����û�</td>
		<td> Ŀǰ��<%=CID(team.execute("Select Count(*) From ["& IsForum &"User] Where UserGroupID=5 ")(0))%> ���û��ȴ�����ˡ�[<a href="Admin_User.asp?action=Activation"><FONT COLOR="blue">�鿴��ϸ</FONT></a>]</td>
	</tr>
	<tr class="a4">
		<td>�鿴ϵͳ���׶���</td>
		<td>���� <%=CID(team.execute ("Select Count(ID) From ["&Isforum&"BankLog] Where Makes = 0")(0))%> �����׶���δ����[<a href="Admin_plus.asp?action=buyalipays"><FONT COLOR="blue">�鿴��ϸ</FONT></a>]
		</td>
	</tr>
	</table>
</td></tr>
</table>

<BR />
<table cellpadding="3" cellspacing="1" border="0" width="90%" class="a2">
<tr><td class=a1>TEAM��̳�����ĵ�˵��</td></tr>
<tr><td class=a4>
<table cellpadding="3" cellspacing="0" border="0" width="100%">
<tr class=a4><td width=30%>��Ȩ����:</td> <td>Team Studio</td></tr>
<tr class=a3><td>���򿪷�:</td> <td>DayMoon,�Ķ�����</td></tr>
<tr class=a4><td>�ٷ���̳:</td> <td>http://www.team5.cn</td></tr>
<tr class=a3><td>�ʼ���ַ:</td> <td>teamserver@163.com</td></tr>
</table>
</td></tr></table>
<%
	ldb = MyDbPath & LogDate
	lConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ldb)
	Set lConn = Server.CreateObject("ADODB.Connection")
	lConn.Open lConnStr
	lConn.execute("Delete from [SaveLog] Where Datediff('d',LogTime,'"&Now()&"') > "& team.Forum_setting(47) &" * 30 ")
	Footer()
End sub


Sub leftbody
	Dim Admin_Class
%>
<script language="JavaScript">
function ClearAllDeploy(){
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var userdeploy='';
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		userdeploy=deployitem.substring(0,admin_start);
	}
	for(i=0;i<20;i++){
		obj=document.getElementById("cate_"+"id"+i);	
		img=document.getElementById("img_"+"id"+i);
		if(obj && obj.style.display=="none"){
			obj.style.display="";
			img_re=new RegExp("_open\\.gif$");
			img.src=img.src.replace(img_re,'_fold.gif');
		}
	}
	deployitem=userdeploy+"\n\t\t";
	SetCookie("deploy",deployitem);
}
function SetAllDeploy(){
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var userdeploy='';
	var admindeploy='';
	var i;
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		userdeploy=deployitem.substring(0,admin_start);
	}
	for(i=0;i<20;i++){
		obj=document.getElementById("cate_"+"id"+i);	
		img=document.getElementById("img_"+"id"+i);
		if(obj && obj.style.display==""){
			obj.style.display="none";
			img_re=new RegExp("_fold\\.gif$");
			img.src=img.src.replace(img_re,'_open.gif');
		}
		admindeploy=admindeploy+"id"+i+"\t";
	}
	deployitem=userdeploy+"\n\t"+admindeploy;
	SetCookie("deploy",deployitem);
}
function IndexDeploy(ID,type){
	obj=document.getElementById("cate_"+ID);	
	img=document.getElementById("img_"+ID);
	if(obj.style.display=="none"){
		obj.style.display="";
		img_re=new RegExp("_open\\.gif$");
		img.src=img.src.replace(img_re,'_fold.gif');
		SaveDeploy(ID,type,false);
	}else{
		obj.style.display="none";
		img_re=new RegExp("_fold\\.gif$");
		img.src=img.src.replace(img_re,'_open.gif');
		SaveDeploy(ID,type,true);
	}
	return false;
}
function SaveDeploy(ID,type,is){
	var foo=new Array();
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var admindeploy='';
	var userdeploy='';
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		admindeploy= deployitem.substring(admin_start+1,deployitem.length);
		userdeploy = deployitem.substring(0,admin_start);
	}
	if(deployitem!=null){
		if(admin_start!=-1){
			deployitem = type==0 ? userdeploy : admindeploy;
		}
		deployitem=deployitem.split("\t");
		for(i in deployitem){
			if(deployitem[i]!=ID && deployitem[i]!=""){
				foo[foo.length]=deployitem[i];
			}
		}
	}
	if(is){
		foo[foo.length]=ID;
	}
	deployitem = type==0 ? "\t"+foo.join("\t")+"\t\n"+admindeploy : userdeploy+"\n\t"+foo.join("\t")+"\t";
	SetCookie("deploy",deployitem)
}
function SetCookie(name,value){
	expires=new Date();
	expires.setTime(expires.getTime()+(86400*365));
	document.cookie=name+"="+escape(value)+"; expires="+expires.toGMTString()+"; path=/";
}
function FetchCookie(name){
	var start=document.cookie.indexOf(name);
	var end=document.cookie.indexOf(";",start);
	return start==-1 ? null : unescape(document.cookie.substring(start+name.length+1,(end>start ? end : document.cookie.length)));
}
</script>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="5" topmargin="5">
<table width="100%" cellspacing="1" cellpadding="4" border="0" class=a4>
	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
        <tr>
          <td class="a4"><a href="#" onClick="return ClearAllDeploy()" class="a_bold">[չ��+]</a>&nbsp;&nbsp;<a href="#" onClick="return SetAllDeploy()" class="a_bold">[�ر�-]</a> </td>
        </tr>
      </table>
	</td></tr>

	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
		<a style="float:right"><img src="images/cate_fold.gif" border="0"></a>
			<a href="?menu=pass" class="a1" target="main"><b>������ҳ</b></a>
		</td></tr>
		</table>
	</td></tr>
	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id0',1)"><img id="img_id0" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id0',1)" class="a1"><b>����ѡ��</b></a>
		</td></tr>
		<tbody id="cate_id0" style="display:none;">
		  <tr>
            <td class="a4"><a href="Admincp.asp#��������" target="main">����������</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#ע������ʿ���" target="main">��ע������ʿ���</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#��������ʾ��ʽ" target="main">����������ʾ��ʽ</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#���������Ż�" target="main">�����������Ż�</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#��̳����" target="main">����̳����</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#��ȫ����" target="main">����ȫ����</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#ʱ��μ���������" target="main">��ʱ�估��������</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#�û�Ȩ��" target="main">���û�Ȩ��</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#��������" target="main">����������</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#JS ����" target="main">��JS ����</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#��������" target="main">����������</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#��������" target="main">����������</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admincp.asp#��������" target="main">����������</a> </td>
          </tr>
		</tbody>
		</table>
	</td></tr>
	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id1',1)"><img id="img_id1" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id1',1)" class="a1"><b>��̳����</b></a>
		</td></tr>
		<tbody id="cate_id1" style="display:none;">
         <tr>
            <td class="a4"><a href="Admin_Forum.asp" target="main">���༭���</a> </td>
          </tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Manage.asp" target="main">����ݹ���</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Manage.asp?Action=readkey" target="main">���������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Update.asp" target="main">��������̳ͳ��</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>

	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id2',1)"><img id="img_id2" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id2',1)" class="a1"><b>�����뼶��</b></a>
		</td></tr>
		<tbody id="cate_id2" style="display:none;">
          <tr>
            <td class="a4"><a href="Admin_Group.asp" target="main">��������</a> </td>
          </tr>
          <tr>
            <td class="a4"><a href="Admin_Group.asp?Action=IsuserGroup" target="main">���û���</a> </td>
          </tr>
		</tbody>
		</table>
	</td></tr>

	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id3',1)"><img id="img_id3" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id3',1)" class="a1"><b>�û�����</b></a>
		</td></tr>
		<tbody id="cate_id3" style="display:none;">
		<tr>
		  <td class="a4">
		  <a href="Admin_User.asp" target="main">���༭�û�</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_User.asp?action=adduser" target="main">������û�</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_User.asp?action=setuser" target="main">���ϲ��û�</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_User.asp?action=Activation" target="main">������û�</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_User.asp?action=getmoney" target="main">�����ʹ���</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>

	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id4',1)"><img id="img_id4" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id4',1)" class="a1"><b>������</b></a>
		</td></tr>
		<tbody id="cate_id4" style="display:none;">
		<tr>
		  <td class="a4">
		  <a href="admin_skins.asp" target="main">���༭ģ��</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="admin_skins.asp?menu=loading" target="main">��ģ�嵼��</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="admin_skins.asp?menu=output" target="main">��ģ�嵼��</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>
	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id5',1)"><img id="img_id5" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id5',1)" class="a1"><b>��������</b></a>
		</td></tr>
		<tbody id="cate_id5" style="display:none;">	
		<tr>
		  <td class="a4">
		  <a href="Admin_Change.asp?action=announcements" target="main">����̳����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Change.asp?action=forumlinks" target="main">����������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Change.asp?action=medals" target="main">��ѫ�±༭</a>
		  </td>
		</tr>	
		<tr>
		  <td class="a4">
		  <a href="Admin_Change.asp?action=adv" target="main">��������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Change.asp?action=onlinelist" target="main">�������б���</a>
		  </td>
		</tr>		
		</tbody>
		</table>
	</td></tr>

	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id6',1)"><img id="img_id6" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id6',1)" class="a1"><b>�������</b></a>
		</td></tr>
		<tbody id="cate_id6" style="display:none;">
		<tr>
		  <td class="a4">
		  <a href="Admin_plus.asp" target="main">���˵�����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_plus.asp?action=makeonline" target="main">������������Ա</a>
		  </td>
		</tr>
		<%
		Dim Rs
		Set Rs=team.Execute("Select Name,url From "&IsForum&"Menu Where Newtype=0 Order By SortNum")
		Do While not Rs.Eof
			Echo "<tr><td class=""a4"">"
			Echo " <a href="""&Rs(1)&""" target=""main"">��"&Rs(0)&"</a> "
			Echo "</td></tr>"
			Rs.MoveNext
		Loop
		Rs.close:Set Rs=Nothing
		%>
		</tbody>
		</table>
	</td></tr>
	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id7',1)"><img id="img_id7" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id7',1)" class="a1"><b>��̳ά��</b></a>
		</td></tr>
		<tbody id="cate_id7" style="display:none;">
		<tr>
		  <td class="a4">
		  <a href="Admin_dbmake.asp" target="main">�����ݿ����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_dbmake.asp?action=updates" target="main">�����ݿ�����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_dbmake.asp?action=reforums" target="main">������������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_dbmake.asp?action=upfiles" target="main">����������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_dbmake.asp?action=clearmsg" target="main">�����Ź���</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_dbmake.asp?action=savelog" target="main">��������¼</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>
	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id8',1)"><img id="img_id8" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id8',1)" class="a1"><b>ͳ����Ϣ</b></a>
		</td></tr>
		<tbody id="cate_id8" style="display:none;">
		<tr>
		  <td class="a4">
		  <a href="Admin_Path.asp" target="main">��������������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Path.asp?action=discreteness" target="main">�����֧�����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_Path.asp?action=statroom" target="main">��ͳ��ռ�ÿռ�</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>

	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id9',1)"><img id="img_id9" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id9',1)" class="a1"><b>��̨Ȩ��</b></a>
		</td></tr>
		<tbody id="cate_id9" style="display:none;">
		<tr>
		  <td class="a4">
		  <a href="Admin_maste.asp" target="main">������Ա���</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a href="Admin_maste.asp?action=manages" target="main">������Ȩ������</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>

	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onclick="return IndexDeploy('id10',1)"><img id="img_id10" src="images/cate_fold.gif" border="0"></a>
			<a href="#" onclick="return IndexDeploy('id10',1)" class="a1"><b>��������</b></a>
		</td></tr>
		<tbody id="cate_id10" style="display:">
		<tr>
		  <td class="a4">
		  <a href="../" target="_blank">����̳��ҳ</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4">
		  <a target="_top" href="?menu=out">���˳�����</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>
	<tr>
    <td height="2"></td>
  </tr>
  <tr>
    <td height="2"><a href="http://www.TEAM5.CN" target="_blank" class="a_bold">&copy;&nbsp;TEAM's�ٷ���̳</a></td>
  </tr>
</table>
<%
end sub


Sub ManageIndex
	%>
	<frameset cols="160,*" frameborder="no" border="0" framespacing="0" rows="*">
	<frame name="menu" noresize scrolling="yes" src="?menu=leftbody">
	<frameset rows="25,*" frameborder="no" border="0" framespacing="0" cols="*">
	<frame name="a1" noresize scrolling="no" src="?menu=topbanner">
	<frame name="main" noresize scrolling="yes" src="?menu=pass">
	</frameset></frameset></html>
	<%
End  Sub 
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>