<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<!-- #include file="../inc/MD5.asp" -->
<%
Public boards
Dim Admin_Class,Uid
Call Master_Us()

Uid = Cid(Request("uid"))
If Cid(Session("UserMember")) <> 1 Then 
	SuccessMsg "对不起，只有管理员方可查看此版的内容 。 "
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
		SuccessMsg "参数错误。"
	Else
		team.Execute("Delete From ["&isforum&"Admin] Where ID="&UID)
		SuccessMsg "此后台登陆用户已经被删除。"
	End If
End Sub

Sub upkeyok
	If Uid="" or Not IsNumeric(Uid) Then
		SuccessMsg "参数错误。"
	Else
		team.Execute("Update ["&isforum&"Admin] Set adminname='"&request.Form("adminname")&"',forumname='"&request.Form("forumname")&"',adminpass='"&Md5(request.Form("adminpass"),16)&"' Where ID="&UID)
		SuccessMsg "后台管理密码修改完成。"
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
      <td colspan="3">修改后台管理员密码</td>
    </tr>
    <tr class="a3">
      <td align="center" width="40%">后台登陆名称：</td>
      <td width="30%"><input name="adminname" size="30" value="<%=RS(1)%>"></td>
      <td width="30%">(可与注册名不同) </td>
    </tr>
    <tr class="a4">
      <td align="center">后台登陆密码：</td>
      <td><input name="adminpass" size="30"></td>
      <td>(可与注册密码不同)</td>
    </tr>
    <tr class="a3">
      <td align="center">前台用户名称：</td>
      <td><input name="forumname" size="30" value="<%=RS(2)%>"></td>
      <td>一个前台用户名可绑定多个后台名称登陆!</td>
    </tr>
  </table>
  <br>
  <center>
  <input type=submit value="确定">
  <center>
</form>
<%
	End if
End Sub

Sub edmasterok	
	Dim Admin_s
	Admin_s = Replace(Request.Form("Admin_Pass")," ","")
	If Uid="" or Not IsNumeric(Uid) Then
		SuccessMsg "参数错误。"
	Else
		team.execute("Update ["&isforum&"admin] set adminclass='"&Admin_s&"' Where ID="&UID)
		SuccessMsg "后台权限设置完成。"
	End If
End Sub

Sub edmaster
	dim menu(4,3),trs,k
	menu(0,0)=" 论坛后台管理权限分配 "
	menu(0,1)="<a href=Admincp.asp> 基本选项 </a> [包括：所有项目 ]@@1"
	menu(1,1)="<a href=Admin_Forum.asp> 论坛设置  </a> [包括：编辑版块  ]@@2"
	menu(1,2)="<a href=Admin_Manage.asp> 快捷管理  </a> [包括：快捷管理，帖子审核]@@3"
	menu(1,3)="<a href=Admin_Update.asp> 更新论坛统计  </a> [包括：更新论坛统计 ]@@4"
	menu(2,1)="<a href=Admin_Group.asp> 分组与级别   </a> [包括：管理组 ，用户组 ]@@5"
	menu(2,2)="<a href=Admin_User.asp> 用户管理  </a> [包括：编辑用户，添加用户 ，合并用户 ，审核用户 ，工资管理 ]@@6"
	menu(3,1)="<a href=admin_skins.asp> 界面风格 </a> [包括：编辑模板 ，模板导入，模板导出 ]@@7"
	menu(3,2)="<a href=Admin_Change.asp>其他设置  </a> [包括：论坛公告 ，友情链接 ，勋章编辑 ，广告管理 ，在线列表定制 ]@@8"
	menu(3,3)="<a href=Admin_plus.asp>插件设置 </a> [包括：菜单管理 ，虚拟在线人员 ]@@9"
	menu(4,1)="<a href=Admin_dbmake.asp> 论坛维护 </a> [包括：数据库管理 ，数据库升级 ，回帖表设置 ，附件管理 ，短信管理 ，操作记录 ]@@10"
	menu(4,2)="<a href=Admin_Path.asp> 统计信息  </a> [包括：主机环境变量 ，组件支持情况，统计占用空间 ]@@11"

	dim j,tmpmenu,menuname,menurl,rs
	Set Rs=Team.Execute("Select forumname,adminclass From "&IsForum&"Admin Where ID="&UID )	
	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>TEAM系统可以进入后台管理的等级组默认是 <B>管理员</B> 和 超级版主 ，管理员拥有所有权限， 而超级版主拥有除了对后台管理用户的添加删除等管理功能之外的所有权限，但是其他权限必须在<FONT COLOR="red">权限给予</FONT>的情况下才可以执行。所以建议对超级版主给予一定的后台权限，以降低管理员的工作强度 。
      </ul></td>
  </tr>
</table>
<br>
<form action="?action=edmasterok" method="post">
  <table cellpadding="3" cellspacing="1" border="0" width="95%" class="a2" align="center">
    <tr class="a1">
      <td><b>管理员权限管理</b>( 请选择相应的权限分配给管理员 <%=rs("forumname")%> )</td>
    </tr>
    <tr>
      <td class="a3"><table cellpadding="3" cellspacing="1" border="0" width="95%" class="a2" align="center">
          <tr>
            <td class="a1">全局权限</td>
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
              选择所有权限</td>
          </tr>
        </table>
        <BR>
        <center>
          <input type="submit" name="Submit" value="更新">
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
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>你可以在此进行后台管理员的添加和后台密码的更改，查看各管理员的登陆情况，修改管理员的权限等等。
      </ul>
      <ul>
        <li>此版的功能只有拥有管理员等级的用户方可登陆管理。
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=masterupdate">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
<tr align="center" class="a1">
  <td colspan="7">查看后台管理登陆状况 </td>
</tr>
<tr class="a3" align="center">
  <td>登陆用户</td>
  <td>前台用户名</td>
  <td>登陆IP</td>
  <td>最后登陆时间</td>
  <td>权限</td>
  <td>密码</td>
  <td>管理</td>
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
		Echo "	<td> <a href=""?action=edmaster&uid="&RS(0)&""">编辑权限</a> </td>"
		Echo "	<td> <a href=""?action=upkey&uid="&RS(0)&""">修改密码</a> "
		Echo "	<td> <a href=""?action=killmaster&uid="&RS(0)&""">删除?</a></td>	"
		Echo "</tr>"
		RS.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo " </table><br><center><input type=""submit"" name=""detailsubmit"" value=""提 交""></form><br>"
End Sub



Sub masterupdate
	Dim Rs,adminname,adminpass,forumname,Us
	Adminname = Replace(Request.Form("Adminname"),"'","''")
	Adminpass = Replace(Request.Form("adminpass"),"'","''")
	Forumname = Replace(Request.Form("forumname"),"'","''")
	Set Us = team.Execute("Select UserGroupID from ["&Isforum&"User] Where UserName='"&Forumname&"'")
	If Us.Eof And Us.bof Then
		SuccessMsg "此前台用户名不存在,请重新设置。"
	Else
		Set Rs= team.execute("Select adminname,forumname from ["&Isforum&"admin] ")
		Do While Not Rs.Eof
			If LCase(Rs(0))=Lcase(Adminname) Then 
				SuccessMsg " 此后台用户名已经存在,请重新设置! 如果您需要修改密码,请使用修改密码的功能 。"
			End If
			If LCase(Rs(1))=Lcase(Forumname) Then 
				SuccessMsg " 此用户名已经建立了后台管理用户名称 。"
			End if
			Rs.Movenext
		Loop
		Rs.Close:Set Rs=Nothing
		If Len(Adminpass)< 6  Then 
			SuccessMsg " 管理密码不能少于6位数 。"
		Else
			team.Execute( "insert into "&Isforum&"admin (adminname,adminpass,forumname) values ('"&Adminname&"','"&MD5(Adminpass,16)&"','"&Forumname&"')" )
			If Int(Us(0)) >2 Then
				team.Execute("Update ["&Isforum&"User] Set UserGroupID=2,LevelName='超级版主||||||18||0',Members='超级版主' Where UserName='"&Forumname&"'")
			End if
			SuccessMsg " 后台管理员添加成功。<BR> 默认添加的用户属于超级版主组，如果您需要将此用户加入管理员组,请 转入<A HREF=""Admin_User.asp""><B>编辑用户 </B></A>选项，或者请等待系统自动返回到 <a href=Admin_maste.asp?action=manages>管理权限设置  </a> 页面，请选择 <B>编辑权限</B> 选项，对后台用户进行后台管理权限编辑 。<meta http-equiv=refresh content=3;url=Admin_maste.asp?action=manages>。 "
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
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>设置此选项，可以给予前台用户登入后台管理的权限，请谨慎使用此权限。
      </ul>
      <ul>
        <li>前台用户如果绑定为后台管理员，将自动加入管理组的超级版主组，拥有超版的所有权限。
      </ul>
      <ul>
        <li>初次设置后台管理用户，请先用默认用户名 admin 登陆到后台，然后按照一下步骤执行：
          <ul>
            <li> 1. <A HREF="Admin_User.asp?action=adduser"><B>添加新的用户</B></A> </li>
          </ul>
          <ul>
            <li> 2. 然后转入<A HREF="Admin_User.asp"><B>编辑用户 </B></A>选项，搜索此用户，搜索完成后点击 <B>用户属性</B>，将 <B>用户所属组类别</B> 设置为 <B>管理员</B> 。</li>
          </ul>
          <ul>
            <li> 3. 转入 <A HREF="Admin_maste.asp"><B>管理员添加 </B></A>选项,添加新的后台管理用户名称，将<B>前台用户名称</B>设置为刚才添加的用户名称。</li>
          </ul>
          <ul>
            <li> 4. 最后转入 <A HREF="Admin_maste.asp?action=manages"><B>管理权限设置 </B></A>选项,给新添加的后台用户设置详细的后台管理权限。</li>
          </ul>
        </li>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=masterupdate">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
      <td colspan="2">添加后台管理员</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="60%"><B>后台登陆名称：</B><BR>
        管理员登陆后台时候使用的登陆名，此登陆名称与前台的注册名可以不同，但是不能与已经存在的后台用户名称重复，每个管理者都拥有独立的后台登陆名称，此用户名仅在登陆后台管理时有效。 </td>
      <td bgcolor="#F8F8F8"><input name="adminname" size="30" value="<%=TK_UserName%>">
      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="60%"><B>后台登陆密码：</B><BR>
        管理员登陆到后台需要输入的密码，每个后台管理者都拥有一个独立的管理密码，管理员如果需要登陆到后台管理，必需输入正确密码才可以登入。 </td>
      <td bgcolor="#F8F8F8"><input name="adminpass" size="30" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="60%"><B>前台用户名称：</B><BR>
        每个后台管理名称都对应一个前台的管理员名称， 如 默认名称 <FONT COLOR="RED">admin</FONT> 其对应的用户为管理员 <FONT COLOR="RED">admin</FONT> ，此处输入的名称必须为前台的用户，如果此用户不存在，请先在 <A HREF="Admin_User.asp?action=adduser"><B>添加用户</B></A> 选项添加新的用户，然后再设置绑定名称。 </td>
      <td bgcolor="#F8F8F8"><input name="forumname" size="30" value="">
      </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="提 交">
</form>
<br>
<%
End Sub

footer()
%>
