<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/MD5.asp" -->
<%
Dim x1,x2,fid
team.Headers(Team.Club_Class(1) &" -登陆后台管理")
If Not team.UserLoginED Then
	Response.Redirect "Login.asp"
ElseIf Not (team.IsMaster Or team.SuperMaster) Then 
	team.Error " 请查看您的权限,确认您有没有登陆后台的权限! "
End If
X1=" <a href=Admin.asp>登陆后台</a> "
x2=" "
Select Case Request("menu")
	Case "Logins"
		Call Logins()
	Case Else
		Call Main()
End Select
Sub Main()
	Dim MSCode
	If IsSqlDataBase = 1 Then
		MSCode="SQL"
	Else
		MSCode="ACC"
	End If
	Echo team.MenuTitle
%>
<form action="?menu=Logins" method="post">
<table cellpadding="5" cellspacing="1" border="0" align="center" class="a2" width="98%">
<tr>
   <td class="a1" colspan="2">TEAM5论坛管理登录 </td>
</tr>
<tr class="a4">
   <td colspan="2">
	 <b>论坛版本：<%=team.Forum_setting(8)%> - <%=MSCode%>版</b><BR />
     <b>发布网站：<a href="http://www.team5.cn" target="_blank">TEAM5.CN</a></a></b><BR />
   </td>
</tr>
<tr class="a4">
    <td width="40%" align="right"><b>用户名：</b></td>
    <td><INPUT name="username" type="text" value="<%=TK_UserName%>" onBlur="this.className='colorblur';" onfocus="this.className='colorfocus';this.value='';" class="colorblur"></td>
</tr>
<tr class="a4">
    <td align="right"><b>密　码：</b></font></td>
    <td><INPUT name="password" type="password" onBlur="this.className='colorblur';" onfocus="this.className='colorfocus';" class="colorblur"></td>
</tr>
<tr class="a4">
    <td align="right"><b>附加码：</b></td>
    <td><input size="20" name="code" onBlur="this.className='colorblur';" onfocus="this.className='colorfocus';" onclick="get_Code();" class="colorblur" > &nbsp;<span id="imgid" style="color:red">点击获取验证码</span></td>
</tr>
</table>
<BR><center><input class="button" type="submit" name="submit" value="登 录"></center>
</form>
<%
End Sub
team.footer

Sub  Logins
	Dim UserName,userpass,code,SQL,RS
	UserName = Replace(Trim(Request.Form("username")),"'","")
	Userpass=md5(Trim(Request.Form("password")),16)
	Code=Trim(Request.Form("code"))
	if Username=Empty or Userpass=Empty then 
		team.Error "用户名或密码没有填写,返回后请刷新登录页面后重新输入正确的信息。"
		Exit Sub
	End If
	if Not Team.CodeIsTrue(Code) Then
		team.Error "验证码错误,请刷新后重新填写. "
		'========================================
	End If
	SQL="select adminname,adminpass,AdminClass,forumname from ["&IsForum&"admin] where adminname='"&username&"'"
	Set Rs=Team.Execute(SQL)
	If RS.Eof and Rs.bof Then 
		session("Admin_Pass")=""
		team.Error("管理员名称错误")
		Exit Sub
	Else
		If Lcase(Trim(Username))<>Lcase(Trim(RS(0))) or Userpass<>RS(1)  Then 
			session("Admin_Pass")=""
			team.Error("管理员名称或密码错误")
			Exit Sub
		ElseIf Lcase(Trim(TK_UserName))<>Lcase(Trim(RS(3))) Then
			session("Admin_Pass")=""
			team.Error("此用户没有绑定后台用户名")
			Exit Sub
		Else
			session("Admin_Pass")=RS(2)
			Session("UserMember")= team.UserGroupID
			Team.execute("Update ["&IsForum&"admin] set Loginip='"&replace(Request.ServerVariables("REMOTE_ADDR"),"'","")&"',Logintime="&SqlNowString&" where adminname='"&username&"'")
			Response.Redirect ManagePath &"index.asp"
		End If
	End If
	RS.Close:Set RS=Nothing
End Sub
%>


