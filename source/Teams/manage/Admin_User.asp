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
team.SaveLog ("用户管理 [包括：编辑用户，添加用户 ，合并用户 ，审核用户 ，工资管理 ]  ")
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
		If NewMembers =""  Then SuccessMsg " 组名称不能为空。"
		If team.execute("Select Members from ["&Isforum&"Wages] where Members='"&NewMembers&"' ").Eof Then
			Set Gs = team.execute("Select ID,GroupName From ["&IsForum&"UserGroup] Where ID = "& NewMembers)
			If Not Gs.Eof Then
				team.execute("insert into ["&Isforum&"Wages] (Members,WageMach,WageGroupID) values ('"&Gs(1)&"',"&NewWageMach&","&Gs(0)&") ")
			End if
			Gs.Close:Set Gs=Nothing
		Else
			SuccessMsg  " 此用户组已经存在! "
		End If
	End if
	SuccessMsg  " 工资图表设置完成。 "
End Sub

Sub getmoney 
	%>
	<br>
	<br>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?action=getmoneyok">
	<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr class="a1">
	 <td>技巧提示</td>
	</tr>
	<tr class="a4">
	 <td><br>
      <ul>
        <li>TEAM's 支持对各用户组发放工资，此功能需要在基本选项里面开启交易积分设置 。
		<li>打开此功能后，系统将在每月的第一天自动发放工资（工资名称为交易积分设置值）到用户账户 。
      </ul></td>
	 </tr>
	</table>
	<BR>
	<table cellspacing="1" cellpadding="3" border="0" width="95%" align="center" class="a2">
	<tr class="a1">
		<td align="center" colspan="3">本月工资额度管理</td>
	</tr>
	<tr class="a3">
		<td align="center" width="80"><input type="checkbox" name="chkall" onclick="checkall(this.form, 'wagid')" class="a3">删? </td>
		<td align="center">组对象</td><td align="center">工资额度</td>
	</tr>
	<tbody <%if team.Forum_setting(96)=0 Then%>disabled<%end if%>>
	<%
	Dim Rs,WsValue,m
	i = 0
	Set Rs= team.execute("Select ID,Members,WageMach,WageGroupID From ["&Isforum&"Wages]")
	If Rs.Eof Then
		Echo "<tr class=""a4"" align=""center""><td colspan=""3"">目前没有需要的发放工资的组对象</td></tr>"
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
		<td>新增:</td>
		<td><select name="newMembers" style="width:100%">
		<option value=""> 请选择用户组 </option>
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
	<input type="submit" name="medalsubmit" value="提 交" <%if team.Forum_setting(96)=0 Then%>disabled<%end if%>>
	</center></form>
<%
End Sub

Sub Activationok
	Dim Ho
	for each ho in request.form("checkuid")
		team.execute("Update ["&Isforum&"User] Set UserGroupID=27,Levelname='附小一年级||||||0||0',Members='注册用户' Where ID="&ho)
	Next
	team.SaveLog ("审核用户操作")
	SuccessMsg " 审核用户操作成功，请等待系统自动返回到 <a href=Admin_User.asp?action=Activation>审核用户  </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_User.asp?action=Activation>。 "
End Sub

Sub Activation %>
	<br>
	<br>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?action=activationok">
	<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr class="a1">
	 <td>技巧提示</td>
	</tr>
	<tr class="a4">
	 <td><br>
      <ul>
        <li>TEAM's 提供了3种新用户注册验证方式，点击查看 <a href="http://localhost/1/Manage/Admincp.asp#注册与访问控制"> 注册与访问控制 。</a>
		<li>如何您设置了人工审核，那么就需要手动将用户进行激活 。
      </ul></td>
	 </tr>
	</table>
	<BR>
	<table cellspacing="1" cellpadding="3" border="0" width="95%" align="center" class="a2">
	<tr class="tab1">
		<td width="15%"><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio"> 审核 </td><td>待审核用户列表</td><td>目前等级</td>
	</tr>
	<%
	Dim Rs
	Set Rs = team.execute("Select * From ["&IsForum&"User] Where UserGroupID=5")
	Do While Not Rs.Eof
		Echo "<tr class=""a4"">"
		Echo "	<td align=""center""><input type=""checkbox"" name=""checkuid"" value="""&RS("ID")&""" class=""radio""></td><td title=""点击可查看用户的详细资料""><a href=""Admin_User.asp?action=editgroups&uid="&RS("ID")&""">"&Rs("UserName")&"</a></td><td>"&Split(Rs("Levelname"),"||")(0)&"</td>"
		Echo "</tr>"
		Rs.MoveNext
	Loop
	Echo "</table></table><br><center><input type=""submit"" name=""onlinesubmit"" value=""提 交""></center></form>"
	Rs.Close:Set Rs=Nothing
End Sub

Sub editmedalsok
	Dim tmp,newid
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " 参数错误 "
	Else
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			If CID(Request.Form("medals"&i))>0 Then
				If Request.Form("reason"&i)&"" = "" Then 
					SuccessMsg "您必须输入授予理由。"
				End if
				tmp = tmp & Request.Form("medalname"&i) &"&&&"&Request.Form("reason"&i) & "$$$"
			End if
		Next
		team.Execute("Update ["&Isforum&"User] set Medals='"&tmp&"' Where ID= "& Uid )
	end if
	SuccessMsg " 用户的勋章设置完成。 "
End Sub

Sub editmedals
	Dim Rs,Rs1,Medals,u,i,Medalsinfo,Medaltext,MyMedals
	Dim SetMys,SetMsg
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " 参数错误 "
	Else
		Set Rs = team.execute("Select UserName,Medals From ["&Isforum&"User] Where ID="& Uid)
		If Rs.Eof Then
			SuccessMsg " 指定的用户不存在。 "
		Else
			Set Rs1=team.execute("Select ID,MedalName,Medalimg From ["&Isforum&"Medals] Where MedalSet=1 Order By ID asc")
			If Rs1.Eof Then
				SuccessMsg " 目前没有启用的勋章，请到 <A HREF=""Admin_Change.asp?action=medals"">勋章编辑</A> 功能中设定可用的勋章后再编辑。  "
			Else
				MyMedals = Rs1.GetRows(-1)
			End If
			Rs1.Close:Set Rs1=Nothing
			Echo "<br><br> "
			Echo "<body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10"">"
			Echo "<form method=""post"" action=""?action=editmedalsok&uid="&UID&""">"
			Echo "<table cellspacing=""1"" cellpadding=""4"" width=""95%"" align=""center"" class=""a2"">"
			Echo "<tr class=""a1"">"
			Echo "	<td colspan=""4"">勋章编辑 - "&Rs(0)&"</td>"
			Echo "</tr>"
			Echo "<tr class=""a4"" align=""center"">"
			Echo "		<td>勋章图片</td><td>名称</td><td>授予该勋章</td><td>授勋理由</td>"
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
			Echo "<input type=""submit"" name=""medalsubmit"" value=""提 交"">"
			Echo "</center></form><br><br>"
		End if
	End if
	Rs.Close:Set Rs=Nothing
End Sub
Sub editcreditsok
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " 参数错误 "
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
	SuccessMsg " 积分设置完成 。"
End Sub

Sub editcredits	
	Dim Gs,Value,i,m,Rs,u,UserInfo
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " 参数错误 "
	Else
		Set Rs = team.execute("Select Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,UserName,LevelName From ["&Isforum&"User] Where ID="& Uid)
		If Rs.Eof Then
			SuccessMsg " 指定的用户不存在。 "
		Else
%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>TEAM's 支持对用户 8 种扩展积分的设置，只有被启用的积分才允许您进行编辑。
		<li>对用户的积分，奖励为正数，惩罚为负数 。
      </ul></td>
  </tr>
</table>
<br>
<form name="input" method="post" action="?action=editcreditsok&uid=<%=UID%>">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="10">编辑用户积分 - <%=Rs(8)%>(<%=Split(RS(9),"||")(0)%>)</td>
    </tr>
    <tr class="a3" align="center">
      <td width="14%">用户详细积分</td>
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
    <input type="submit" name="creditsubmit" value="提 交">
  </center>
</form>
<%		End if
	End if
End Sub

Sub editgroupsok
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " 参数错误 "
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
			tmp = "用户 "&Trim(Request.Form("olduser"))&" 已经改名为 "&Trim(Request.Form("usernamenew"))&" "
		End if
		If Request.Form("clearquestion") = 1 Then
			team.execute("Update ["&Isforum&"User] Set Question='',Answer=''  Where ID="& Uid)
		End if
		If not IsValidEmail(Request.Form("emailnew")) Then
			SuccessMsg "邮件地址错误"
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
				SuccessMsg "您没有提升管理员的权限"
			ElseIf int(Request.Form("mygroups"))=2 Then
				If UserGroupID = 2 Then SuccessMsg "您没有提升用户到相同等级的权限"
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
		SuccessMsg tmp & "<BR> 用户信息更新成功。"
	End if
End sub

Sub SetUserMamdber(s,m)
	Dim Rs
	Set Rs= team.execute("Select GroupName,UserColor,UserImg,rank,Members From ["&IsForum&"UserGroup] Where ID="& Int(s))
	If Rs.Eof Then
		SuccessMsg "用户权限表损坏，请重新导入。"
	Else
		team.Execute("Update ["&Isforum&"User] Set Levelname='"&Rs(0)&"||"&Rs(1)&"||"&Rs(2)&"||"&Rs(3)&"||0',Members='"&Rs(4)&"' Where ID="& Int(m))
	End if
End Sub

Sub editgroups	
	Dim Gs,Value,i,m,Rs,u,UserInfo,uGID,Us,UserGroupID
	If Uid = "" Or Not isNumeric(Uid) Then 
		SuccessMsg " 参数错误 "
	Else
		Set Us = team.execute("Select UserGroupID From ["&Isforum&"User] Where UserName='"& tk_userName &"'")
		If Not Us.Eof Then
			UserGroupID = Int(Us(0))
		End If 
		Us.Close:Set Us=Nothing 
		Set Rs = team.execute("Select UserName,UserGroupID,Posttopic,Postrevert,Deltopic,Goodtopic,Usermail,Userhome,Userface,UserCity,UserSex,Honor,Birthday,Sign,Degree,RegIP,Regtime,Landtime,UserInfo,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members From ["&Isforum&"User] Where ID="& Uid)
		If Rs.Eof Then
			SuccessMsg " 指定的用户不存在。 "
		Else
			If Int(Rs(1))=1 Then
				If Not (Trim(tk_userName) = Trim(WebSuperAdmin)) Then SuccessMsg "管理员等级用户的资料只能被内置的网站管理员修改, 内置管理员的设置请打开Conn.asp,修改里面的"
			Elseif Int(Rs(1))=2 Then
				If UserGroupID=2 Then  SuccessMsg "您不能在后台修改高于或与您一样等级的用户资料,包括您自己的资料"
			End if
			UserInfo = Split(Rs(18),"|")
%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>如果需要提升用户权限到管理员，请打开Conn.asp 修改里面 Const WebSuperAdmin = "admin"		'默认内置管理员, 
		WebSuperAdmin 就是您的管理员的名称， 只有设置成为默认内置管理员的用户，才有提升其他用户到管理员的权限。
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=editgroupsok&uid=<%=UID%>">
  <input type="hidden" value="<%=Rs(0)%>" name="olduser">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">用户所属组类别</td>
    </tr>
    <tr class="a4">
      <td><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr class="a4">
		  <input type="hidden" name="oldgroup" value="<%=Rs(1)%>"> <%
	Set Gs = team.execute("Select Max(ID) as ID,GroupName From ["&IsForum&"UserGroup] Group By GroupName")
	If Gs.Eof or Gs.Bof Then
		SuccessMsg " 用户权限表数据损坏,请手动导入新表! "
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
      <td colspan="2">编辑用户 [ <%=Rs(0)%> ] </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>用户名:</b><br>
        <span class="a3">如不是特别需要，请不要修改用户名</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="usernamenew" value="<%=Rs(0)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>新密码:</b><br>
        <span class="a3">如果不更改密码此处请留空</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="passwordnew" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>清除用户安全提问:</b><br>
        <span class="a3">选择“是”将清除用户安全提问，该用户将不需要回答安全提问即可登录；选择“否”为不改变用户的安全提问设置</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="clearquestion" value="1">
        是 &nbsp; &nbsp;
        <input type="radio" name="clearquestion" value="0" checked>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>个人头衔:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="honor" value="<%=Rs(11)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>性别:</b></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="gendernew" value="1" <%If Rs(10)=1 Then%>checked<%End if%>>
        男
        <input type="radio" name="gendernew" value="2" <%If Rs(10)=2 Then%>checked<%End if%>>
        女
        <input type="radio" name="gendernew" value="0" <%If Rs(10)=0 Then%>checked<%End if%>>
        保密</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>Email:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="emailnew" value="<%=Rs(6)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>发帖数:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="postsnew" value="<%=Rs(2)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>回帖数:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="postsre" value="<%=RS(3)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>精华帖数:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="digestpostsnew" value="<%=Rs(5)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>总在线时间(分钟):</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="totalnew" value="<%=Rs(14)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>注册 IP:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="regipnew" value="<%=Rs(15)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>注册时间:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="regdatenew" value="<%=Rs(16)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>上次访问:</b></td>
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
  <a name="用户资料"></a>
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">用户资料</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>主页:</b></td>
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
      <td width="60%" bgcolor="#F8F8F8" ><b>淘宝:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="taobao" value="<%=UserInfo(4)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>支付宝:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="alibuy" value="<%=UserInfo(5)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>来自:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="locationnew" value="<%=Rs(9)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>生日:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="bdaynew" value="<%=Rs(12)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>头像:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="avatarnew" value="<%=Rs(8)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>签名:</b></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="signaturenew" cols="30" style="height:70;overflow-y:visible;"><%=Rs(13)%></textarea></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="searchsubmit" value="编辑" <%If Session("UserMember") = 2 and  Rs(1) = 1 Then%>disabled<%end if%>>
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
		SuccessMsg "你没有选定需要删除的用户ID"
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
					team.execute("Update ["&IsForum&"bbsconfig] set Board_Last='<A href=Thread.asp?tid="&rs1(0)&" target=""_blank"">"&Cutstr(rs1(1),200)&"</a> →$@$"&TK_UserName&"$@$"&Now()&"' where id="&RS(0))
				End If
				Rs.MoveNext
			Loop
			team.execute("Delete from ["&isforum&"User] Where ID="&ho)
			Cache.DelCache("BoardLists")
		next
		SuccessMsg " 指定的用户已经删除。"
	End If
End Sub

Sub findmembers
	Dim username,lookperm,Looks,findpages,usermail,regip,usersign,maxpost,maxlogin,regtime
	Dim CountNum,Rs,PageSNum,i,NextSeach
	Dim tmp,u,ExtCredits,ExtSort
	PageSNum = Cid(Request.Form("findpages"))		'查询结果分页
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
	If Request.Form("searchsubmit") = "搜索用户" Then
		Echo "<body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10"">"
		Echo "<table cellspacing=""1"" cellpadding=""4"" width=""95%"" align=""center"" class=""a2"">"
		Echo " <form method=""post"" action=""?action=members"">"
		Echo "<tr align=""center"" class=""a1"">"
		Echo "	<td width=""48""><input type=""checkbox"" name=""chkall"" onclick=""checkall(this.form, 'delete')"">删?</td>"
		Echo "	<td>用户名</td>"
		Echo "	<td>发帖数</td><td>用户组</td><td>注册时间</td><td>登陆时间</td><td>编辑</td></tr>"
		If Rs.Eof Then
			Echo "<tr align=""center"" class=""a4""><td Colspan=""7""> 对不起，没有找到符合条件的用户。</td></tr>"
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
			Echo "		<td><a href=""?action=editgroups&uid="&RS(0)&""">[用户属性]</a> <a href=""?action=editcredits&uid="&RS(0)&""">[积分]</a> <a href=""?action=editmedals&uid="&RS(0)&""">[勋章]</a> </td>"
			Echo "	</tr>"
			Rs.MoveNext
		Loop
		Echo "</table><br><center> <input type=""submit"" name=""searchsubmit"" value=""删除用户""></center></form>"
	End If
	If Request.Form("newslettersubmit") = "论坛通知" Then
		Echo " <BR><BR><form method=""post"" action=""?action=annonces"">"
		Do While Not Rs.Eof
			Echo  "	<input type=""hidden"" name=""msgid"" value="""&RS(0)&""">"
		Rs.MoveNext
		Loop
		%>
		<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
		<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
		<tr class="a1"><td colspan="9">符合条件的会员数: <%=TopUser%></td></tr>
		<tr>
			<td bgcolor="#F8F8F8">标题:</td>
			<td bgcolor="#FFFFFF"><input type="text" name="subject" size="80" value=></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8" valign="top">内容:</td><td bgcolor="#FFFFFF">
			<textarea cols="80" rows="10" name="message"></textarea></td></tr>
		<tr>
			<td bgcolor="#F8F8F8">发送方式:</td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="email" name="sendvia"> Email<input type="radio" value="pm" checked name="sendvia"> 短消息</td>
		</tr>
		</table><br>
		<center><input type="submit" name="sendsubmit" value="提 交"></center></form><br><br>
		<%
	End If
	If Request.Form("creditsubmit") = "积分奖惩" Then 
		Echo " <BR><BR><form method=""post"" action=""?action=edituserexc"">"
		Do While Not Rs.Eof
			Echo  "	<input type=""hidden"" name=""msgid"" value="""&RS(0)&""">"
			Rs.MoveNext
		Loop
	%>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	 <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
		<tr class="a1">
		  <td colspan="10">符合条件的会员数: <%=TopUser%></td>
		</tr>
		<tr class="a3" align="center">
			<td width="14%">用户详细积分</td>
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
			Echo " </tr><tr align=""center"" class=""a4""><td bgcolor=""#F8F8F8""> 奖惩数值 </td>"
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
		<tr class="a1"><td colspan="9"><input class="a1" type="checkbox" name="sendcreditsletter" value="1">发送积分变更通知</td></tr>
		<tr>
			<td bgcolor="#F8F8F8">标题:</td>
			<td bgcolor="#FFFFFF"><input type="text" name="subject" size="80" value=></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8" valign="top">内容:</td><td bgcolor="#FFFFFF">
			<textarea cols="80" rows="10" name="message"></textarea></td></tr>
		<tr>
			<td bgcolor="#F8F8F8">发送方式:</td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="email" name="sendvia"> Email<input type="radio" value="pm" checked name="sendvia"> 短消息</td>
		</tr>
		</table><br>
	 <center><input type="submit" name="creditsubmit" value="提 交"></center></form><%
	End If
	Rs.Close:Set Rs=Nothing
End Sub

Sub edituserexc
	If Request.Form("msgid") = "" Then
		SuccessMsg "  不存在需要发送的用户。"
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
					team.Execute( "insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('系统消息','"&team.Execute("Select UserName From ["&isforum&"User] Where ID="& ho)(0)&"','"&Request.Form("message")&"',"&SqlNowString&",'"&Request.Form("subject")&"')" )
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
	SuccessMsg " 积分设置完成，请等待系统自动返回到 <a href=Admin_User.asp>编辑用户 </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_User.asp> 。"
End Sub


Sub annonces
	Dim MsgName,ho,msgmail
	If Request.Form("msgid") = "" Then
		SuccessMsg "  不存在需要发送的用户。"
	Else
		If Len(Request.Form("message"))<5 or Request.Form("subject") = "" Then 
			SuccessMsg " 内容或标题不能为空 。"
		Else
			If request.Form("sendvia") = "pm" Then
				for each ho in Request.Form("msgid")
					team.execute("Update ["&isforum&"User] Set Newmessage=Newmessage+1 Where ID="&ho)
					team.Execute( "insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('系统消息','"&team.Execute("Select UserName From ["&isforum&"User] Where ID="& ho)(0)&"','"&Request.Form("message")&"',"&SqlNowString&",'"&Request.Form("subject")&"')" )
				next
			Else
				for each ho in Request.Form("msgid")
					If IsValidEmail(team.Execute("Select UserMail From ["&isforum&"User] Where ID="& ho)(0)) Then
						Call Emailto ( team.Execute("Select UserMail From ["&isforum&"User] Where ID="& ho)(0), Request.Form("subject") , Request.Form("message"))
					End if
				next
			End if
		End if
		SuccessMsg " 信息已发送成功，请等待系统自动返回到 <a href=Admin_User.asp>编辑用户 </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_User.asp> 。"
	End If
End Sub

Sub setuserok
	Dim source1,source2,source3,target,UserTrg,rs,MsgTrg,MsgTrg1
	Source1 = HtmlEncode(Trim(Request.Form("source1")))
	Source2 = HtmlEncode(Trim(Request.Form("source2")))
	Source3 = HtmlEncode(Trim(Request.Form("source3")))
	Target = HtmlEncode(Trim(Request.Form("target")))
	If Source1 & Source2 & Source3 &"" = "" Then
		SuccessMsg "原用户名栏至少需要一行数据，不能全部为空。"
	ElseIf Source1&"" = "" and Source2 & Source3 &"" <>"" Then 
		SuccessMsg "用户名输入请从 <FONT COLOR=""red""><B>原用户名 1</B></FONT> 栏开始。" 	
	Else	
		If Source1 & ""<>"" Then
			If team.execute("Select * from ["&Isforum&"User] Where UserName='"&Source1&"'").Eof Then
				SuccessMsg " 系统不存在名为"&Source1&"的用户名 。" 
			End If
		End If
		If Source2 & ""<>"" Then
			If team.execute("Select * from ["&Isforum&"User] Where UserName='"&Source2&"'").Eof Then
				SuccessMsg " 系统不存在名为"&Source2&"的用户名 。" 
			End if
		End If
		If Source3 & ""<>"" Then
			If team.execute("Select * from ["&Isforum&"User] Where UserName='"&Source3&"'").Eof Then
				SuccessMsg " 系统不存在名为"&Source3&"的用户名 。" 	
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
		'删除原用户
		Team.Execute("Delete From ["&Isforum&"User] Where "&UserTrg&" ")
	End If
	SuccessMsg " 用户何并成功，原用户的主贴，积分，已全部转入目标用户，同时原用户已被删除 。"
End Sub

Sub  Setuser%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?action=setuserok">
  <table cellspacing="1" cellpadding="4" width="85%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">合并用户 - 原用户的帖子、积分全部转入目标用户，同时删除原用户</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">原用户名 1:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="source1" size="20"></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">原用户名 2:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="source2" size="20"></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">原用户名 3:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="source3" size="20"></td>
    </tr>
    <tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">目标用户名:</td>
      <td bgcolor="#FFFFFF" width="60%"><input type="text" name="target" size="20"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="mergesubmit" value="提 交">
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
		SuccessMsg "用户名不能为空"
	End If
	If Newpassword = "" Or IsNull(Newpassword) Then
		SuccessMsg "密码不能为空"
	End If	
	If Not IsValidEmail(Newemail) Then
		SuccessMsg "邮件格式错误 !"
	End If
	CheckStr=Array("=","%",chr(32),"?","&",";",",","'",",",chr(34),chr(9),"","$","|")
	For i=0 To Ubound(CheckStr)
		If Instr(NewUserName,CheckStr(i))>0 then SuccessMsg "用户名中不能含有特殊符号"
	Next
	If Not team.execute("Select * From ["&Isforum&"User] Where UserName='"&NewUserName&"'").Eof Then
		SuccessMsg " 用户名重复，请重新输入一个用户名。"
	End If
	ExtCredits= Split(team.Club_Class(21),"|")
	team.Execute( "insert into ["&Isforum&"User] (UserName,Userpass,UserGroupID,Usermail,UserSex,Newmessage,Posttopic,Postrevert,Deltopic,Goodtopic,Extcredits0,Extcredits1,Extcredits2,Regtime,Landtime,Postblog,UserInfo,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,LevelName,Members) values('"&NewUserName&"','"&MD5(Newpassword,16)&"',26,'"&Newemail&"',0,0,0,0,0,0,"&Cid(Split(ExtCredits(0),",")(2))&","&Cid(Split(ExtCredits(1),",")(2))&","&Cid(Split(ExtCredits(2),",")(2))&","&SqlNowString&","&SqlNowString&",0,'|||||',"&Cid(Split(ExtCredits(3),",")(2))&","&Cid(Split(ExtCredits(4),",")(2))&","&Cid(Split(ExtCredits(5),",")(2))&","&Cid(Split(ExtCredits(6),",")(2))&","&Cid(Split(ExtCredits(7),",")(2))&",'附小一年级||||||0||0','注册会员')" )

	If Request.Form("emailnotify") = "yes" Then
		Dim Mailtopic,Body
		Mailtopic="请确认用户名注册通知。"
		Body="亲爱的"&NewUserName&", 您好!"&vbCrlf&""&vbCrlf&" [本邮件由系统自动发送] 恭喜您得到 "&team.Club_Class(1)&" 的注册用户权限。"&vbCrlf&""&vbCrlf&"　* 您的帐号是:"&NewUserName&"　密码是:"&Newpassword&" "&vbCrlf&""&vbCrlf&"　* "&vbCrlf&""&vbCrlf&" * 最后, 有几点注意事项请您牢记"&vbCrlf&"1、请遵守《计算机信息网络国际联网安全保护管理办法》里的一切规定。"&vbCrlf&"2、使用轻松而健康的话题，所以请不要涉及政治、宗教等敏感话题。"&vbCrlf&"3、承担一切因您的行为而直接或间接导致的民事或刑事法律责任。"&vbCrlf&""&vbCrlf&""&vbCrlf&"论坛服务由 "&team.Club_Class(1)&"("&team.Club_Class(2)&") 提供 。"&vbCrlf&"[本论坛源程序由:TEAM5.CN提供]"&vbCrlf&""&vbCrlf&""&vbCrlf&""
		Call Emailto(Newemail,Mailtopic,Body)
	End If
	SuccessMsg " 新用户 "&NewUserName&" 已经添加完成，默认密码为 "&Newpassword&"。 "
End Sub

Sub AddUser %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=adduserok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="2">添加新用户</td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">用户名:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="text" name="newusername"></td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">密码:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="text" name="newpassword"></td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">Email:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="text" name="newemail"></td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">发送通知到上述地址:</td>
      <td align="right" bgcolor="#FFFFFF"><input type="checkbox" name="emailnotify" value="yes"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="addsubmit" value="提 交">
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
		SuccessMsg " 用户权限表数据损坏,请手动导入新表! "
	Else
		Value = Gs.GetRows(-1)
	End If
	Gs.Close:Set Gs=Nothing
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
  <tr class="a1">
    <td>TEAM's 提示</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li> 系统共有用户 <B><%=Application(CacheName&"_UserNum")%></B> 位。</li>
        <li> 查找用户可以使用直接输入用户，也可以使用相关资料来模糊查找。</li>
        <li> 系统提供3个提交按钮功能，对用户的快捷管理可以通过此处直接执行。</li>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=findmembers">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr class="a1">
      <td colspan="2" align="center">用户管理 </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">用户名包含：</td>
      <td bgcolor="#FFFFFF"><input type="text" name="username" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">24小时内登录的用户：</td>
      <td bgcolor="#FFFFFF"><input type="checkbox" name="userlogin" value="1">
        选定 </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">24小时内注册的用户：</td>
      <td bgcolor="#FFFFFF"><input type="checkbox" name="newuserreg" value="1">
        选定 </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">按用户组别查询：</td>
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
    <input type="submit" name="searchsubmit" value="搜索用户">
    &nbsp;
    <input type="submit" name="newslettersubmit" value="论坛通知">
    &nbsp;
    <input type="submit" name="creditsubmit" value="积分奖惩">
    &nbsp;
  </center>
  <BR>
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr class="a1">
      <td align="center" colspan="2">高级查询</td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">最高查询数：</td>
      <td bgcolor="#FFFFFF"><input type="text" name="findpages" size="40" value="20">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">Email包含：</td>
      <td bgcolor="#FFFFFF"><input type="text" name="usermail" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">注册IP包含：</td>
      <td bgcolor="#FFFFFF"><input type="text" name="regip" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">用户签名包含：</td>
      <td bgcolor="#FFFFFF"><input type="text" name="usersign" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">发贴+回帖总数：
        <input type="radio" name="maxpost1" value="1" checked>
        &nbsp;多于&nbsp;
        <input type="radio" name="maxpost1" value="0">
        &nbsp;少于</td>
      <td bgcolor="#FFFFFF"><input type="text" name="maxpost" size="40" value="">
      </td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">在线时长：
        <input type="radio" name="maxlogin1" value="1" checked>
        &nbsp;多于&nbsp;
        <input type="radio" name="maxlogin1" value="0">
        &nbsp;少于</td>
      <td bgcolor="#FFFFFF"><input type="text" name="maxlogin" size="40" value="">
        (分钟)</td>
    </tr>
    <tr>
      <td bgcolor="#F8F8F8">注册日期：
        <input type="radio" name="regtime1" value="0" checked>
        &nbsp;早于&nbsp;
        <input type="radio" name="regtime1" value="1">
        &nbsp;等于&nbsp;
        <input type="radio" name="regtime1" value="2">
        &nbsp;晚于 (yyyy-mm-dd):</td>
      <td bgcolor="#FFFFFF"><input type="text" name="regtime" size="40" value="">
      </td>
    </tr>
    <%
	Dim ExtCredits,ExtSort
	ExtCredits= Split(team.Club_Class(21),"|")
	For U=0 to 2
		ExtSort=Split(ExtCredits(U),",")
		Echo "<tr>"
		Echo " <td bgcolor=""#F8F8F8"">"&ExtSort(0)&"："
		Echo "	<input type=radio name=""Nums"&U&""" value=""1"" checked>&nbsp;多于&nbsp; "
        Echo "	<input type=radio name=""Nums"&U&""" value=""0"">&nbsp;少于</td>"
		Echo "	<td bgcolor=""#FFFFFF""><input type=""text"" name=""MyCred"&U&""" size=""40"" value=""""></td>"
		Echo "</tr>"
	Next
	%>
  </table>
  <BR>
  <center>
    <input type="submit" name="searchsubmit" value="搜索用户">
    &nbsp;
    <input type="submit" name="newslettersubmit" value="论坛通知">
    &nbsp;
    <input type="submit" name="creditsubmit" value="积分奖惩">
    &nbsp;
  </center>
  <BR>
</form>
<%
end sub

footer()
%>
