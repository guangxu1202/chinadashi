<!-- #include file="CONN.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim x1,x2,fID
team.Headers(Team.Club_Class(1))
Call ForUserBoard
Select Case Request("action")
	Case "killname"
		Call killname
	Case "managesok"
		Call managesok
	Case "killuserok"
		Call killuserok
	Case "gotoname"
		Call gotoname
	Case "boardlist"
		Call boardlist
	Case "boardlistok"
		Call boardlistok
	Case "forumcheck"
		Call forumcheck
	Case "forumcheckok"
		Call forumcheckok
	Case Else
		Call Main()
End Select
Call team.footer


Sub forumcheckok
	Dim ho,mo
	If Request("forumlinksubmit")="" Then 
		for each ho in request.form("checktid")
			team.execute("Update ["&Isforum&"Forum] Set Auditing=0 Where ID="&Int(ho))
		next
		for each mo in request.form("checkrid")
			team.execute("Update ["&Isforum & team.Club_Class(11) &"] Set Auditing=0 Where ID="&Int(mo))
		Next
		team.SaveLog ("执行帖子的审核操作")
		team.Error "帖子的审核操作成功。请等待系统自动返回到 <a href=BoradServer.asp?action=forumcheck> 审核帖子  </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=forumcheck>"
	Else
		for each ho in request.form("checktid")
			team.execute("Update ["&Isforum&"Forum] Set Deltopic=1 Where ID="&Int(ho))
		next
		for each mo in request.form("checkrid")
			team.execute("Delete From ["&Isforum & team.Club_Class(11) &"] Where Auditing=1 and ID="&Int(mo))
		Next
		team.SaveLog ("删除未审核的帖子")
		team.Error "选定的未审核帖子已经被删除。请等待系统自动返回到 <a href=BoradServer.asp?action=forumcheck> 审核帖子  </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=forumcheck>"
	End if
End Sub

Sub forumcheck
	Dim tmp,SQL,SqlQueryNum,RS,Maxpage,PageNum,iRs,Rmp,DispCount,i,Page,Chcheid,j,Rs1,MyRs,m
	Dim Nexs
	x1 = " <a href=""BoradServer.asp?action=boardlist"">前台管理界面</a> "
	tmp = Replace(Team.ElseHtml (8),"{$weburl}",team.MenuTitle)
	Rmp = "<form name=""myform"" method=""post"" action=""?action=forumcheckok"">"
	Rmp = Rmp & "<tr class=""tab1""><td align=""center"" width=""20%""><input type=""checkbox"" name=""chkall"" onClick=""checkall(this.form)"" class=""radio"">删除 / 审核</td> <td align=""center"">(操作提示：删除帖子请使用“提交”建，审核请使用“恢复”键) 标题 </td></tr>"
	DispCount = team.execute("Select Count(ID) From ["&IsForum&"forum] Where Auditing=1 and Deltopic=0 ")(0)
	SQL="Select ID,Topic,UserName From ["&IsForum&"forum] Where Auditing=1 and Deltopic=0 Order By Lasttime DESC"
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum = SqlQueryNum+1
		Maxpage = Cid(team.Forum_setting(19))		'每页分页数
		PageNum = Abs(int(-Abs(DispCount/Maxpage)))	'页数
		Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		iRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
	If Isarray(iRs) Then
		For i=0 To Ubound(iRs,2)
			If Nexs = "" Then
				Nexs = iRs(0,i)
			Else
				Nexs = Nexs &  "," & iRs(0,i)
			End if
			Rmp = Rmp & "<tr class=""a4""><td align=""center""><input type=""checkbox"" name=""checktid"" class=""radio"" value="""&iRs(0,i)&"""></td><td> <a href=""SeeDeltop.asp?tid="&iRs(0,i)&""" target=""_blank""> "& iRs(1,i) &" </a> <A href=""Profile.asp?username="& iRs(2,i) &""">("& iRs(2,i) &")</a> </td></tr>"
			MyRs=""
			Set Rs1 = team.execute("Select ID,ReTopic,content,username From ["& IsForum & team.Club_Class(11) &"] Where Auditing=1 And topicid="& Int(iRs(0,i)))
			If Not Rs1.Eof Then
				MyRs = Rs1.GetRows(-1)
			End If
			Rs1.Close:Set Rs1=Nothing
			If IsArray(MyRs) Then
				For m=0 To UBound(MyRs,2)
					Rmp = Rmp & "<tr class=""a4"">"
					Rmp = Rmp & "	<td align=""center""></td>"
					Rmp = Rmp & "	<td>回复： <input type=""checkbox"" name=""checkrid"" class=""radio"" value="""&MyRs(0,m)&"""> <a href=""SeeDeltop.asp?tid="&iRs(0,i)&"#RID"&MyRs(0,m)&""" target=""_blank""> "& MyRs(2,m) &"</a> </td></tr>"				
				Next
			End If
		Next
	End If 
	If Nexs = "" Then
		Set Rs1 = team.execute("Select ID,ReTopic,content,username,topicid From ["& IsForum & team.Club_Class(11) &"] Where Auditing=1")
	Else
		Set Rs1 = team.execute("Select ID,ReTopic,content,username,topicid From ["& IsForum & team.Club_Class(11) &"] Where Auditing=1 and Not topicid In ("& Nexs &") ")
	End if
	If Not Rs1.Eof Then
		MyRs = Rs1.GetRows(-1)
	End If
	Rs1.Close:Set Rs1=Nothing
	If IsArray(MyRs) Then
		For m=0 To UBound(MyRs,2)
			Rmp = Rmp & "<tr class=""a4"">"
			Rmp = Rmp & "	<td align=""center""><input type=""checkbox"" name=""checkrid"" class=""radio"" value="""&MyRs(0,m)&"""></td><td>审核回复： [ <a href=""thread.asp?tid="&MyRs(4,m)&"#"&MyRs(0,m)&""" target=""_blank""> 浏览主题 </a> ] / "& MyRs(2,m) &"  </td></tr>"				
		Next
	End If
	tmp=Replace(tmp,"{$forumlist}",Rmp)
	tmp=Replace(tmp,"{$userkill}","")
	tmp=Replace(tmp,"{$pagecount}",PageNum)
	tmp=Replace(tmp,"{$dispcount}",DispCount)
	Echo tmp
End sub

Sub gotoname
	Dim UID,Rs,uName
	UID = HRF(2,2,"uid")
	uname = HRF(2,1,"uname")
	Set Rs = team.execute("Select * From ["&IsForum&"User] Where ID="& Int(UID) )
	If Rs.Eof Then
		team.Error "系统不存在此用户。请等待系统自动返回到 <a href=BoradServer.asp?action=killname> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=killname>"
	Else
		team.execute("Update ["&IsForum&"User] Set UserGroupID=27,Levelname='附小一年级||||||0||0' Where ID="& Int(UID))
		team.SaveLog ("将用户"&uname&"恢复为正常状态的操作。")
		team.Error "选定的用户已经恢复为正常的状态。请等待系统自动返回到 <a href=BoradServer.asp?action=killname> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=killname>"
	End if
End Sub

Sub killuserok
	Dim getuser,getusermeber,RS
	GetUser = HRF(1,1,"myname") 
	GetuserMeber = HRF(1,2,"getusermeber")
	If GetuserMeber = 0 Then
		team.Error "您没有选定操作选项。请等待系统自动返回到 <a href=BoradServer.asp?action=killname> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=killname>"
	End if
	Set Rs = team.execute("Select UserGroupID From ["&IsForum&"User] Where UserName = '"&GetUser&"'")
	If Rs.Eof And Rs.Bof Then
		team.Error "系统不存在此用户。请等待系统自动返回到 <a href=BoradServer.asp?action=killname> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=killname>"
	Else
		If Int(Rs(0))=1 Or Int(Rs(0))=2 Or Int(Rs(0))=3 Or Int(Rs(0))=4  Then
			team.Error "您不能对管理等级的用户进行此项操作。"
		End If
		If GetuserMeber = 6 Then
			team.execute("Update ["&IsForum&"User] Set UserGroupID=6,Levelname='禁止发言||||||0||0' Where UserName='"&GetUser&"'")
			team.SaveLog ("将用户"&GetUser&"进行禁止发言的操作。")
		ElseIf GetuserMeber = 7 Then
			team.execute("Update ["&IsForum&"User] Set UserGroupID=7,Levelname='禁止访问||||||0||0' Where UserName='"&GetUser&"'")
			team.SaveLog ("将用户"&GetUser&"进行禁止访问的操作。")
		End If
		team.Error "此用户已经被设置为选定的等级。请等待系统自动返回到 <a href=BoradServer.asp?action=killname> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=killname>"
	End if
End Sub

Sub boardlistok
	Dim fid,ho
	for each ho in request.form("fid")
		team.execute("Update ["&Isforum&"bbsConfig] Set Readme='"&HRF(1,1,"Readme"&ho&"")&"',Board_Key='"&HRF(1,1,"Board_Key"&ho&"")&"' Where ID="& Int(ho))
		Cache.DelCache("ForumsBoards_"&ho)
		Cache.DelCache("Boards_"&ho)
		team.SaveLog ("修改版块Forums.asp?fid="& ho &"的资料")
	Next
	Cache.DelCache("BoardLists")
	
	team.Error "版块信息已经修改完成。请等待系统自动返回到 <a href=BoradServer.asp?action=boardlist> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp?action=boardlist>"
End Sub

Sub boardlist
	Dim tmp,rmp,RS,wmp,t,Board_Setting,twhere
	x1 = " <a href=""BoradServer.asp?action=boardlist"">前台管理界面</a> "
	tmp = Replace(Team.ElseHtml (8),"{$weburl}",team.MenuTitle)
	Rmp = "<form name=""myform"" method=""post"" action=""?action=boardlistok"">"
	t = ""
	Set Rs = team.execute("Select BoardID From ["&IsForum&"Moderators] Where ManageUser ='"& tk_UserName &"'")
	Do While Not Rs.Eof
		If t = "" Then
			t = Rs(0)
		Else
			t = t & "," & Rs(0) 
		End If 
		Rs.MoveNext
	Loop
	Rs.close:Set Rs = Nothing
	If Not team.IsMaster or Not team.SuperMaster Then
		Set Rs = team.execute("Select ID,bbsname,Readme,Board_Key,Board_Setting From ["&IsForum&"bbsConfig] Where followid>0 ")
	Else 
		If t <>"" Then
			Set Rs = team.execute("Select ID,bbsname,Readme,Board_Key,Board_Setting From ["&IsForum&"bbsConfig] Where ID in ("&t&") and followid>0 ")
		Else
			Set Rs = team.execute("Select ID,bbsname,Readme,Board_Key,Board_Setting From ["&IsForum&"bbsConfig] Where ID =0 ")
		End If
	End If
	Do While Not Rs.Eof
		Board_Setting = ""
		Board_Setting = Split(Rs(4),"$$$")
		If Int(Board_Setting(1)) = 0 And (Not team.IsMaster and Not team.SuperMaster) Then
			Rmp = Rmp & "<tr class=""a4""><td colspan=""2""> 论坛设置了版主不能修改版规和介绍。</td></tr>"
		Else
			Rmp = Rmp & "<input type=""hidden"" name=""fid"" value="""&Rs(0)&"""><tr class=""tab1""><td> 版块名称：<a href=""Forums.asp?fid="&RS(0)&""">"& Rs(1) &"</a> </td><td> 编辑详细 </td></tr>"
			Rmp = Rmp & "<tr class=""a4""><td width=""50%""> <b>论坛简介:</b><br> 将显示于论坛名称的下面，提供对本论坛的简短描述，支持html代码  </td><td><textarea rows=""5"" name=""Readme"&Rs(0)&""" cols=""30"" style=""height:70;overflow-y:visible;width:100%;"">"& ReplaceStr(RS(2),"<BR>",VbCrlf) &"</textarea> </td></tr>"
			Rmp = Rmp & "<tr class=""a4""><td> <b>本论坛规则:</b><br> 显示于主题列表页的当前论坛规则，支持 html 代码，留空为不显示 </td><td><textarea rows=""5"" name=""Board_Key"&Rs(0)&""" cols=""30"" style=""height:70;overflow-y:visible;width:100%;"">"&ReplaceStr(RS(3),"<BR>",VbCrlf)&"</textarea> </td></tr>"
			Rmp = Rmp & "<tr class=""a1""><td colspan=""2"" height=""5""></td></tr>"
		End if
		Rs.MoveNext
	Loop
	Rs.close:Set Rs = Nothing
	tmp=Replace(tmp,"{$forumlist}",Rmp)
	tmp=Replace(tmp,"{$userkill}","")
	tmp=Replace(tmp,"{$pagecount}",1)
	tmp=Replace(tmp,"{$dispcount}",1)
	Echo tmp
End Sub


Sub killname
	Dim tmp,rmp,RS,wmp
	x1 = " <a href=""BoradServer.asp?action=killname"">前台管理界面</a> "
	tmp = Replace(Team.ElseHtml (8),"{$weburl}",team.MenuTitle)
	Rmp = "<form name=""myform"" method=""post"" action=""?action=killuserok"">"
	Rmp = Rmp & "<tr class=""tab1""><td> 用户名称 </td><td> 管理操作</td></tr>"
	Rmp = Rmp & "<tr class=""tab4""><td> <input type=""text"" name=""myname"" size=""25"" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';"" class=""colorblur""> </td><td><select name=""getusermeber""><option value="""">设置该用户的等级</option>"
	Set Rs = team.execute("Select ID,GroupName from ["&IsForum&"UserGroup] Where ID=6 or ID=7")
	Do While Not Rs.Eof
		Rmp = Rmp & "<option value="""&Rs(0)&""">"&Rs(1)&"</option>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	Rmp = Rmp & "</select></td></tr>"
	tmp=Replace(tmp,"{$forumlist}",Rmp)
	wmp = "<br><table cellspacing=""1"" cellpadding=""3"" width=""100%"" border=""0"" align=""center"" class=""a2""><tr class=""tab1""><td width=""60%"">被执行操作的用户列表名称</td><td>操作</td></tr>"
	Set Rs = team.execute("Select ID,Username From ["&IsForum&"User] Where UserGroupID=6 or UserGroupID=7 ")
	Do While Not Rs.Eof
		wmp = wmp & "<tr class=""tab4""><td>"&RS(1)&"</td><td alt=""点击将此用户等级设置为注册基本等级""><a href=""?action=gotoname&uid="&Rs(0)&"&uname="&Rs(1)&""" title=""点击将此用户等级设置为注册基本等级""><img Src="""&team.Styleurl&"/get.gif"" border=""0""></a></td></tr>"
		Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	wmp = wmp & "</table>"
	tmp=Replace(tmp,"{$userkill}",wmp)
	tmp=Replace(tmp,"{$pagecount}",1)
	tmp=Replace(tmp,"{$dispcount}",1)
	Echo tmp
End Sub

Sub managesok
	Dim ho,mFso,fPath,Rs,fName
	If Request.form("deleteid") = "" Then 
		team.Error2 "您选择要操作的ID"
	End If
	If Request("resubmit")="" Then
		If team.UserGroupID=1 And tk_UserName = WebSuperAdmin Then 
			for each ho in Request.form("deleteid")
				Set Rs = team.execute("Select ReList From ["&Isforum&"forum] Where ID="& Int(ho))
				Do While Not Rs.Eof 
					team.execute("Delete from ["&Isforum & Rs(0) &"] Where topicid = "& Int(ho))
					Rs.MoveNext
				Loop
				team.execute("Delete from ["&Isforum&"forum] Where ID="& Int(ho))
			Next
			fPath = "Images/Upfile/"
			Set mFso = Server.CreateOBject("Scripting.FileSystemObject")
			for each ho in Request.form("deleteid")
				Set Rs = team.execute("Select FileName,UserName From ["&IsForum&"Upfile] Where ID="& Int(ho) )
				If Not Rs.Eof Then
					fName = fPath & Rs(0)
					If mFso.FileExists(Server.mappath(fName)) Then
						On Error Resume Next
						mFso.deletefile(Server.mappath(fName))
					End  If
					UpdateUserpostExc(Rs(1))
				End if
				team.execute("Delete from ["&IsForum&"Upfile] Where ID="&Int(ho))
			Next
			team.SaveLog ("删除回收箱的操作。")
			team.Error "回收箱指定的帖子已经被彻底删除了。请等待系统自动返回到 <a href=BoradServer.asp> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp>"
		Else 
			team.Error "您没有清空回收箱的权利"
		End If 
	Else
		for each ho in Request.form("deleteid")
			team.execute("Update ["&Isforum&"forum] Set deltopic=0 Where ID="& Int(ho))
		Next
		team.SaveLog ("复原回收箱指定的帖子的操作。")
		team.Error "回收箱指定的帖子已经被复原了。请等待系统自动返回到 <a href=BoradServer.asp> 前台管理 </a> 页面 。<meta http-equiv=refresh content=3;url=BoradServer.asp>"
	End if
End Sub

Sub UpdateUserpostExc(uName)
	'用户积分部分
	Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u
	Dim UserPostID,My_ExtSort
	If Not team.UserLoginED Then  Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	For U=0 to Ubound(ExtCredits)
		ExtSort=Split(ExtCredits(U),",")
		MustSort=Split(MustOpen(U),",")
		If ExtSort(3)=1 Then
			If U = 0 Then
				UExt = UExt &"Extcredits0=Extcredits0-"&MustSort(3)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"-"&MustSort(3)&""
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where UserName='"& HtmlEncode(uName)&"'")
End Sub

Sub Main()
	Dim tmp,SQL,SqlQueryNum,RS,Maxpage,PageNum,iRs,Rmp,DispCount,i,Page,Chcheid,j
	Dim MyIds
	x1 = " <a href=""BoradServer.asp"">前台管理界面</a> "
	Chcheid = team.BoardList
	MyIds = ""
	Set Rs = team.execute("Select BoardID From ["&IsForum&"Moderators] Where ManageUser='"&tk_UserName&"'")
	Do While Not Rs.Eof
		If MyIds = "" Then
			MyIds = " and forumid = "& Rs(0)
		Else
			MyIds = " or forumid = "& Rs(0)
		End if
		Rs.moveNext
	Loop
	Rs.close:Set Rs= Nothing 
	If team.IsMaster Or team.SuperMaster Then
		MyIds =  ""
	End If
	tmp = Replace(Team.ElseHtml (8),"{$weburl}",team.MenuTitle)
	DispCount = team.execute("Select Count(ID) From ["&IsForum&"forum] Where deltopic=1 ")(0)
	SQL="Select ID,Topic,UserName,Views,Replies,LastText,forumid From ["&IsForum&"forum] Where deltopic=1  "&MyIds&" Order By Lasttime DESC"
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum = SqlQueryNum+1
		Maxpage = Cid(team.Forum_setting(19))		'每页分页数
		PageNum = Abs(int(-Abs(DispCount/Maxpage)))	'页数
		Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		iRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
	If Not Isarray(iRs) Then
		tmp=Replace(tmp,"{$forumlist}","<tr class=""tab1""><td> 主题 </td><td> 所属版块 </td><td> 作者/管理者 </td><td> 回复/查看 </td></tr><tr class=""tab4""><td colspan=""5"">暂无删贴</td></tr>")
	Else
		Rmp ="<form name=""myform"" method=""post"" action=""?action=managesok""><tr class=""tab1""><td width=""7%""><input type=""checkbox"" name=""chkall"" class=""radio"" onClick=""checkall(this.form,'delete')"">全</td><td width=""55%""> 主题(查看/回复) </td><td width=""18%""> 所属版块 </td><td> 作者/管理者 </td></tr>"
		For i=0 To Ubound(iRs,2)
			Rmp = Rmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td><input type=""checkbox"" name=""deleteid"" value="""&iRs(0,i)&""" class=""radio""></td><td><a href=""SeeDeltop.asp?tid="&iRs(0,i)&""" target=""_blank"">"&iRs(1,i)&"</a> ("&iRs(3,i)&" / "&iRs(4,i)&") <img src="""&team.styleurl&"/new.gif"" border=""0"" align=""absmiddle""></td><td align=""center"">"
			If isarray(Chcheid) Then
				For j=0 to Ubound(Chcheid,2)
					If Cid(iRs(6,i)) = Cid(Chcheid(0,j)) Then
						Rmp = Rmp & "[ <A href=Forums.asp?fid="&Chcheid(0,j)&" target=""_blank"">"&Chcheid(1,j)&"</a> ]"
					End if
				Next
			End If
			Rmp = Rmp & "</td><td align=""center""> "&iRs(2,i)&" / "&Split(iRs(5,i),"$@$")(0)&" </td></tr> "
		Next
		tmp=Replace(tmp,"{$forumlist}",Rmp)
	End If
	tmp=Replace(tmp,"{$pagecount}",PageNum)
	tmp=Replace(tmp,"{$dispcount}",DispCount)
	tmp=Replace(tmp,"{$userkill}","")
	Echo tmp
End Sub

Sub ForUserBoard
	If Not team.UserLoginED Then
		team.Error " 你未登陆论坛。<meta http-equiv=refresh content=3;url=login.asp> "
	End if
	If Not team.ManageUser Then
		team.Error " 您的权限不够，不能参与论坛管理 。"
	End if
End Sub
%>
