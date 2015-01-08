<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim fID,tID,x1,x2,rID,Mstemp,GInfo
Dim MasterMenu,Inputinfo,temp
Public Values,Board_Setting
fID = HRF(1,2,"fid")
tID = HRF(1,2,"tid")
Call ManageClass()

Select Case Request("action")
	Case "move"
		MasterMenu="移动主题"
		Inputinfo = "moves"
		call movepage
	Case "moves"
		Call forummove
	Case "deltopic"
		MasterMenu="删除帖子"
		Inputinfo = "deltopics"
		Call movepage
	Case "deltopics"
		Call deltopics
	case "movenew"
		MasterMenu="拉前主题"
		Inputinfo = "movenews"
		call movepage
	case "movenews"
		call movenews
	Case "islockpage"
		MasterMenu="锁定/解除锁定帖子"
		Inputinfo = "islockpages"
		call movepage
	Case "islockpages"
		islockpages
	Case "isclosepage"
		MasterMenu="关闭/打开主题"
		Inputinfo = "isclosepages"
		call movepage
	Case "isclosepages"
		isclosepages
	case "settops"
		MasterMenu="置顶/解除置顶"
		Inputinfo = "settopss"
		call movepage
	case "settopss"
		call settopss
	Case "getlike"
		MasterMenu="批量设置主题分类"
		Inputinfo = "getlikes"
		call movepage
	Case "getlikes"
		Call getlikes
	Case "digest"
		MasterMenu="加入/解除精华"
		Inputinfo = "digests"
		call movepage	
	Case "iskillofget"
		MasterMenu="奖励/惩罚"
		Inputinfo = "iskillofget"
		call movepage
	Case "digests"
		Call digests
	Case Else
		team.Error "参数错误"
End Select	

Sub movepage 
	If request.form("ismanage")="" and request.form("fismanage")="" Then 
		team.Error("您没有选择主题或相应的管理操作，请返回修改。")
	End If
	Echo "<form method=""post"" action=""?action="&Inputinfo&""">"
	Echo "<input type=""hidden"" value="""&fID&""" name=""fid"">"
	Echo "<input type=""hidden"" value="""&Request("rid")&""" name=""rid"">"
	Echo "<input type=""hidden"" value="""&Request("fismanage")&""" name=""fismanage"">"
	Echo "<table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=""a2"">"
	Echo " <tr>"
	Echo "		<td class=""a1"" colspan=""2"">TEAM's提示: "&MasterMenu&"</td>"
	Echo " </tr>"
	Echo " <tr class=""a4"">"
	Echo "		<td width=""40%""><B>用户名</B>:</td><td>"&TK_UserName&"</td>"
	Echo "</tr>"
	If Request("action")="settops" Then
		Echo "<tr class=""a4"">"
		Echo "	<td><b>操作:</b></td>"
		Echo "	<td><input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked> 置顶帖子 &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;  "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> 解除置顶 "
		Echo "	</td>"
		Echo "</tr>	 "
		Echo "<tr class=""a4"">"
		Echo "	<td>模式:</td>"
		Echo "	<td><input type=""radio"" name=""istohead"" value=""1"" class=""radio"" checked> 本版置顶 &nbsp; &nbsp; <input type=""radio"" name=""istohead"" value=""2"" class=""radio""> 总置顶 "
		Echo "</td>"
		Echo "</tr> "
	End If
	If Request("action")="settops" or Request("action")="digest" Then 
		Echo "<tr class=a4>"
		Echo "	<td><b>高亮设置:</b></td>"
		Echo "	<td>"
		Echo "	<input type=""radio"" name=""icolor"" value=""999"" class=""radio"" checked> 默认设置 &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""icolor"" value=""1"" class=""radio""> <span style=""width:10px;background:#808080;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""2"" class=""radio""> <span style=""width:10px;background:#808000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""3"" class=""radio""> <span style=""width:10px;background:#008000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""4"" class=""radio""> <span style=""width:10px;background:#0000ff;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""5"" class=""radio""> <span style=""width:10px;background:#800000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""6"" class=""radio""> <span style=""width:10px;background:#ff0000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""7"" class=""radio""> <span style=""width:10px;background:#cc0066;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""0"" class=""radio""> 取消高亮 &nbsp; &nbsp; "
		Echo "</td>"
		Echo "</tr>"
		Echo "<tr class=a4>"
		Echo "	<td><b>后续操作:</b></td>"
		Echo "	<td>"
		If Request("action")="digest" Then
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""1"" class=""radio"" checked> 加为精华帖子 &nbsp; &nbsp; "
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""2"" class=""radio""> 解除精华"
		Else
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""0"" class=""radio"" checked> 无设置 &nbsp; &nbsp; "
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""1"" class=""radio""> 加为精华帖子 &nbsp; &nbsp; "
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""2"" class=""radio""> 解除精华"
		End if
		Echo "</td>"
		Echo "</tr>"
	End If
	If Request("action")="isclosepage" Then
		Echo "<tr class=""a4"">"
		Echo "	<td><b>操作:</b></td>"
		Echo "	<td><input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked>  关闭主题 &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> 打开主题"
		Echo "	</td>"
		Echo "</tr> "
	End If
	If Request("action")="islockpage" Then
		Echo "<tr class=""a4""> "
		Echo "		<td><b>操作:</b></td> "
		Echo "	<td><input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked> 锁定帖子 &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> 解除锁定帖子"
		Echo "	</td>"
		Echo "</tr>"
	End If
	If Request("action")="move" Then
		Echo " <tr class=""a4""> "
		Echo "		<td><b>目标论坛/分类:</b></td><td>"
		Echo "		<select name=moveid> "
		Echo "		<option selected value="""">将主题移动到...</option>"
		Echo			team.BBs_Value_List(0,0)
		Echo "		</Select></td> "
		Echo " </tr>"
	End If
	If Request("action")="getlike" Then
		Dim Special,utmp,u
		Special = ""
		If Int(Board_Setting(15))=1 and Int(Board_Setting(17))=1 Then
			If Instr(Board_Setting(19),Chr(13)&Chr(10))>0 Then
				utmp = Split(Board_Setting(19),Chr(13)&Chr(10))
				For U=0 To Ubound(utmp)
					Special = Special &" <option value="""&U&""">"& utmp(u) &"</option>" 
				Next
			Else
				Special = "<option value=""0"">"& Board_Setting(19) &"</option>"
			End if	
		End If
		Echo " <tr class=""a4""> "
		Echo "	<td>批量设置主题分类:</td><td>"
		Echo "	<input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked> 加入分类 &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> 移出分类 "
		Echo "	</td> "
		Echo " </tr>"
		Echo " <tr class=""a4""> "
		Echo "		<td>批量设置主题分类:</td><td>"
		Echo "		<select name=posttopic><option value=""999"">请选择以下分类专题</option>"
		Echo		Special
		Echo "		</Select></td> "
		Echo " </tr>"
	End If
	Echo "<tr class=""a4"">"
	Echo "<td><B>操作原因:</B><BR> 版主及以下等级进行管理必须填写操作原因，超版以上等级可以无需填写操作原因。<BR> "
	If Not team.IsMaster Then Echo " <B>您必填填写才可以进行管理操作。</B> "
	Echo " <br><input type=""checkbox"" name=""sendpm"" value=""1"" class=""radio"" "
	If request("action")="movenew" Then 
		Echo "disabled"
	End  If
	If Not team.IsMaster Then
		Echo " checked "
	End if
	Echo "> 发短消息通知作者"
	Echo "</td><td>"
	Echo "<select name=""selectreason"" size=""6"" style=""height: 8em; width: 8em"" onchange=""this.form.reason.value=this.value"">"
	Echo"<option value="""">自定义</option>"
	Dim KillInfo,i
	If Instr(team.Club_Class(8),Chr(13)&Chr(10))>0 Then
		KillInfo = Split(team.Club_Class(8),Chr(13)&Chr(10))
		For i = 0 To Ubound(KillInfo)
			Echo "<option value="""&KillInfo(i)&""">"&KillInfo(i)&"</option>"
		Next
	Else
		Echo "<option value="""&team.Club_Class(8)&""">"&team.Club_Class(8)&"</option>"
	End if	
	Echo " </select> "
	Echo " <textarea name=""reason"" style=""height: 8em; width: 18em""></textarea></td>"
	Echo " </tr>"
	if Request("action")<>"movenew" Then
		Echo "<tr class=a4>"
		Echo "<td><B>用户操作:</B><BR>将操作值设置为负数，即扣除积分，相反操作值为正数，则为加分。 </td>"
		Echo "<td>"
		Dim ExtCredits,m,ExtSort
		ExtCredits = Split(team.Club_Class(21),"|")
		For m = 0 To UBound(ExtCredits)
			ExtSort=Split(ExtCredits(M),",")
			If Split(ExtCredits(M),",")(3)=1 Then
				Echo ExtSort(0) & " <select name=""ExtCredits"&M&"""  size=""1""> "
				Call Kills
				Echo "</select> &nbsp;"
			End if
		Next
		Echo "</td>"
		Echo "</tr>"
		Echo "<tr class=a4>"
		Echo "<td><B>追加扣分:</B><BR>最佳扣分及在默认扣分的前提下，再次追加扣分数量，用于对某用户的积分进行奖罚。 </td>"
		Echo "<td><input type=""radio"" name=""douser"" value=""0"" class=""radio"" checked>默认扣分 &nbsp; <input type=""radio"" name=""douser"" value=""1"" class=""radio"">追加扣分"
		Echo "</td></tr>"
	End if
	Echo "	 </table><br /><center><input type=""submit"" value="" 确 定 ""><br /> "
	Dim Rs2,SQL2,ho,rs1
	If request("rid")="" or (Not isnumeric(Request("rid"))) Then
		Echo "<br /><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=a2><tr class=a1><td>标题</td><td>作者</td><td>回复</td><td>最后发表</td></tr>"
		for each ho in request.form("ismanage")
			Set Rs2=Team.Execute("select id,topic,UserName,Replies,Lasttime from ["&IsForum&"Forum] where id="&CID(ho))
			If Not Rs2.Eof Then
				Echo "<tr class=a4><td><input type=checkbox name=ismanages value="&rs2(0)&" checked><a href=Thread.asp?tid="&rs2(0)&" target=_blank>"&Rs2(1)&"</a></td><td>"&Rs2(2)&"</td><td>"&Rs2(3)&"</td><td>"&Rs2(4)&"</td></tr>"
			End if
			Rs2.Close:Set Rs2=Nothing
		next
		Echo "</table></form>"
	Else
		If Request("menu")="move" and (Request.Form("fismanage")="" or not isnumeric(Request.Form("fismanage"))) Then
			team.error " 您只能对主题进行移动操作，请选定主题ID。"
		End If
		if isnumeric(Request.Form("fismanage")) and Request.Form("fismanage")<>"" Then
			Response.Write "<br /><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=a2><tr class=a1><td width=""30%"">标题</td><td width=""10%"">作者</td><td width=""10%"">回复</td><td width=""20%"">最后发表</td><td width=""10%"">用户IP</td></tr>"
			Set Rs2=Team.Execute("select id,topic,UserName,Replies,Lasttime,postip from ["&IsForum&"Forum] where id="& CID(Request.Form("fismanage")) )
			If Not Rs2.Eof Then
				Response.Write "<tr class=a4><td><input type=checkbox name=ismanages value="&rs2(0)&" checked><a href=Thread.asp?tid="&rs2(0)&" target=_blank>"&Rs2(1)&"</a></td><td>"&Rs2(2)&"</td><td>"&Rs2(3)&"</td><td>"&Rs2(4)&"</td><td>"&Rs2(5)&"</td></tr>"
			End if
			Rs2.Close:Set Rs2=Nothing
			Response.Write "</table>"
		Else
			Set Rs1=Team.Execute("select Relist,id,topic from ["&IsForum&"Forum] where id="&CID(request("rid")))
			If Not Rs1.Eof Then
				Echo "<br /><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=a2><tr class=a1><td>主贴</td><td>回帖ID</td><td>作者</td><td>回复时间</td><td>用户IP</td></tr>"
				Echo "<input type=""hidden"" value="&Rs1(0)&" name=relistname>"
				for each ho in request.form("ismanage")
					Set Rs2=Team.Execute("select id,UserName,posttime,postip from ["&IsForum & Rs1(0)&"] where id="&CID(ho))
					If Not Rs2.Eof Then
						Echo "<tr class=a4><td><input type=checkbox name=ismanages value="&rs2(0)&" checked> <a href=Thread.asp?tid="&rs1(1)&"#"&Rs2(0)&" target=_blank>"&Rs1(2)&"</a></td><td>"&Rs2(0)&"</a></td><td>"&Rs2(1)&"</td><td>"&Rs2(2)&"</td><td>"&Rs2(3)&"</td></tr>"
					End if
					Rs2.Close:Set Rs2=Nothing
				Next
				Echo "</table></form>"
			End If
			Rs1.Close:Set Rs1=Nothing
		End if
	End If
End Sub

Sub getlikes()
	Dim ho,rs,posttopic
	posttopic = HRF(1,2,"posttopic")
	If CID(team.Group_Manage(12)) = 1 then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select username,topic from ["&IsForum&"forum] where id="&CID(ho))
			If Not Rs.BOF then
				If request.form("isclose")=0 Then
					Team.execute("update ["&IsForum&"forum] set PostClass="&Int(posttopic)&" where id="&CID(ho))
					Mstemp = "批量设置主题分类 "
					Call delpiont(RS(0))
					GInfo = "批量设置主题分类"
					Call Pmsetto(rs(0),rs(1))
					temp="批量设置主题分类成功"
				Else
					Team.execute("update ["&IsForum&"forum] set PostClass=999 where id="&CID(ho))
					temp="批量移出主题分类成功"
					Mstemp =  "批量移出主题分类"
					Call delpiont(RS(0))
					GInfo = "批量移出主题分类"
					Call Pmsetto(rs(0),Rs(1))
				End if
			End if
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error("<li>您所在的组 "&team.Levelname(0)&" 没有批量设置主题分类的权限")
	End if
	Call Serverend(Mstemp)
End Sub

Sub Kills()
	Dim i,t
	For i = -50 To -1
		t = t &"<option value="""&i&""">"&i&"</option>"
	Next
	t = t &"<option value=""0"" selected>0</option>"
	For i = 1 To 50
		t = t &"<option value="""&i&""">"&i&"</option>"
	Next
	Echo t
End Sub

Sub digests
	Dim ho,rs,iSetColor
	If team.Group_Manage(8) = "1" then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select username,id,topic from ["&IsForum&"forum] where id="&CID(ho))
			If Not Rs.BOF Then
				If Request.Form("icolor") <> 999 Then
					iSetColor = ",color=" & CID(Request.Form("icolor"))
				End if
				If request.form("togoodtopic")=1 Then
					team.execute("update ["&IsForum&"forum] set goodtopic=1"& iSetColor &" where id="&CID(ho))
					team.execute("update ["&IsForum&"user] set goodtopic=goodtopic+1 where username='"&RS(0)&"'")
					UpdateUserpostExc RS(0),2
					Mstemp =  "用户"&tk_UserName&"将主题[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]加入精华区。"
					Temp = "帖子加入精华区成功 !"
					GInfo = "加精"
				ElseIf request.form("togoodtopic")=2 Then
					team.execute("update ["&IsForum&"forum] set goodtopic=0"& iSetColor &" where id="&CID(ho))
					team.execute("update ["&IsForum&"user] set goodtopic=goodtopic-1 where username='"&RS(0)&"'")
					KillUpdateUserpostExc RS(0),2
					Mstemp = "用户"&tk_UserName&"将主题[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]移出精华区。"
					Temp = "帖子移出精华区成功 !"
					GInfo = "取消精华"
				End If
			End if
			Call delpiont(RS(0))
			Call Pmsetto(RS(0),RS(2))
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error " 您所在的组 "&team.Levelname(0)&" 没有加入/解除精华主题的权限。" 
	End if
	Call Serverend(Mstemp)
End Sub

Sub UpdateUserpostExc(uname,m)
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
				UExt = UExt &"Extcredits0=Extcredits0+"&MustSort(m)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"+"&MustSort(m)&""
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where UserName = '"& HTmlEncode(uname) &"' ")
End Sub

Sub KillUpdateUserpostExc(uname,m)
	'用户积分部分
	Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u,Rs
	Dim UserPostID,My_ExtSort
	If Not team.UserLoginED Then  Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	UExt = ""
	Set Rs = team.execute("Select Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7 From ["&IsForum&"User] Where UserName = '"& HTmlEncode(uname) &"' ")
	If Rs.Eof And Rs.Bof Then
		Exit Sub
	Else 
		For U=0 to Ubound(ExtCredits)
			ExtSort=Split(ExtCredits(U),",")
			MustSort=Split(MustOpen(U),",")
			If ExtSort(3)=1 Then
				If U = 0 Then
					If (Rs(0+U)-MustSort(4))-MustSort(8)<=0 Then
						UExt = UExt &"Extcredits0=0"
					Else
						UExt = UExt &"Extcredits0=Extcredits0-"&MustSort(m)&""
					End if
				Else
					If (Rs(0+U)-MustSort(4))-MustSort(8)<=0 Then
						UExt = UExt &",Extcredits"&U&"=0"
					Else
						UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"-"&MustSort(m)&""
					End if
				End If
			End If
		Next
		team.execute("Update ["&IsForum&"User] Set "&UExt&" Where UserName = '"& HTmlEncode(uname) &"' ")
	End If 
End Sub

Sub settopss
	Dim ho,rs,iSetColor
	If CID(team.Group_Manage(1)) >= 1 then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select Username,ID,Topic,Goodtopic from ["&IsForum&"forum] Where ID="&CID(ho))
			If Not Rs.BOF Then
				If Request.Form("icolor") <> 999 Then
					iSetColor = ",color=" & CID(Request.Form("icolor"))
				End if
				If request.form("isclose")=0 Then
					If request.form("istohead")=1 and CID(team.Group_Manage(1)) >= 1 Then
						team.execute("update ["&IsForum&"forum] set toptopic=1"& iSetColor &" where id="&CID(ho))
						Mstemp = "用户"&tk_UserName&"将主题 [<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]进行了本版置顶操作。 "
						GInfo = "本版置顶"
					Elseif request.form("istohead")=2 and CID(team.Group_Manage(1)) = 2 Then
						team.execute("update ["&IsForum&"forum] set toptopic=2"& iSetColor &" where id="&CID(ho))
						Mstemp = "用户"&tk_UserName&"将主题 [<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]进行了总置顶操作。 "
						GInfo = "总置顶"
					End If
					Call delpiont(RS(0))
					temp="设置主题置顶成功。"
				Else
					If CID(team.Group_Manage(1)) >= 1 Then
						team.execute("update ["&IsForum&"forum] set toptopic=0"& iSetColor &" where id="&CID(ho))
						Mstemp = "用户"&tk_UserName&"将置顶的主题 [<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]设置为正常帖子状态。"
						Call delpiont(RS(0))
						temp="解除主题置顶成功。"
						GInfo = "解除置顶"
					End if
				End if
				If request.form("togoodtopic")=1 and team.Group_Manage(8) = 1 Then
					If Rs(3) = 0 Then
						team.execute("update ["&IsForum&"forum] set goodtopic=1"& iSetColor &" where id="&CID(ho))
						team.execute("update ["&IsForum&"user] set goodtopic=goodtopic+1 where username='"&RS(0)&"'")
						Mstemp = Mstemp & "<br>用户"&tk_UserName&"将主题[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]加入精华区。"
					End if
				ElseIf request.form("togoodtopic")=2 and team.Group_Manage(8) = 1 Then
					If Rs(3) = 1 Then
						team.execute("update ["&IsForum&"forum] set goodtopic=0"& iSetColor &" where id="&CID(ho))
						team.execute("update ["&IsForum&"user] set goodtopic=goodtopic-1 where username='"&RS(0)&"'")
						Mstemp = Mstemp & "<br>用户"&tk_UserName&"将主题[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]移出精华区。"
					End if
				End If
				Call Pmsetto(rs(0),Rs(2))
			End if
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error "您所在的组 "&team.Levelname(0)&" 没有加入/解除精华主题的权限"
	End if
	Call Serverend(Mstemp)
End Sub

Sub isclosepages
	Dim ho,rs
	If team.Group_Manage(7) = 1 then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select username,topic from ["&IsForum&"forum] where id="&CID(ho))
			If Not Rs.BOF then
				If request.form("isclose")=0 Then
					Team.execute("update ["&IsForum&"forum] set CloseTopic=1 where id="&CID(ho))
					Mstemp = "关闭主题  [url=Thread.asp?tid="&CID(ho)&"]"&RS(1)&"[/url] "
					Call delpiont(RS(0))
					GInfo = "关闭主题"
					Call Pmsetto(rs(0),rs(1))
					temp="关闭主题成功"
				Else
					Team.execute("update ["&IsForum&"forum] set CloseTopic=0 where id="&CID(ho))
					temp="打开被关闭的主题成功"
					Mstemp =  "打开被关闭主题  [url=Thread.asp?tid="&CID(ho)&"]"&RS(1)&"[/url]"
					Call delpiont(RS(0))
					GInfo = "打开关闭主题"
					Call Pmsetto(rs(0),Rs(1))
				End if
			End if
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error("<li>您所在的组 "&team.Levelname(0)&" 没有关闭/打开主题的权限")
	End if
	Call Serverend(Mstemp)
End Sub

Sub movenews
	Dim ho
	If team.Group_Manage(5) = 1 then
		for each ho in request.form("ismanages")
			team.execute("update ["&IsForum&"forum] set Lasttime="&SqlNowString&" Where id="&CID(ho))
			temp="拉前主题成功"
			Mstemp =  "拉前主题"
		next
	Else
		Team.Error "您所在的组 "&team.Levelname(0)&" 没有拉前主题的权限"
	End if
	Call Serverend(Mstemp)
End Sub

Sub islockpages
	Dim ho,rs,RName
	If team.Group_Manage(6) = 1 then
		for each ho in request.form("ismanages")
			If request.form("isclose")=0 Then
				if isnumeric(Request("rid")) and Request("rid")<>"" and (Request.Form("fismanage")="" or not isnumeric(Request.Form("fismanage"))) then
					set rs=team.execute("Select ReList,username,Topic from ["&IsForum&"forum] where id="&Request("rid"))
					If Not Rs.BOF then
						team.execute("update ["&IsForum & rs(0)&"] set lock=1 where topicid="&CID(Request("rid"))&" and id="&CID(ho))
						RName = team.execute("select username from ["&IsForum & rs(0)&"] where id="&CID(ho))(0)
						Mstemp =  "锁定回贴"
						Call delpiont(RName)
						GInfo = "锁贴"
						Call Pmsetto(RName,RS(2))
						Temp="锁定回贴成功。"
					end if
					Rs.Close:Set Rs=Nothing
				else
					set rs=team.execute("Select username,topic from ["&IsForum&"forum] where id="&CID(ho))
					If Not Rs.BOF then
						team.execute("update ["&IsForum&"forum] set Locktopic = 1 where id="&CID(ho))
						Mstemp =  "锁定主题 [url=Thread.asp?tid="&CID(ho)&"]"&RS(1)&"[/url]"
						Call delpiont(RS(0))
						GInfo = "锁贴"
						Call Pmsetto(RS(0),RS(1))
					End If
					Rs.Close:Set Rs=Nothing
					Temp="锁定主题成功。"
				end if
			Else
				if isnumeric(Request("rid")) and Request("rid")<>"" then
					set rs=team.execute("Select ReList,username,Topic from ["&IsForum&"forum] where id="&Request("rid"))
					If Not Rs.BOF then
						team.execute("update ["&IsForum & rs(0)&"] set lock=0 where topicid="&CID(Request("rid"))&" and id="&CID(ho))
						Mstemp =  "解除回贴锁定"
						GInfo = "锁贴"
						Call delpiont(team.execute("select username from ["&IsForum & rs(0)&"] where id="&CID(ho))(0))
						Call Pmsetto(team.execute("select username from ["&IsForum & rs(0)&"] where id="&CID(ho))(0),RS(2))
						Temp="解除回贴锁定成功。"
					end If
					Rs.Close:Set Rs=Nothing
				else
					set rs=team.execute("Select username,Topic from ["&IsForum&"forum] where id="&CID(ho))
					If Not Rs.BOF then
						team.execute("update ["&IsForum&"forum] set Locktopic =0 where id="&CID(ho))
						Mstemp =  "解除主题锁定"
						Call delpiont(RS(0))
						GInfo = "锁贴"
						Call Pmsetto(RS(0),RS(1))
					End If
					Rs.Close:Set Rs=Nothing
					Temp="解除主题锁定成功。"
				end if
			End If
		next
		Application.Contents.RemoveAll()
	Else
		Team.Error "您所在的组 "&team.Levelname(0)&" 没有锁定/解除帖子的权限"
	End if
	Call Serverend(Mstemp)
End Sub

Sub forummove()
	Dim Ts,UpID,ho,Rs,SQL
	Dim Board_Setting
	if team.Group_Manage(4) = 1 then
		if Request("moveid")="" then 
			team.Error "您没有选择要将主题移动哪个论坛!"
		End if
		If Request("moveid")=Request("fid") Then 
			team.Error "你选择的论坛和源论坛相同!"
		End If
		Board_Setting = team.Execute("Select Board_Setting From ["&IsForum&"bbsconfig] where ID="&CID(Request("moveid")))(0)
		if Split(Board_Setting,"$$$")(2) = 1 Then
			team.Error "目标论坛属于审核版块，不能转入。"
		End if
		for each ho in request.form("ismanages")
			Set Rs = team.execute("Select forumid,topic,Toptopic,Locktopic,Lasttime,UserName,ID from ["&IsForum&"forum] where id="&CID(ho))
			If Not (Rs.BOF and Rs.EOF) Then
				team.execute("Update ["&IsForum&"forum] set forumid="&int(Request("moveid"))&",topic='"&RS(1)&"',Toptopic=0,Locktopic=0,Lasttime="&SqlNowString&"  Where ID="&CID(ho))
				GInfo = "移动主题"
				Call delpiont(RS(5))
				Call Pmsetto(RS(5),Rs(1))
				Mstemp = "移动帖子ID : [url=Thread.asp?tid="&RS(6)&"]"&RS(1)&"[/url][BR]"
			End If
			Temp="移动主题成功"
			Rs.Close:Set Rs=Nothing
			UpID = Team.Execute("Select Max(ID) From ["&IsForum&"Forum] Where deltopic=0 and forumid="& Request("fid"))(0)
			set Ts=team.execute("select top 1 topic,Lasttime,username,ID from ["&IsForum&"forum] where ID="& UpID )
			If Not Ts.Eof Then
				team.execute("update ["&IsForum&"bbsconfig] set Board_Last='<A href=Thread.asp?tid="&TS(3)&" target=""_blank"">"&Cutstr(TS(0),200)&"</a> →$@$"&TS(2)&"$@$"&Now()&"' where id="&Request("fid"))
			End If
			Ts.Close:Set Ts=Nothing
		Next
		Cache.DelCache("BoardLists")
	Else
		team.Error " 您所在的组 "&team.Levelname(0)&" 没有移动主题的权限 "
	End if
	Call Serverend(Mstemp)
End Sub

Sub deltopics
	if team.Group_Manage(3) = 1 then
		Dim Forum_ID,Max_ID,rs1,ho,rs,Isnames,DayDel
		for each ho in request.form("ismanages")
			if Request("rid")<>"" and isnumeric(Request("rid")) And (Request.Form("fismanage")="" or not isnumeric(Request.Form("fismanage"))) then
				set rs=team.execute("Select forumid,ReList,Topic,Posttime from ["&IsForum&"forum] where id="&Request("rid"))
				If Not Rs.BOF then
					Isnames= team.execute("select username from ["&IsForum & rs(1)&"] where id="&CID(ho))(0)
					Call delpiont(Isnames)
					GInfo = "删除回贴"
					Call Pmsetto(Isnames,RS(2))
					KillUpdateUserpostExc Isnames,1
					team.execute("delete from ["&IsForum & rs(1)&"] where id="&CID(ho))
					team.execute("update ["&IsForum&"forum] set replies=replies-1 where id="&CID(Request("rid")))
					If DateDiff("d",RS(3),Date())=0 Then
						DayDel = "today=today-1,"
					End If
					team.execute("update ["&IsForum&"bbsconfig] set "& DayDel &"tolrestore=tolrestore-1 where id="&rs(0))
					team.execute("update ["&IsForum&"user] set postrevert=postrevert-1 where username='"&Isnames&"'")
					If CID(DateDiff("d",CDate(RS(3)),Now()))=1 Then
						team.LockCache "TodayNum" , Application(CacheName&"_TodayNum")-1
					End If
				End If
				Temp="删除回贴成功"
				Mstemp = "删除回贴"
				Rs.close:Set Rs=Nothing
			Else
				'If Request.Form("fismanage")<>"" And IsNumeric(Request.Form("fismanage")) Then
					'set rs=team.execute("Select forumid,topic,toptopic,goodtopic,Locktopic,lasttime,UserName,id,ReList from ["&IsForum&"forum] where id="& Int(Request.Form("fismanage")))
				'Else
				set rs=team.execute("Select forumid,topic,toptopic,goodtopic,Locktopic,lasttime,UserName,id,ReList from ["&IsForum&"forum] where id="&CID(ho))
				'End if
				If Not Rs.BOF then
					team.execute("update ["&IsForum&"user] set deltopic=deltopic+1 where username='"&rs(7)&"'")
					team.execute("update ["&IsForum&"forum] set toptopic=0,deltopic=1,lasttime="&SqlNowString&",LastText='"&tk_UserName &"$@$此贴已被"&tk_UserName &"删除' where deltopic=0 and id="&CID(ho))
					'处理其他表
					team.execute("delete from ["&IsForum &"FVote] where RootID="&CID(ho))
					team.execute("delete from ["&Isforum&"Activity] where RootID="&CID(ho))
					team.execute("delete from ["&Isforum&"ReActivity] where RootID="&CID(ho))
					team.execute("delete from ["&Isforum&"ActivityUser] where RootID="&CID(ho))
					If DateDiff("d",RS(3),Date())=0 Then
						DayDel = "today=today-1,"
					End If
					Max_ID=Team.Execute("Select Max(ID) from ["&IsForum&"forum] where deltopic=0 and Forumid="&rs(0))(0)
					If Max_ID<>"" Then
						Set Rs1=Team.Execute("Select ID,topic,username,posttime from ["&IsForum&"forum] where deltopic=0 and id="&Max_ID)
						if Not rs1.eof then
							team.execute("update ["&IsForum&"bbsconfig] set "&DayDel&"toltopic=toltopic-1,Board_Last='<A href=Thread.asp?tid="&rs1(0)&" target=""_blank"">"&Cutstr(rs1(1),200)&"</a> →$@$"&rs1(2)&"$@$"&Now()&"' where id="&rs(0))
						End If
						Rs1.Close:Set Rs1 = Nothing
					Else
						team.execute("update ["&IsForum&"bbsconfig] set "&DayDel&"toltopic=toltopic-1,Board_Last='暂无帖子$@$"&TK_UserName&"$@$"&Now()&"' where id="&rs(0))
					End If

					Call delpiont(RS(6))
					GInfo = "删除主题 "
					Call Pmsetto(RS(6),RS(1))
					KillUpdateUserpostExc RS(6),0
				End If
				Temp = "删除主题成功"
				Mstemp =  Mstemp & " 删除主题 : ["& Rs(1) &"]"
			End If
		Next
		Cache.DelCache("BoardLists")
	Else
		Team.Error "您所在的组 "&team.Levelname(0)&" 没有删除帖子的权限"
	End if
	Call Serverend(Mstemp)
End Sub

Sub Serverend(s)
	team.SaveLOG("用户"&TK_UserName&"操作: "&s)
	team.Error1 ("<li>"&temp&"<li><a href=""Forums.asp?fid="&request("fid")&""">返回论坛</a><li><a href=""Default.asp"">返回论坛首页</a><meta http-equiv=refresh content=3;url=""Forums.asp?fid="&request("fid")&""">")
End Sub

Sub delpiont(s)
	Dim ExtCredits,m,ExtSort,GetMyExs,ExcName,MustOpen,MustSort,Rs
	If HRF(1,2,"douser") = 0 Then Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	GetMyExs=""
	Set Rs = team.execute("Select Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7 From ["&IsForum&"User] Where UserName = '"& HTmlEncode(s) &"' ")
	If Rs.Eof And Rs.Bof Then
		team.Error "不存在此用户"
	Else 
		For m = 0 To UBound(ExtCredits)
			ExtSort=Split(ExtCredits(M),",")
			MustSort=Split(MustOpen(m),",")
			If Split(ExtCredits(M),",")(3)=1 Then
				If Request.Form("ExtCredits"&M) <> "0" Then
					If GetMyExs = "" Then
						If (Rs(0+M)-Request.Form("ExtCredits"&M))-MustSort(8)<=0 Then
							GetMyExs = "ExtCredits"& M& "=0"
						Else
							GetMyExs = "ExtCredits"& M& "=ExtCredits"& M& "+"& Request.Form("ExtCredits"&M)
						End if
					Else
						If (Rs(0+M)-Request.Form("ExtCredits"&M))-MustSort(8)<=0 Then
							GetMyExs = GetMyExs & ",ExtCredits"& M& "=0"
						Else
							GetMyExs = GetMyExs & ",ExtCredits"& M& "=" & "ExtCredits"& M& "+ "& Request.Form("ExtCredits"&M)
						End if
					End If
					ExcName = ExcName & ExtSort(0) &" : "& Request.Form("ExtCredits"&M)
				End If
			End If
		Next
	End If 
	If GetMyExs <>"" Then
		if s=TK_UserName then 
			team.Error "你不能对自己进行操作!"
		Else
			team.execute("Update ["&IsForum&"User] Set "& GetMyExs &" Where UserName='"&Htmlencode(s)&"' ")
		End If
		temp = temp & "<br>" & ExcName
	End if
End Sub
Sub Pmsetto(s,m)
	If request("sendpm") = "1" Then 
		Dim Istemp,ho
		If Not team.IsMaster and len(request("reason"))<2 Then 
			team.error2 "你没有填写操作原因"
		Else
			Istemp = "这是由论坛系统自动发送的通知短消息。[br] "
			If request("rid")<>"" or isnumeric(Request("rid")) Then 
				Istemp = Istemp & " 您在主题：[url=Thread.asp?tid="&request("rid")&"]"&m&"[/url] 的回复帖子 [br] "
			Else
				Istemp = Istemp & " 您所发表的主题： "
				for each ho in request.form("ismanages")
					Istemp = Istemp & "  [url=Thread.asp?tid="&CID(ho)&"] "&m&"[/url] [br]"
				Next
			End if
			Istemp = Istemp & " 被 "&tk_UserName&" 执行 "& GInfo  &" 操作 [br] 操作理由:  "&request("reason")&" 。"
		End If
		Team.Execute("insert into ["&IsForum&"message](author,incept,content,Sendtime,MsgTopic) values ('"&TK_UserName&"','"&HTmlEncode(s)&"','"&HTmlEncode(Istemp)&"',"&SqlNowString&",'[系统消息]您发表的帖子被执行管理操作!')")
		Team.Execute("update ["&IsForum&"user] set newmessage=newmessage+1 where username='"&HtmlEncode(s)&"'")
	End If
End Sub

Sub ManageClass()
	Dim Rs
	team.ChkPost()
	Set Rs = team.Execute("Select ID,bbsname,Board_Setting From ["&IsForum&"bbsconfig] Where ID="&fID)
	If Rs.Eof Then
		team.Error " 参数错误。"
	Else
		Values = Rs.GetRows(-1)
	End If
	If isarray(Values) Then
		Board_Setting = Split(Values(2,0),"$$$")
	End if
	team.Headers("论坛帖子管理中心 - "& Values(1,0))
	x1="<a href=""Forums.asp?fid="&fID&""">"& Values(1,0)  &"</a> "
	x2=" 论坛帖子管理中心 "
	Echo team.MenuTitle
	If Not team.UserLoginED Then
		team.Error " 你未登陆论坛。<meta http-equiv=refresh content=3;url=login.asp> "
	End if
	If Not team.ManageUser Then
		team.Error " 您的权限不够，不能参与论坛管理 。"
	Else
		If Not team.IsMaster and Not team.SuperMaster Then
			If team.execute("Select ID from ["&Isforum&"Moderators] Where ManageUser='"& tk_username &"' and BoardID = "& fid).eof Then 
				team.Error " 您不是此版的版主,不能参与此版的管理"
			End If
		End if
	End if
End Sub
Team.footer
%>