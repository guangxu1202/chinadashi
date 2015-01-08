<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim tID,fID,x1,x2,MyNews,WebUrl
fID = HRF(2,2,"fid")
WebUrl = team.Club_Class(2)
Set MyNews = New ClNews

Class ClNews
	Private Sub Class_Initialize()
		If CID(team.Forum_setting(85)) = 0 Then
			Call NotCallNews
		End If
		If Not CheckServer() then
			OutNews "数据被保护,禁止被其他站点调用!"
			Response.End	
		End If
		team.LoadTemps
		Select Case Request("action")
			'========调用类型======
			Case "recallaffiche"
				Call recallaffiche
			Case "recallmembers"
				Call recallmembers
			Case "recallinfos"
				Call recallinfos
			Case "recallboardfids"
				Call recallboardfids
			Case "recallshowlinks"
				Call recallshowlinks
			'========结束==========
			Case "Next1"
				Call Next1
			Case "Next2"
				Call Next2
			Case Else
				Call Main()
		End Select
	End Sub

	Sub recallshowlinks
		Dim tmp,myclass
		myclass = HRF(2,2,"myclass")
		tmp = Stringhtml(Team.HtmlNews (2))
		tmp = Replace(tmp,"{$myclass}",BBs_Value_List(0,0,myclass))
		OutNews Fixjs(tmp)
	End Sub

	Function BBs_Value_List(a,b,c)
		Dim tmp,i,Boards
		Dim U,Y
		Boards = team.myBoardJump()
		If isArray(Boards) Then
			For i=0 To Ubound(Boards,2)
				If Boards(2,i) = a Then
					U = 1+b
					For Y=0 To U
						tmp = tmp & "&nbsp; &nbsp;"
					Next
					If a = 0 Then tmp = tmp & "<br />"
					If C = 0 Then
						If a = 0 Then
							tmp = tmp & "╋"
						Else
							tmp = tmp & "├"
						End If
					End If 
					tmp = tmp & "<a href="""& WebUrl &"/Forums.asp?fid="& Boards(0,i)&""" target=""_blank"">"& Boards(1,i)&"</a>"& Vbcrlf
					If C = 0 Then
						tmp = tmp & "<br />" & Vbcrlf
					Else
						tmp = tmp & " &nbsp;" & Vbcrlf
						If a = 0 Then tmp = tmp & "<br />"
					End if	
					tmp = tmp & BBs_Value_List(Boards(0,i),U,C) 
				End if
			Next
		End if
		BBs_Value_List = tmp
	End function

	Sub recallboardfids
		Dim tmp,psize,formattime,showboards,showorders,icon,showtypes,total,showname
		Dim Rs,t,i,SQL,Orders,ts,u,m,p
		icon = HRF(2,2,"icon")
		psize = HRF(2,2,"psize")
		total = HRF(2,2,"total")
		showname = HRF(2,2,"showname")
		showtypes = HRF(2,2,"showtypes")
		formattime = HRF(2,2,"formattime")
		showboards = HRF(2,2,"showboards")
		showorders = HRF(2,2,"showorders")
		SQL = "" : Orders=""
		t = team.BoardList
		If showboards > 0 Then
			SQL = " and forumid = " & Int(showboards)
		End If
		If showtypes = 0 Then
			SQL = " and goodtopic = 1 "
		End If
		Select Case showorders
			Case "1"
				Orders = " PostTime "
			Case "2"
				Orders = " Lasttime "
			Case "3"
				Orders = " Views "
			Case "4"
				Orders = " Replies "
			Case Else
				Orders = " ID "
		End Select
		Set Rs = team.execute("Select ID,Topic,UserName,Views,Replies,Posttime,ICON,Forumid From ["&Isforum&"Forum] Where deltopic=0 "& SQL &" Order By "& Orders &" Desc ")
		If Not Rs.Eof Then
			ts = Rs.GetRows(total)
		End If
		Rs.Close:Set Rs=Nothing
		If IsArray(ts) Then
			m = "" : P=""
			For u = 0 To UBound(ts,2)
				If INT(icon) = 1 Then
					If Int(ts(6,u))>0 Then
						m = "<img src="""&WebUrl&"/images/brow/icon"&ts(6,u)&".gif"" align=""absmiddle"" border=""0""> "
					End if
				End If 
				If INT(showname) = 1 Then
					If IsArray(t) Then
						For I = 0 To UBound(t,2)
							If Int(t(0,i)) = Int(ts(7,u)) Then
								p = " <a href="""&WebUrl&"/Forums.asp?fid="&t(0,i)&""" target=""_blank"">["& t(1,i) &"]</a> " 
							End if
						Next
					End If
				End If
				tmp = tmp & Stringhtml(Team.HtmlNews (1))
				tmp = Replace(tmp,"{$class}",m)
				tmp = Replace(tmp,"{$show}",p)
				tmp = Replace(tmp,"{$weburl}",WebUrl)
				tmp = Replace(tmp,"{$tid}",ts(0,u))
				tmp = Replace(tmp,"{$title}",Cutstr(ts(1,u),psize))
				tmp = Replace(tmp,"{$infos}","发表用户："& ts(2,u) &"&#xA;查看数："& ts(3,u) &"&#xA;回复数："& ts(4,u) &"")
				tmp = Replace(tmp,"{$times}",FormatDateTime(ts(5,u),formattime))
			Next
		End If
		OutNews Fixjs(tmp)
	End Sub


	Sub recallinfos
		Dim tmp,psize,formattime,Onlinemany
		psize = HRF(2,2,"psize")
		formattime = HRF(2,2,"formattime")
		tmp = Stringhtml(Team.HtmlNews (0))
		Onlinemany = team.execute("Select Count(*) From ["&IsForum&"Online]")(0)
		tmp = Replace(tmp,"{$weburl}",WebUrl)
		tmp = Replace(tmp,"{$TopicNum}",Application(CacheName&"_PostNum"))
		tmp = Replace(tmp,"{$PostNum}",Application(CacheName&"_ConverPostNum"))
		tmp = Replace(tmp,"{$Regmember}",Application(CacheName&"_UserNum"))
		tmp = Replace(tmp,"{$AllOnline}",Onlinemany)
		tmp = Replace(tmp,"{$LastReg}",team.Club_Class(12))
		tmp = Replace(tmp,"{$TodayPostNum}",Application(CacheName&"_TodayNum"))
		tmp = Replace(tmp,"{$OLdPost}",Application(CacheName&"_OldTodayNum"))
		tmp = Replace(tmp,"{$TopOnline}",Split(Team.Club_Class(20),"|")(0))
		tmp = Replace(tmp,"{$StarDay}",FormatDateTime(team.Club_Class(29),formattime))
		OutNews Fixjs(tmp)
	End Sub

	Sub recallmembers
		Dim tmp,total,psize,userorder,t,i
		Dim ExtCredits,u,m,Byorders,Rs,S
		ExtCredits = Split(team.Club_Class(21),"|")
		total = HRF(2,2,"total")
		psize = HRF(2,2,"psize")
		userorder = CID(Request("orderys"))
		Select Case userorder
			Case "0"
				Byorders = "Regtime Desc"
			Case "1"
				Byorders = "Posttopic Desc"
			Case "2"
				Byorders = "Posttopic+Postrevert Desc"
			Case "3"
				Byorders = "Goodtopic Desc"
			Case "4"
				Byorders = "Extcredits0 Desc"
			Case "5"
				Byorders = "Extcredits1 Desc"
			Case "6"
				Byorders = "Extcredits2 Desc"
			Case "7"
				Byorders = "Extcredits3 Desc"
			Case "8"
				Byorders = "Extcredits4 Desc"
			Case "9"
				Byorders = "Extcredits5 Desc"
			Case "10"
				Byorders = "Extcredits6 Desc"
			Case "11"
				Byorders = "Extcredits7 Desc"
			Case Else
				Byorders = "Regtime Desc"
		End Select
    	Cache.Reloadtime = CID(team.Forum_setting(86))
		Cache.Name="recallmembers_" & userorder
		If Cache.ObjIsEmpty() Then
			Set Rs = team.execute("Select UserName,Posttopic,Postrevert,goodtopic,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7 From ["&IsForum&"User] Where not (UserGroupID=5) Order By "& Byorders&" ")
			If Not Rs.Eof Then
				t = Rs.GetRows(total)
				Cache.Value = t
			End If
			Rs.Close:Set Rs=Nothing 
		End If
		t = Cache.Value
		If IsArray(t) Then
			For i = 0 To UBound(t,2)
				If i =>total Then Exit For
				tmp = tmp & Stringhtml(Team.HtmlNews (3))
				tmp = Replace(tmp,"{$weburl}",WebUrl)
				tmp = Replace(tmp,"{$Csslist}",team.Styleurl)
				tmp = Replace(tmp,"{$username}",t(0,i))
				tmp = Replace(tmp,"{$title}",Cutstr(HtmlEncode(t(0,i)),psize))
				tmp = Replace(tmp,"{$postnum}",t(1,i))
				tmp = Replace(tmp,"{$repostnum}",t(2,i))
				tmp = Replace(tmp,"{$godnum}",t(3,i))
				s = ""
				for m = 0 to ubound(ExtCredits)
					If Split(ExtCredits(m),",")(4) =1 Then
						s = s & ""& Split(ExtCredits(m),",")(0) & "&nbsp;"& t(4+m,i) &"&nbsp;"& Split(ExtCredits(m),",")(1) &"&#xA;"
					End if
				Next
				tmp = Replace(tmp,"{$infos}",s)
			Next
		End If
		OutNews Fixjs(tmp)
	End Sub

	Sub recallaffiche
		Dim Rs,i,tmp,total,psize,formattime
		total = HRF(2,2,"total")
		psize = HRF(2,2,"psize")
		formattime = HRF(2,2,"formattime")
		Rs = team.Affiche
		If IsArray(Rs) Then
			For i = 0 To UBound(Rs,2)
				If i =>total Then Exit For 
				tmp = tmp & Stringhtml(Team.HtmlNews (4))
				tmp = Replace(tmp,"{$weburl}",WebUrl)
				tmp = Replace(tmp,"{$Csslist}",team.Styleurl)
				tmp = Replace(tmp,"{$fid}",RS(0,i))
				tmp = Replace(tmp,"{$title}",Cutstr(RS(1,i),psize))
				tmp = Replace(tmp,"{$infos}",FormatDateTime(RS(4,i),formattime))
			Next
		End If
		OutNews Fixjs(tmp)
	End Sub

	Private Sub Main()
		team.Headers(Team.Club_Class(1) & "- 论坛调用指南")
		Echo "<form action=""?action=Next1"" method=""post""><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
		Echo "<tr class=""tab1""><td> 您好，欢迎管理论坛调用向导 －－ 》，请首先选择您需要的调用内容。然后点击下一步。</td></tr>"
		Echo "	<tr class=""tab3""><td><SELECT NAME=""NewsType""> "
		Echo "		<option value=""0"">选取调用类型</option>"
		Echo "		<option value=""1"">论坛参数调用</option>"
		Echo "		<option value=""2"">帖子调用</option>"
		Echo "		<option value=""3"">版块调用</option>"
		Echo "		<option value=""4"">会员调用</option>"
		Echo "		<option value=""5"">公告调用</option>"
		Echo "	</SELECT> </td></tr>"
		Echo "	</table><BR><center><input type=""submit"" name=""submit"" value=""下 一 步""></center></form>"
		team.footer
	End Sub

	Private Sub Next1()
		Dim NewsType
		team.Headers(Team.Club_Class(1) & "- 论坛调用指南")
		NewsType = HRF(1,2,"NewsType")
		If NewsType = 0 Then
			team.Error "请输入调用类型。"
		End If
		Select Case NewsType
			Case "1"
				Echo "<form action=""?action=Next2"" method=""post""><input type=""hidden"" name=""myjs"" value=""1""><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
				Echo "<tr class=""tab1""><td colspan=""2""> 论坛参数调用</td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">标题长度：</td><td class=""altbg2""><input type=""text"" name=""psize"" size=""30"" value=""20""></td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">时间格式：</td><td class=""altbg2""> <select name=""formattime"" ID=""formatTime"">"
				Echo "	<option value=""0"" SELECTED>YYYY-M-D H:M:S(长格式)</option>"
				Echo "	<option value=""1"">YYYY年M月D</option>"
				Echo "	<option value=""2"">YYYY-M-D</option>"
				Echo "	<option value=""3"">H:M:S</option>"
				Echo "	<option value=""4"">hh:mm</option>"
				Echo "	</select> [服务器设置: "& Now &"] </td></tr>"
				Echo "	</table><BR><center><input type=""submit"" name=""submit"" value=""生成JS""></center></form>"
			Case "2"
				Echo "<form action=""?action=Next2"" method=""post""><input type=""hidden"" name=""myjs"" value=""2""><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
				Echo "<tr class=""tab1""><td colspan=""2""> 帖子调用</td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">心情图标：</td><td class=""altbg2""> <input type=""radio"" name=""icon"" value=""1"" class=""radio"" CHECKED> 是 &nbsp; &nbsp; <input type=""radio"" name=""icon"" value=""0"" class=""radio""> 否</td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">显示类型：</td><td class=""altbg2""> <input type=""radio"" name=""showtypes"" value=""1"" class=""radio"" CHECKED> 显示所有主题 &nbsp; &nbsp; <input type=""radio"" name=""showtypes"" value=""0"" class=""radio""> 只显示精华主题 </td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">显示版块名称：</td><td class=""altbg2""> <input type=""radio"" name=""showname"" value=""1"" class=""radio"" CHECKED> 是 &nbsp; &nbsp; <input type=""radio"" name=""showname"" value=""0"" class=""radio""> 否</td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">排序方式：</td><td class=""altbg2""> <select name=""showorders"">"
				Echo "	<option value=""0"" SELECTED>默认最新排序(推荐使用)</option>"
				Echo "	<option value=""1"">按照时间(按发表时间)</option>"
				Echo "	<option value=""2"">按照时间(按回复时间)</option>"
				Echo "	<option value=""3"">按照点击数(最热帖)</option>"
				Echo "	<option value=""4"">按照回复数(最热帖)</option>"
				Echo "	</select> </td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">调用版面：</td><td class=""altbg2""> <select name=""showboards"">"
				Echo "	<option value=""0"" SELECTED>显示所有版块</option>"
				Echo team.BBs_Value_List(0,0)
				Echo "	</select> </td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">标题长度：</td><td class=""altbg2""><input type=""text"" name=""psize"" size=""30"" value=""20""></td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">显示记录数：</td><td class=""altbg2""><input type=""text"" name=""total"" size=""30"" value=""10""></td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">时间格式：</td><td class=""altbg2""> <select name=""formattime"" ID=""formatTime"">"
				Echo "	<option value=""0"" SELECTED>YYYY-M-D H:M:S(长格式)</option>"
				Echo "	<option value=""1"">YYYY年M月D</option>"
				Echo "	<option value=""2"">YYYY-M-D</option>"
				Echo "	<option value=""3"">H:M:S</option>"
				Echo "	<option value=""4"">hh:mm</option>"
				Echo "	</select> [服务器设置: "& Now &"] </td></tr>"
				Echo "	</table><BR><center><input type=""submit"" name=""submit"" value=""生成JS""></center></form>"
			Case "3"
				Echo "<form action=""?action=Next2"" method=""post""><input type=""hidden"" name=""myjs"" value=""3""><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
				Echo "<tr class=""tab1""><td colspan=""2""> 版块调用</td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">板块样式：</td><td class=""altbg2""><input type=""radio"" name=""myclass"" value=""0"" class=""radio"" CHECKED> 树型结构 &nbsp; &nbsp; <input type=""radio"" name=""myclass"" value=""1"" class=""radio""> 地图结构</td></tr>"
				Echo "	</table><BR><center><input type=""submit"" name=""submit"" value=""生成JS""></center></form>"
			Case "4"
				Echo "<form action=""?action=Next2"" method=""post""><input type=""hidden"" name=""myjs"" value=""4""><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
				Echo "<tr class=""tab1""><td colspan=""2""> 会员调用</td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">显示记录数：</td><td class=""altbg2""><input type=""text"" name=""total"" size=""30"" value=""10""></td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">标题长度：</td><td class=""altbg2""><input type=""text"" name=""psize"" size=""30"" value=""20""></td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">会员排序：</td><td class=""altbg2""> <select name=""userorder"" ID=""userorder"">"
				Echo "	<option value=""0"" SELECTED>按注册时间</option>"
				Echo "	<option value=""1"">按用户发帖数</option>"
				Echo "	<option value=""2"">按用户总贴数</option>"
				Echo "	<option value=""3"">按用户精华贴数</option>"
				Dim ExtCredits,u
				ExtCredits = Split(team.Club_Class(21),"|")
				for u = 0 to ubound(ExtCredits)
					If Split(ExtCredits(u),",")(4)=1 Then
						Echo "<option value="""& u+4 &""">按用户" & Split(ExtCredits(u),",")(0) &"</option>"
					End if
				Next
				Echo "	</select>  </td></tr>"
				Echo "	</table><BR><center><input type=""submit"" name=""submit"" value=""生成JS""></center></form>"
			Case "5"
				Echo "<form action=""?action=Next2"" method=""post""><input type=""hidden"" name=""myjs"" value=""5""><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
				Echo "<tr class=""tab1""><td colspan=""2""> 公告调用</td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">显示记录数：</td><td class=""altbg2""><input type=""text"" name=""total"" size=""30"" value=""10""></td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">标题长度：</td><td class=""altbg2""><input type=""text"" name=""psize"" size=""30"" value=""20""></td></tr>"
				Echo "	<tr><td class=""altbg1"" align=""center"" width=""40%%"">时间格式：</td><td class=""altbg2""> <select name=""formattime"" ID=""formatTime"">"
				Echo "	<option value=""0"" SELECTED>YYYY-M-D H:M:S(长格式)</option>"
				Echo "	<option value=""1"">YYYY年M月D</option>"
				Echo "	<option value=""2"">YYYY-M-D</option>"
				Echo "	<option value=""3"">H:M:S</option>"
				Echo "	<option value=""4"">hh:mm</option>"
				Echo "	</select> [服务器设置: "& Now &"] </td></tr>"
				Echo "	</table><BR><center><input type=""submit"" name=""submit"" value=""生成JS""></center></form>"
		End Select
		team.footer
	End Sub

	Sub Next2
		Dim MyJs,MakeJs
		team.Headers(Team.Club_Class(1) & "- 论坛调用指南")
		MyJs = HRF(1,2,"myjs")
		Select Case MyJs
			Case "1"
				MakeJs = "FNews.asp?action=recallinfos&psize="&HRF(1,2,"psize")&"&formattime="&HRF(1,2,"formattime")&""
			Case "2"
				MakeJs = "FNews.asp?action=recallboardfids&formattime="&HRF(1,2,"formattime")&"&showboards="&HRF(1,2,"showboards")&"&showorders="&HRF(1,2,"showorders")&"&icon="&HRF(1,2,"icon")&"&showtypes="&HRF(1,2,"showtypes")&"&total="&HRF(1,2,"total")&"&psize="&HRF(1,2,"psize")&"&showname="&HRF(1,2,"showname")&""
			Case "3"
				MakeJs = "FNews.asp?action=recallshowlinks&myclass="&HRF(1,2,"myclass")&""
			Case "4"
				MakeJs = "FNews.asp?action=recallmembers&total="&HRF(1,2,"total")&"&psize="&HRF(1,2,"psize")&"&orderys="&HRF(1,2,"userorder")&""
			Case "5"
				MakeJs = "FNews.asp?action=recallaffiche&total="&HRF(1,2,"total")&"&psize="&HRF(1,2,"psize")&"&formattime="&HRF(1,2,"formattime")&""
		End Select
		Echo "<script language=""JavaScript"">"
		Echo "	<!-- "
		Echo "	function oCopy(obj){ "
		Echo "		obj.select(); "
		Echo "		var js=obj.createTextRange(); "
		Echo "		js.execCommand('Copy');"
		Echo "	}"
		Echo "	//-->"
		Echo "</script>"
		Echo "<table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
		Echo "<tr class=""tab1""><td>调用JS代码</td></tr>"
		Echo "	<tr class=""tab3""><td><textarea name=""makemyjs"" ID=""makemyjs"" style=""width:100%;height:50"" rows=""5"" cols=""50"">&lt;script src=&quot;" & weburl & MakeJs & "&quot;&gt;&lt;/script&gt;</textarea></td></tr></table><BR>"
		Echo "<table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""10"" align=""center"" class=""a2"">"
		Echo "<tr class=""tab1""><td>JS代码调用演示</td></tr>"
		Echo "<tr class=""a4""><td><script src=""" & weburl & MakeJs & """></script></td></tr>"
		Echo "	</table><BR><center><input type=""submit"" name=""submit"" value=""拷贝调用代码"" onclick=""oCopy($('makemyjs'));""></center>"
		team.footer
	End Sub

	Private Sub NotCallNews
		OutNews("系统禁止了调用功能。")
		Response.End
	End Sub

	Private Sub OutNews(s)
		Response.Write "document.write('"
		Response.Write s
		Response.Write "');"
		Response.Write vbNewline
	End Sub

	Private Function CheckServer()
		Dim i,servername,str
		Str = Trim(team.Club_Class(28))
		If str = "" Or IsNull(str) Then
			CheckServer = True
			Exit Function
		Else
			CheckServer = False
		End If
		servername=Request.ServerVariables("HTTP_REFERER")
		If instr(Cstr(str),Chr(13)&Chr(10)) > 0 Then
			str=split(Cstr(str),Chr(13)&Chr(10))
			For i=0 to Ubound(str)
				If Right(str(i),1)="/" Then str(i)=left(Trim(str(i)),Len(str(i))-1)
				If Lcase(left(servername,Len(str(i))))=Lcase(str(i)) then
					checkserver = True
					Exit For
				Else
					checkserver = False
				End if
			Next
		Else
			If Right(str,1)="/" Then str(i)=left(Trim(str),Len(str(i))-1)
			If Lcase(left(servername,Len(str)))=Lcase(str) then
				checkserver = True
				Exit Function
			Else
				checkserver = False
			End if		
		End if
	End Function

	Private Function Stringhtml(str)
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="<!--(.[^>]*)>"
		str=re.replace(str, "")
		Stringhtml=str
	End Function
	Function Fixjs(Strings)
		Dim Str
		Str = Strings
		str = Replace(str, CHR(39), "\'")
		str = Replace(str, CHR(13), "")
		str = Replace(str, CHR(10), "")
		str = Replace(str, "]]>","]]&gt;")
		Fixjs = str
	End Function
End Class
%>
