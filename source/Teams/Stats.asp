<!-- #include file="CONN.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim x1,x2,fID,Page
team.Headers(Team.Club_Class(1))
Select Case Request("action")
	Case "onilnes"
		Call onilnes
	Case "master"
		Call master
	Case "forums"
		Call forums
	Case "goodtopic"
		Call goodtopic
	Case Else
		Call Main()
End Select
Call team.footer

Sub Goodtopic

End Sub

Sub Main()
	Dim tmp,SQL,SqlQueryNum,RS,Maxpage,PageNum,iRs,Rmp,DispCount
	Dim i
	x1 = " <a href=""Stats.asp"">论坛近日新帖</a> "
	tmp = Replace(Team.ElseHtml (2),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"newpost"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"onlines"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"master"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"forums"))
	DispCount = team.execute("Select Count(ID) From ["&IsForum&"forum] Where deltopic=0 and PostTime> "&SqlNowString&"-3")(0)
	SQL="Select ID,Topic,UserName,Views,Replies,Lasttime From ["&IsForum&"forum] Where deltopic=0 and PostTime> "&SqlNowString&"-3 Order By Toptopic Desc,Lasttime DESC"
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
		tmp=Replace(tmp,"{$newposts}","")
	Else
		Rmp ="<tr class=""tab1""><td> 主题 </td><td> 作者 </td><td> 查看/回复 </td><td> 最后更新 </td></tr>"
		For i=0 To Ubound(iRs,2)
			Rmp = Rmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td><a href=""thread.asp?tid="&iRs(0,i)&""" target=""_blank"">"&iRs(1,i)&"</a> <img src="""&team.styleurl&"/new.gif"" border=""0"" align=""absmiddle""></td><td align=""center""> "&iRs(2,i)&" </td><td align=""center""> "&iRs(3,i)&" / "&iRs(4,i)&"</td> <td align=""center""> "&iRs(5,i)&" </td></tr> "
		Next
		tmp=Replace(tmp,"{$newposts}",Rmp)
	End If
	tmp=Replace(tmp,"{$pagecount}",PageNum)
	tmp=Replace(tmp,"{$dispcount}",DispCount)
	Echo tmp
End Sub

Sub onilnes
	Dim tmp,SQL,SqlQueryNum,RS,Maxpage,PageNum,iRs,Rmp,DispCount
	Dim i
	x1 = " <a href=""Stats.asp?action=onilnes"">在线列表</a> "
	Call team.OnlinActions("0,查看在线列表,查看在线列表")
	tmp = Replace(Team.ElseHtml (2),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"onlines"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"newpost"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"master"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"forums"))
	DispCount = team.execute("Select Count(*) From ["&IsForum&"Online]")(0)
	SQL="Select UserName,Ip,Eremite,Bbsname,Act,Acturl,Lasttime,Levelname From ["&IsForum&"Online] Order By Cometime DESC"
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
		tmp=Replace(tmp,"{$newposts}","")
	Else
		Rmp ="<tr class=""tab1""><td> 用户名 </td><td> 时间 </td><td> 用户IP </td><td> 所在论坛 </td><td> 当前动作 </td></tr>"
		For i=0 To Ubound(iRs,2)
			Rmp = Rmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td align=""center"">" 
			If iRs(0,i)="游客" Then
				Rmp = Rmp & iRs(0,i)
			Else
				If iRs(2,i) = 1 Then
					Rmp = Rmp & "隐身用户"
				Else 
					Rmp = Rmp & "<a href=""Profile.asp?username="&iRs(0,i)&""" target=""_blank"">"&iRs(0,i)&"</a>"
				End If 
			End If
			Rmp = Rmp & "</td><td align=""center""> "&iRs(6,i)&" </td><td align=""center""> "& iRs(1,i) &" </td> <td align=""center""> "&iRs(3,i)&" </td><td align=""center""> <a href="""&iRs(5,i)&""">"&iRs(4,i)&"</a></td></tr> "
		Next
		tmp=Replace(tmp,"{$newposts}",Rmp)
	End If
	tmp=Replace(tmp,"{$pagecount}",PageNum)
	tmp=Replace(tmp,"{$dispcount}",DispCount)
	Echo tmp
End Sub

Sub master
	If IsSqlDataBase = 1 Then team.Error "系统暂时停止统计数据的查看！"
	TestUser()
	Dim tmp,SQL,SqlQueryNum,RS,Maxpage,PageNum,iRs,Rmp,DispCount
	Dim i,Rs1
	x1 = " <a href=""Stats.asp?action=master"">管理团队</a> "
	tmp = Replace(Team.ElseHtml (2),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"master"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"newpost"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"onlines"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"forums"))
	DispCount = team.execute("Select Count(ID) From ["&IsForum&"User] Where UserGroupID =1 or UserGroupID = 2 or UserGroupID = 3 ")(0)
	SQL="Select UserName,Levelname,Posttopic,Postrevert,Regtime,Landtime,degree From ["&IsForum&"User] Where UserGroupID =1 or UserGroupID = 2 or UserGroupID = 3  Order By UserGroupID Asc"
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
		tmp=Replace(tmp,"{$newposts}","")
	Else
		Rmp ="<tr class=""tab1""><td> 论坛 </td><td> 用户名 </td><td> 用户等级 </td><td> 注册时间  </td><td> 上次登陆  </td><td> 在线时间 </td><td> 最近30天的发帖数 </td></tr>"
		For i=0 To Ubound(iRs,2)
			Dim Rsk,kNum,mNum
			Set Rsk = team.execute("Select Count(*) From ["&IsForum&"Forum] Where year(PostTime)=year(date()) and Month(PostTime)=month(Date()) and UserName = '"& iRs(0,i) &"'")
			If Not Rsk.Eof Then
				kNum = Rsk(0)
			End If
			Rsk.Close:Set Rsk=Nothing
			Set Rsk = team.execute("Select Count(*) From ["&IsForum & team.Club_Class(11) &"] Where year(PostTime)=year(date()) and Month(PostTime)=month(Date()) and UserName = '"& iRs(0,i) &"'")
			If Not Rsk.Eof Then
				mNum = Rsk(0)
			End If
			Rsk.Close:Set Rsk=Nothing
			Rmp = Rmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td>"
			Set Rs1 = team.execute("Select T.ID,T.BbsName From ["&IsForum&"bbsConfig] T Inner Join ["&IsForum&"Moderators] L On T.ID=L.BoardID Where L.ManageUser='"&iRs(0,i)&"'")
			Do While Not Rs1.Eof
				Rmp = Rmp & " <A href=""Forums.asp?fid="& Rs1(0) &""" target=""_blank"">"& Rs1(1) &"</a> "
				Rs1.MoveNext
			Loop
			Rs1.Close:Set Rs1 = Nothing
			Rmp = Rmp & " </td><td align=""center""> <a href=""Profile.asp?username="&iRs(0,i)&""" target=""_blank"">"&iRs(0,i)&"</a> </td><td align=""center""> "&Split(iRs(1,i),"||")(0)&" </td><td align=""center""> "&Formatdatetime(iRs(4,i),2)&" </td><td align=""center""> <span "
			If DateDiff("d",iRs(5,i),Now())>30 Then
				Rmp = Rmp & " style='color:red' title='超过30天未登陆'"
			End if
			Rmp = Rmp &" >"&iRs(5,i)&"</span></td><td align=""center""> 大约 "& CID(iRs(6,i)/60) &" 小时 </td><td align=""center"">" & CID(kNum) + CID(mNum) & "</td></tr> "
		Next
		tmp=Replace(tmp,"{$newposts}",Rmp)
	End If
	tmp=Replace(tmp,"{$pagecount}",PageNum)
	tmp=Replace(tmp,"{$dispcount}",DispCount)
	Echo tmp
End Sub

Sub forums
	TestUser()
	If IsSqlDataBase = 1 Then team.Error "系统暂时停止统计数据的查看！"
	Dim tmp,SQL,SqlQueryNum,RS,Maxpage,PageNum,iRs,Rmp,DispCount,iUs
	Dim i,vmp,uMaster,uStar,uStar1,mStar,mStar1,tStar,tStar1,pName
	x1 = " <a href=""Stats.asp?action=forums"">帖子统计</a> "
	tmp = Replace(Team.ElseHtml (2),"{$weburl}",team.MenuTitle)
	tmp = iHtmlEncode(BlackTmp(HtmlEncode(tmp),"forums"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"newpost"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"onlines"))
	tmp = iHtmlEncode(TempCode(HtmlEncode(tmp),"master"))
	uMaster = team.execute("Select Count(ID) From ["&IsForum&"User] Where UserGroupID >0 and UserGroupID<4")(0)

	Set Rs = team.execute("Select top 1 username,count(username) From ["&IsForum&"Forum] Where DateDiff('d',PostTime,"&SqlNowString&")>0 group by username order by count(username) desc")
	If Not Rs.Eof Then
		uStar = Rs(0)
	End If
	RS.Close:Set Rs=Nothing
	Set Rs = team.execute("Select count(username),username From ["&IsForum&"Forum] Where PostTime > Date() group by username order by count(username) desc")
	If Not  Rs.Eof Then
		uStar1 = Rs(0) : iUs = Rs(1)
	End If
	Set Rs = team.execute("Select count(username),username From ["&IsForum&"reForum] Where PostTime > Date() group by username order by count(username) desc")
	If Not  Rs.Eof Then
		If Trim(iUs) = Rs(1) Then
			uStar1 = Rs(0)+uStar1
		Else
			If uStar1 > Rs(0) Then
				uStar1 = uStar1
			Else
				uStar1 = Rs(0)
			End If 
		End if
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = team.execute("Select top 1 username,count(username) From ["&IsForum&"Forum] Where PostTime > Date()-weekday(date(),2) group by username order by count(username) desc")
	If Not  Rs.Eof Then
		mStar = Rs(0) 
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = team.execute("Select top 1 count(username),username From ["&IsForum&"Forum] Where PostTime > Date()-weekday(date(),2) group by username order by count(username) desc")
	If Not  Rs.Eof Then
		mStar1 = Rs(0): iUs = Rs(1)
	End If
	Set Rs = team.execute("Select top 1 count(username),username From ["&IsForum&"ReForum] Where PostTime > Date()-weekday(date(),2) group by username order by count(username) desc")
	If Not  Rs.Eof Then
		If Trim(iUs) = Rs(1) Then
			mStar1 = Rs(0)+mStar1
		Else
			If mStar1 > Rs(0) Then
				mStar1 = mStar1
			Else
				mStar1 = Rs(0)
			End If 
		End if
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = team.execute("Select top 1 username,count(username) From ["&IsForum&"Forum] Where year(PostTime)=year(date()) and Month(PostTime)=month(Date()) group by username order by count(username) desc")
	If Not  Rs.Eof Then
		tStar = Rs(0) 
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = team.execute("Select top 1 count(username),username From ["&IsForum&"Forum] Where year(PostTime)=year(date()) and Month(PostTime)=month(Date()) group by username order by count(username) desc")
	If Not  Rs.Eof Then
		tStar1 = Rs(0) : iUs = Rs(1)
	End If
	Set Rs = team.execute("Select top 1 count(username),username From ["&IsForum&"ReForum] Where year(PostTime)=year(date()) and Month(PostTime)=month(Date()) group by username order by count(username) desc")
	If Not  Rs.Eof Then
		If Trim(iUs) = Rs(1) Then
			tStar1 = Rs(0)+tStar1
		Else
			If mStar1 > Rs(0) Then
				tStar1 = tStar1
			Else
				tStar1 = Rs(0)
			End If 
		End if
	End If
	Rs.Close:Set Rs = Nothing
	Set Rs = team.execute("Select Count(ID) From ["&IsForum&"User] Where posttopic=0 and postrevert ")
	If Not  Rs.Eof Then
		pName = Rs(0)
	End If
	Rs.Close:Set Rs = Nothing
	vmp = "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td width=""30%""> 注册会员 </td><td width=""20%""> "& Application(CacheName&"_UserNum") &" </td><td width=""30%""> 发帖用户 </td><td width=""20%""> "& Application(CacheName&"_UserNum") - CID(pName) &" </td></tr>"
	vmp = vmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td> 管理团队 </td><td> "& uMaster &" </td><td> 未发帖用户 </td><td> "&CID(pName)&" </td></tr>"
	vmp = vmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td> 新入用户 </td><td> <a href=""Profile.asp?username="& Team.Club_Class(12) &""">"& Team.Club_Class(12) &"</a> </td><td> 发帖会员占总数 </td><td> "& FormatNumber((Application(CacheName&"_UserNum") - CID(pName))/Application(CacheName&"_UserNum"),4)*100 &" % </td></tr>"
	vmp = vmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td> 今日之星 </td><td> <a href=""Profile.asp?username="& uStar &""">"& uStar &"</a> ("& uStar1&") </td><td> 平均每人发帖数 </td><td> "& FormatNumber((Application(CacheName&"_UserNum") - CID(pName))/Application(CacheName&"_ConverPostNum"),4)*100 &" </td></tr>"
	vmp = vmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td> 本周之星 </td><td> <a href=""Profile.asp?username="& uStar &""">"& mStar &"</a> ("& mStar1&") </td><td> 本月之星 </td><td> <a href=""Profile.asp?username="& uStar &""">"& tStar &"</a> ("& tStar1&") </td></tr>"
	tmp=Replace(tmp,"{$newposts}",vmp)
	vmp = ""
	vmp = "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td width=""50%""> 论坛数 </td><td> "& team.execute("Select Count(*) From ["&IsForum&"BbsConfig] Where hide=0")(0)&" </td></tr>"
	vmp = vmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td> 主题数 </td><td> "& Application(CacheName&"_PostNum") &" </td></tr>"
	vmp = vmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td> 帖子数 </td><td> "& Application(CacheName&"_ConverPostNum") &" </td></tr>"
	vmp = vmp & "<tr class=""altbg2"" onMouseOver=""this.className='altbg1'"" onMouseOut=""this.className='altbg2'""><td> 最热门的论坛 </td><td> "& team.execute("Select top 1 BbsName,count(tolrestore) From ["&IsForum&"BbsConfig] Where Hide=0 group by BbsName order by count(tolrestore) desc")(0) &" </td></tr>"
	tmp=Replace(tmp,"{$newposts1}",vmp)
	tmp=Replace(tmp,"{$pagecount}",1)
	tmp=Replace(tmp,"{$dispcount}",1)
	Echo tmp
End Sub




%>
