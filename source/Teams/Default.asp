<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim RootID,Fid
RootID = Cid(Request.QueryString("rootid"))
team.Headers(Team.Club_Class(1))
Call Main_headers()
Call ForumList()
Call team.OnlinActions("0,首页,"&Team.Club_Class(1))
Call Main_footer()
team.Footer()

Sub Main_headers()
	Dim AffTemp,TempStr,AfficheTitle
	Dim i,PostCount
	TempStr = Team.IndexHtml(2)
	TempStr = Replace(TempStr,"{$clubname}",Team.Club_Class(1))
	AfficheTitle = Team.Affiche()
	If IsArray(AfficheTitle) Then
		For i=0 To Ubound(AfficheTitle,2)
			AffTemp=AffTemp & "<a href=""Affiche.asp#"&AfficheTitle(0,i)&""" target=""_blank""><Span Style='"&AfficheTitle(5,i)&"'> "&AfficheTitle(1,i)&" </Span></A>["&AfficheTitle(4,i)&"]&nbsp;&nbsp;&nbsp;"
		Next
		TempStr = Replace(TempStr,"{$affiche}",AffTemp)
	Else
		TempStr = Replace(TempStr,"{$affiche}",AfficheTitle)
	End If
	TempStr = Replace(TempStr,"{$Alltopic}",Application(CacheName&"_PostNum"))
	TempStr = Replace(TempStr,"{$Allpost}",Application(CacheName&"_ConverPostNum"))
	TempStr = Replace(TempStr,"{$Alluser}",Application(CacheName&"_UserNum") & " <span style=""font-weight: normal;"">/ <a href=""my.asp"">我的快捷菜单</a></span>")
	TempStr = Replace(TempStr,"{$todaypost}",Application(CacheName&"_TodayNum"))
	TempStr = Replace(TempStr,"{$oldpost}",Application(CacheName&"_OldTodayNum"))
	TempStr = Replace(TempStr,"{$newloginuser}",Team.Club_Class(12))
	If team.UserLoginED = False Then
		TempStr = Replace(TempStr,"{$levelname}","<b>游客</b>")
		TempStr = Replace(TempStr,"{$oldlogin}",Now())
		TempStr = Replace(TempStr,"{$mypost}","<b>0</b>")
		TempStr = Replace(TempStr,"{$mygood}","<b>0</b>")
		TempStr = Replace(TempStr,"{$myip}",Remoteaddr)
		TempStr = Replace(TempStr,"{$myext}","")
	Else
		Dim R,u,ExtCredits
		ExtCredits = Split(team.Club_Class(21),"|")
		for u = 0 to ubound(ExtCredits)
			If Split(ExtCredits(u),",")(4) =1 Then
				If R = "" Then
					R = Split(ExtCredits(u),",")(0) & " <b>"& Team.User_SysTem(14+u) &"</b>"
				Else
					R = R & " / " & Split(ExtCredits(u),",")(0) & " <b>"& Team.User_SysTem(14+u) &"</b>"
				End If
			End if
		Next
		TempStr = Replace(TempStr,"{$myip}",Remoteaddr)
		TempStr = Replace(TempStr,"{$levelname}","<b>"&team.LevelName(0)&"</b>")
		TempStr = Replace(TempStr,"{$oldlogin}",Team.User_SysTem(10))
		TempStr = Replace(TempStr,"{$mypost}","<b>"&Team.User_SysTem(5)+Cid(Team.User_SysTem(6))&"</b>")
		TempStr = Replace(TempStr,"{$myext}",R)
		TempStr = Replace(TempStr,"{$mygood}","<b>"&Team.User_SysTem(8)&"</b>")
		R = ""
	End If
	TempStr = Replace(TempStr,"{$hotpost}",HotPost())
	Echo TempStr
End Sub    

Sub ForumList()
	Dim ShowBbs,i,tmp
	Showbbs = team.BoardList()
	If Isarray(Showbbs) Then
		tmp = ""
		For i=0 to Ubound(Showbbs,2)
			If ShowBBs(3,i) = 0 Then
				If RootID>0 Then
					If RootID = Showbbs(0,i) Then
						tmp = Replace(Team.IndexHtml(3),"{$myid}",Showbbs(0,i))
						tmp = Replace(tmp,"{$bbsname}",Sign_Code(Showbbs(1,i),1))
						If Showbbs(2,i)=0 Then tmp = tmp & Team.IndexHtml (4)
						tmp = Replace(tmp,"{$num}",iif(Showbbs(2,i)=0,"7",team.Forum_setting(32)))
						Echo tmp
						Echo team.ForumList_tips(Showbbs(0,i))
						Echo Team.IndexHtml (7)
					End if
				Else
					tmp = Replace(Team.IndexHtml(3),"{$myid}",Showbbs(0,i))
					tmp = Replace(tmp,"{$bbsname}",Sign_Code(Showbbs(1,i),1))
					If Showbbs(2,i)=0 Then tmp = tmp & Team.IndexHtml (4)
					tmp = Replace(tmp,"{$num}",iif(Showbbs(2,i)=0,"7",team.Forum_setting(32)))
					Echo tmp
					Echo team.ForumList_tips(Showbbs(0,i))
					Echo Team.IndexHtml (7)
				End If
			End if
		Next
	End If
End Sub

Sub Main_footer()
	Dim tmp,i
	If RootID>0 Then 
		Exit Sub
	End if
	tmp = Team.IndexHtml(8)
	tmp = Replace(tmp,"{$userlogin}",IIf(team.UserLoginED=true,"display:None",""))
	tmp = Replace(tmp,"{$loginusername}",Tk_UserName)
	If team.UserLoginED=True Then
		tmp = Replace(tmp,"{$Levelname}",Team.Levelname(0))
	Else
		tmp = Replace(tmp,"{$Levelname}","游客")
	End if
	tmp = Replace(tmp,"{$SortShowForum}",iif(CID(team.Forum_setting(48))>0,iif(Cid(Session("loginnum"))> CID(team.Forum_setting(48)),"","none"),"none"))
	tmp = Replace(tmp,"{$daysshow}",IIf(team.Forum_setting(38)=1 or team.Forum_setting(38)=3,"","display:None"))
	tmp = Replace(tmp,"{$todaybrds}",IIf(team.Forum_setting(38)=1 or team.Forum_setting(38)=3,IIF(team.TodBds<>"","祝福 "& team.TodBds & " 生日快乐","暂无用户生日"),""))
	tmp = Replace(tmp,"{$linkshow}",IIf(team.Forum_setting(36)=1,"","display:None"))
	tmp = Replace(tmp,"{$tmlinks}",Iif(team.Forum_setting(36)=1 and Team.Forum_Link<>"",Team.Forum_Link,""))
	tmp = Replace(tmp,"{$onshow}",IIf(team.Forum_setting(40)=1,"","display:None"))
	tmp = Replace(tmp,"{$topslink}",Team.Linkshows )
	tmp = Replace(tmp,"{$grouponlines}",Team.OnlineShows)
	Dim Isshowset
	If team.Forum_setting(39)=1 or team.Forum_setting(39)=3 Then
		if Request("showlines")="yes" or Request("showlines")="" Then
			Isshowset = "<a href=""default.asp?showlines=no#online""><img src="""&team.Styleurl&"/collapsed_no.gif"" align=""right"" border=""0"" alt=""点击关闭在线状况"" /></a>"
		Else
			Isshowset = "<a href=""default.asp?showlines=yes#online""><img src="""&team.Styleurl&"/collapsed_yes.gif"" align=""right"" border=""0"" alt=""点击查看在线状况"" /></a>"
		End if
	Else
		if Request("showlines")="yes" Then
			Isshowset = "<a href=""default.asp?showlines=no#online""><img src="""&team.Styleurl&"/collapsed_no.gif"" align=""right"" border=""0"" alt=""点击关闭在线状况"" /></a>"
		Else
			Isshowset = "<a href=""default.asp?showlines=yes#online""><img src="""&team.Styleurl&"/collapsed_yes.gif"" align=""right"" border=""0"" alt=""点击查看在线状况"" /></a>"
		End if
	End if
	tmp = Replace(tmp,"{$showimg}",Isshowset)
	tmp = Replace(tmp,"{$showonlises}",Iif(Request("showlines")="yes" or team.Forum_setting(39)=1 or team.Forum_setting(39)=3,"","display:None"))
	tmp = Replace(tmp,"{$listonlieuser}",Iif(Request("showlines")="yes" or team.Forum_setting(39)=1 or team.Forum_setting(39)=3,team.Showlines(0),""))
	tmp = Replace(tmp,"{$allonline}",Team.Onlinemany)
	tmp = Replace(tmp,"{$regonline}",Team.Regonline)
	tmp = Replace(tmp,"{$lookonline}",Team.GuestOnline)
	tmp = Replace(tmp,"{$maxonline}",Split(Team.Club_Class(20),"|")(0))
	tmp = Replace(tmp,"{$maxtime}",Split(Team.Club_Class(20),"|")(1))
	Echo tmp
End Sub

Function HotPost()
	Dim RePo
	If CID(team.Forum_setting(113)) = 1 Then
		RePo = RePo& "<table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""98%"" align=""center"" class=""a2""><tr align=center><td Width='33%' class=""tab1"">最新图片</td><td Width='33%' class=""tab1"">热门话题</td><td Width='33%' class=""tab1"">最新主题</td></tr>"
		RePo = RePo& "<tr class=a4 Valign=Top><td>"& Advs() &"</td><td>"& LoadTop(1) &"</td><td>"& LoadTop(2) &"</td></tr></table><BR>"
	End If 
	HotPost = RePo
End Function

Function LoadTop(Str)
	Dim SQL,HotPage,Hottopic,m,HotRs
	SQL = "Select Top 12 ID,Topic,Replies,UserName,Views From ["&IsForum&"Forum] Where Deltopic=0 and CloseTopic=0 and Locktopic=0 and Auditing=0"
	Select Case Str
		Case 1
			SQL = SQL & " Order By Views Desc,Lasttime Desc"
		Case 2
			SQL = SQL & " Order By ID desc"
		Case Else
			SQL = SQL & " and Goodtopic=1 Order By Lasttime Desc,ID desc"
	End Select
	Cache.Name = "HotPage"& Str
	Cache.Reloadtime = 2  '默认更新时间为10分钟!可自己修改.时间越长,对资源消耗越小!
	If Not Cache.ObjIsEmpty() Then
		HotPage = Cache.Value
	Else
		Set HotRs=Team.Execute(SQL)
		If Not HotRs.Eof Then
			HotPage=HotRs.GetRows(12)
			Cache.Value=HotPage
		End If
		HotRs.Close:Set HotRs=Nothing
	End If
	If not IsArray(HotPage) Then
		Hottopic = "<center>暂无帖子记录</center>"
	Else
	For M=0 To Ubound(HotPage,2)
		Hottopic = Hottopic & IIF(CID(team.Forum_setting(65))=1,"<img src="""&team.styleurl&"/tip.gif"" align=""middle""> <a href=Thread-"&HotPage(0,m)&".html alt='主题:"&HotPage(1,m)&"&#xA;作者:"&HotPage(3,m)&"&#xA;浏览:"&HotPage(4,m)&"&#xA;回复:"&HotPage(2,m)&"'>"&Cutstr(HotPage(1,m),35)&"</a><br>","<img src="""&team.styleurl&"/tip.gif"" align=""middle""> <a href=Thread.asp?tid="&HotPage(0,m)&" alt='主题:"&HotPage(1,m)&"&#xA;作者:"&HotPage(3,m)&"&#xA;浏览:"&HotPage(4,m)&"&#xA;回复:"&HotPage(2,m)&"'>"&Cutstr(HotPage(1,m),35)&"</a><br>")
	Next
	End If
	LoadTop = Hottopic
End Function

function Advs()
	Dim Rs,i,Rs1,t,R1,R2,R3,U,GID
	i = 0: U=0
	Set Rs = team.execute("Select t.ID,t.topic,t.forumid,t.posttime,u.filename,u.ID,u.FileSize From ["&IsForum&"Forum] T Inner Join ["&IsForum&"Upfile] u on t.ID=U.ID Where U.types='gif' or U.types='jpg' and u.ID>0 and t.Deltopic=0 order by t.ID desc ")
	R1 = "" : R2="" : R3=""
	Do While Not Rs.Eof
		If gID <> Rs(0) Then
			i = i+1
			If I = 1 Then
				R1 = R1 & "images/Upfile/"& RS(4) : R2 = R2 & "Thread.asp?tid="& RS(0) : R3 = R3 & Fixjs(RS(1))
			Else
				R1 = R1 & "|" & "images/Upfile/"& RS(4) : R2 = R2 & "|" & "Thread.asp?tid="& RS(0) : R3 = R3 & "|" & Fixjs(RS(1))
			End If 
			If i >=5 Then Exit Do
		End If 
		gID = Rs(0)
		Rs.moveNext
	Loop
	Rs.Close:Set Rs = Nothing
    t = t & "  <script type=""text/javascript"">" & vbcrlf
    t = t & "  var swf_width=280;	 " & vbcrlf
    t = t & "  var swf_height=190;"  & vbcrlf
    t = t & "  var config='5|0xffffff|0x0099ff|50|0xffffff|0x0099ff|0x000000';"  & vbcrlf
    t = t & "  // config 设置分别为: 自动播放时间(秒)|文字颜色|文字背景色|文字背景透明度|按键数字色|当前按键色|普通按键色"  & vbcrlf
    t = t & "  var files='"& R1 &"';"  & vbcrlf
    t = t & "  var links='"& R2 &"';"  & vbcrlf
    t = t & "  var texts='"& R3 &"';"  & vbcrlf
	t = t & "   document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ swf_width +'"" height=""'+ swf_height +'"">'); "  & vbcrlf
	t = t & "  document.write('<param name=""movie"" value=""adv/focus.swf"" />'); "  & vbcrlf
	t = t & "  document.write('<param name=""quality"" value=""high"" />');"  & vbcrlf
	t = t & "  document.write('<param name=""menu"" value=""false"" />');"  & vbcrlf
	t = t & "  document.write('<param name=wmode value=""opaque"" />');"  & vbcrlf
	t = t & "  document.write('<param name=""FlashVars"" value=""config='+config+'&bcastr_flie='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'"" />');"  & vbcrlf
	t = t & "  document.write('<embed src=""adv/focus.swf"" wmode=""opaque"" FlashVars=""config='+config+'&bcastr_flie='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'"" menu=""false"" quality=""high"" width=""'+ swf_width +'"" height=""'+ swf_height +'"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />');"  & vbcrlf
	t = t & "  document.write('</object>'); </script>"  & vbcrlf
	Advs = t
End Function

%>