<%
Class Cls_Forum
	Public Forum_setting,Club_Class,UserLoginED,User_SysTem,Wid,TK_UserID,Linkshows
	Public UserGroupID,Newmessage,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Postblog,UserMebe,LoginNum,Levelname,UserName,UserPass,UserUp,Cookies_Path
	Public Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,UserColor,UserImg,Rank
	Public Group_Browse,Group_Manage,UserGroup,ActUrl,SkinKey,HtmlTemp,Onlinemany,Regonline,GuestOnline,HtmlNews
	Public Today,Bannertext,Styleurl,SkinID,Allword,IsWeTimes,XmlDoc
	Public IndexHtml,PostHtml,UserHtml,ElseHtml,Admin_Master,IsUbb,ServerUrl
	Public IsMaster,SuperMaster,BoardMaster,IsVips,UserGroupExs,SearcKeywordClass,iBuild
	Private SeeUIP,CloseForum,SearcKeyword

	Private Sub Class_Initialize()
		If Not Response.IsClientConnected Then Response.End
		UserLoginED = False : SkinID = 1 : SeeUIP = False
		IsMaster = false:SuperMaster= False:BoardMaster = False :IsVips = False 
		UserGroupID = 28 : IsUbb = 1 : iBuild = "20090420"
		TK_UserID = CID(Request.Cookies(Forum_sn)("UserID"))
		ActUrl = Replace((Request.ServerVariables("script_name") &"?" & Left(Request.ServerVariables("Query_String"),20)),"'","")
		IsWeTimes=FormatDateTime(Now(),0)'格式化时间
		Cache.Name = "NewCountDate" : ServerUrl = "http://server.team5.cn/"
		Cache.Reloadtime = 14400
		If Cache.ObjIsEmpty() Then
			Cache.Value = Now
		End If
		If DateDiff("d",CDate(Cache.Value),Now())<>0 Then
			UpNewsDate()
			Cache.Value = Now
		End If
	End Sub

	'论坛基本参数Allclass=0,Clubname=1,Cluburl=2,Homename=3,Homeurl=4,Badwords=5,Badip=6,Badlist=7,ManageText=8,CacheName=9,UpFileGenre=10,ReForumName=11,Newreguser=12,agreement=13,Nowdate=14,Today=15,oldday=16,PostNum=17,RepostNum=18,UserNum=19,ForumBest=20,ExtCredits=21,MustOpen=22,ClearMail=23,ClearIP=24,UserKey=25,BodyMeta=26,ClearPost=27,JsUrl=28,29=Starday
 	Public Sub GetForum_Setting()
		Dim Rs,SQL,Temp
     	Cache.Name = "Club_Class"
     	Cache.Reloadtime = 14400
	 	If Not Cache.ObjIsEmpty() Then
	    	Club_Class = Split(Cache.Value,"#@#")
	 	Else
			Set Rs = Execute("Select Allclass,Clubname,Cluburl,Homename,Homeurl,Badwords,Badip,Badlist,ManageText,CacheName,UpFileGenre,ReForumName,Newreguser,agreement,Nowdate,Today,oldday,PostNum,RepostNum,UserNum,ForumBest,ExtCredits,MustOpen,ClearMail,ClearIP,UserKey,BodyMeta,ClearPost,JsUrl,Starday from ["&Isforum&"Clubconfig]")
			Temp = Rs.GetString(,1, "#@#","","")
			Rs.Close:Set Rs=Nothing
			Cache.Value = Temp
			Club_Class = Split(Temp,"#@#")
			Application.Lock
			LockCache "TodayNum" , Club_Class(15)
			LockCache "OldTodayNum" , Club_Class(16)
			LockCache "PostNum" , Club_Class(17)
			LockCache "RepostNum" , Club_Class(18)
			LockCache "UserNum" , Club_Class(19)
	 	End If
		Forum_setting = Split(Club_Class(0),"$$$")
		Server.ScriptTimeout = Forum_setting(91)
		If Application(CacheName&"_TodayNum")="" or Application(CacheName&"_OldTodayNum")="" or Application(CacheName&"_PostNum")="" or Application(CacheName&"_RepostNum")="" or Application(CacheName&"_UserNum")="" Then Cache.DelCache("Club_Class")
		LockCache "ConverPostNum" , CID(Application(CacheName&"_PostNum")) + Application(CacheName&"_RepostNum")
		If DateDiff("d",CDate(Club_Class(14)),Now())<>0 Then
			UpNewsDate()
		End If
		If Trim(Club_Class(9))="ToWage" Then
			UpUserMonPosts
		End If
	End Sub

	Public Sub LockCache(SetName,NewValue)
		Application.Lock	'锁定
		Application(CacheName &"_"&SetName) = NewValue		'赋值
		Application.unLock	'解除锁定
	End Sub 

	Private Sub UpNewsDate
		'更新系统单日统计
		Dim t
		t = CID(Execute("Select SUM(today)From ["&Isforum&"Bbsconfig]")(0))
		Execute("Update ["&Isforum&"Clubconfig] Set Oldday="& t &",Nowdate="&SqlNowString&",Today=0")
		Execute("Update ["&Isforum&"bbsconfig] set today=0")
		If Day(Now) <> "1" Then 
			Execute("Update ["&Isforum&"Clubconfig] Set CacheName='NoWage'")
		Else
			Execute("Update ["&Isforum&"Clubconfig] Set CacheName='ToWage'")
		End if
		Cache.DelCache("Club_Class")
	End Sub

	Private Sub UpUserMonPosts
		'工资管理
		Dim Rs,URs
		If Day(Now) = "1" Then
			team.Execute("Update ["&Isforum&"Clubconfig] Set CacheName='NoWage'")
			Cache.DelCache("Club_Class")
			Set Rs = team.execute("Select WageMach,WageGroupID From ["&Isforum&"Wages]")
			Do While Not Rs.Eof 
				team.execute("Update ["&IsForum&"User] Set Extcredits"&Forum_setting(99)&"=Extcredits"&Forum_setting(99)&"+"&RS(0)&" Where UserGroupID = "& Int(Rs(1)) )
				If URs = "" Then
					URs = Rs(1)
				Else
					URs = URs & "," & Rs(1)
				End if
				Rs.MoveNext
			Loop
			Rs.Close:Set Rs=Nothing
			Set Rs = team.execute("Select UserName From ["&IsForum&"User] Where UserGroupID in ("&URs&") ")
			Do While Not Rs.Eof 
				team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('系统消息','"&Rs(0)&"','您本月的工资已经发放，请注意查收',"&SqlNowString&",'工资发放消息')")
				team.execute("Update ["&Isforum&"User] set Newmessage=Newmessage+1 where UserName='"&Rs(0)&"'")
				Rs.MoveNext
			Loop
			Rs.Close:Set Rs=Nothing
		End If
	End Sub

	'验证用户登陆
	Public Sub CheckUserLogin()
		Dim RS,Rmp
		If TK_UserID > 0 Then
			Set RS = Execute("Select UserName,UserPass,UserGroupID,Levelname,Newmessage,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Postblog,UserUp,LoginNum,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members,Friend,birthday From ["&Isforum&"User] where ID="& TK_UserID)
			If Rs.Eof And Rs.Bof Then
				CheckGuestLogin : Exit Sub
			Else
				If Not (Trim(RS(0))=Trim(TK_UserName) and RS(1)=TK_UserPass ) Then
					CheckGuestLogin : Exit Sub
				ElseIf Not trim(RS(13))=Trim(Request.Cookies(Forum_sn)("LoginNum")) Then 
					CheckGuestLogin : Exit Sub
				Else
					Rmp = Rs.GetString(,1, "#@#","","")
				End If
			End If
			RS.Close:Set Rs=Nothing
			User_SysTem = Split(Rmp,"#@#")
			'判断Session和Cookies用户名
			If User_SysTem(0)<>TK_UserName Then	
				CheckGuestLogin : Exit Sub
			End If
			UserGroupID = User_SysTem(2)
			UserUp = User_SysTem (12)
			UserGroupExs = User_SysTem (14)
			If InStr(User_SysTem(3),"||") > 0 Then
				Levelname = Split(User_SysTem(3),"||")
			Else
				Levelname = Split("附小一年级||||||0||0","||")
			End if
			Newmessage = User_SysTem(4)
			Members = User_SysTem(22)
			Select Case UserGroupID
				Case 1
					IsMaster = True
				Case 2
					SuperMaster = True
				Case 3
					BoardMaster = True
				Case 4 
					IsVips = True
				Case 5
					team.error " 您的帐号尚未激活。<meta http-equiv=refresh content=3;url=""GetUserInfo.asp"">"
				Case 7
					Response.Redirect "Close.asp"
			End Select
			UserLoginED = True
			If Forum_setting(38) >= 2 Then
				'IntoMsg 发起人,接受人,内容,标题
				If IsDate(User_SysTem(24)) And CID(Request.Cookies(Forum_sn)("mybirday"))=0 Then
					Dim mydate,udate
					udate = ""
					Mydate = Split(User_SysTem(24),"-")
					If UBound(Mydate) = 2 Then
						udate = Mydate(1) &"-"& Mydate(2)
						If DateDiff("d",CDate(udate),Date()) = 0 Then
							SetMyCookies "mybirday",1
							IntoMsg "系统消息",User_SysTem(0),Club_Class(1)&"全体用户和管理员祝您生日快乐!","生日祝福短信"
						End If
					End if
				End if
			End If
			GetGroupSetting()
		Else
			CheckGuestLogin
		End If
	End Sub

	Public Function ManageUser()
		ManageUser = False
		If IsMaster Then 
			ManageUser = True
			Exit Function
		End if
		If SuperMaster Then 
			ManageUser = True
			Exit Function
		End If
		If BoardMaster Then 
			ManageUser = True
			Exit Function
		End If
		If IsVips Then 
			If  Admin_Master =1 or Admin_Master =2 Or Admin_Master = 3 Then
				ManageUser = True
				Exit Function
			End If
		End If
	End Function
	
	Public Sub CheckGuestLogin	
		UserGroupID = 28
		UserLoginED = False
		TK_UserID = 0
		'Session(CacheName&"_UserLogin") = ""
		EmptyCookies
		TK_UserName = "游客"& Session.SessionID
		GetGroupSetting()
	End Sub

	Public Sub SetMyCookies(a,b)
		'判断Cookies更新目录
		Dim cookies_path_s,cookies_path_d,cookies_path,i
		cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
		cookies_path_d=ubound(cookies_path_s)
		cookies_path="/"
		For i=1 to cookies_path_d-1
			cookies_path=cookies_path&cookies_path_s(i)&"/"
		Next
		Response.Cookies(Forum_sn)(a) = b
		Response.Cookies(Forum_sn).path=cookies_path
	End Sub

	Public Sub EmptyCookies()
		'判断Cookies更新目录
		Dim cookies_path_s,cookies_path_d,cookies_path,i
		cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
		cookies_path_d=ubound(cookies_path_s)
		cookies_path="/"
		For i=1 to cookies_path_d-1
			cookies_path=cookies_path&cookies_path_s(i)&"/"
		Next
		Response.Cookies(Forum_sn)("username") = ""
		Response.Cookies(Forum_sn)("userpass") = ""
		Response.Cookies(Forum_sn)("LoginNum") = ""
		Response.Cookies(Forum_sn)("UserID") = 0
		Response.Cookies(Forum_sn).path=cookies_path
	End Sub

	Public Function Createpass()'系统分配随机密码
		Dim Ran,i,LengthNum
		LengthNum=16
		Createpass=""
		For i=1 To LengthNum
			Randomize
			Ran = CInt(Rnd * 2)
			Randomize
			If Ran = 0 Then
				Ran = CInt(Rnd * 25) + 97
				Createpass =Createpass& UCase(Chr(Ran))
			ElseIf Ran = 1 Then
				Ran = CInt(Rnd * 9)
				Createpass = Createpass & Ran
			ElseIf Ran = 2 Then
				Ran = CInt(Rnd * 25) + 97
				Createpass =Createpass& Chr(Ran)
			End If
		Next
		Createpass= Createpass
	End Function

	Private Sub GetGroupSetting()
		Dim tmp,Rs,SQL
		Cache.Reloadtime = Cid(Forum_setting(44))
		Cache.Name="GroupSetting_"& UserGroupID
		If Cache.ObjIsEmpty() Then 
			SQL = "Select IsBrowse,IsManage,GroupRank,UserImg,UserColor,GroupName,rank From ["&isforum&"UserGroup] where ID = " & UserGroupID
			Set Rs = Execute(SQL)
			If Rs.Eof Then
				Set Rs=Nothing
				SQL = "Select IsBrowse,IsManage,GroupRank,UserImg,UserColor,GroupName,rank From ["&isforum&"UserGroup] where ID = 28"
				Set Rs = Execute(SQL)
				Cache.value = Rs.GetString(,1, "$$##$$","","")
			Else
				Cache.value = Rs.GetString(,1, "$$##$$","","")
			End If
			Rs.close:Set Rs=nothing
		End If
		tmp = Split(Cache.Value,"$$##$$")
		Group_Browse = Split(tmp(0),"|") : Group_Manage = Split(tmp(1),"|") : Admin_Master = tmp(2) 
		'组名称||颜色||图片||星星||签名UBB
		If UserLoginED Then
			If Not (Trim(tmp(5))=Levelname(0)) Or Not (Trim(tmp(4))=Levelname(1)) Or Not (Trim(tmp(3))=Levelname(2)) Or Not (Trim(tmp(6))=Levelname(3)) Or Not (Int(Group_Browse(21)) = Int(Levelname(4))) Then
				Execute("Update ["&Isforum&"user] set Levelname='"&tmp(5)&"||"&tmp(4)&"||"&tmp(3)&"||"&tmp(6)&"||"&Group_Browse(21)&"',Landtime="&SqlNowString&" Where ID="& TK_UserID)
			End If
		End If
		Call UpUserClass()
	End Sub 

	Private Sub UpUserClass
		If UserLoginED Then
			If Group_Manage(5) = 1 Then
				SeeUIP = True
			End If
			If Not Isdate(User_SysTem(10)) Then User_SysTem(10) = Now()
			If DateDiff("d",User_SysTem(10),Date())<>0 Then
				Execute("Update ["&Isforum&"user] set UserUp='0|"&Now()&"',Landtime="&SqlNowString&" Where ID="& TK_UserID)
			End If
			'更新用户在线时间
			If Not IsDate(Request.Cookies("Class")("UserLogintime")) Then
				Response.Cookies("Class")("UserLogintime") = Now
			End if
			If DateDiff("s",CDate(Request.Cookies("Class")("UserLogintime")),IsWeTimes) > 600 Then
				Execute("update ["&Isforum&"user] set Degree=Degree+10,LastLoginIP='"&RemoteAddr&"' Where ID="& TK_UserID)
				Response.Cookies("Class")("UserLogintime") = Now
			End If
		End if
	End Sub

 	Public Sub LoadTemplates(ID)
		Dim Rs,SQL,value
		ID = INT(ID)
		Cache.Name = "Templates_"&ID
     	Cache.Reloadtime = Cid(Forum_setting(44))
		If Cache.ObjIsEmpty() Then
	    	Set Rs = Execute("Select StyleName,StyleWid,Styleurl,Style_index,Style_post,Style_user,Style_else,StyleCss From ["&Isforum&"Style] Where ID="& ID)
			If Rs.Eof and Rs.Bof Then
				Set Rs = Nothing
				Set Rs = Execute("Select StyleName,StyleWid,Styleurl,Style_index,Style_post,Style_user,Style_else,StyleCss From ["&Isforum&"Style] Where ID="& INT(team.Forum_setting(18)))
				If Rs.Eof And Rs.Bof Then
					Set Rs = Nothing
					Set Rs = Execute("Select StyleName,StyleWid,Styleurl,Style_index,Style_post,Style_user,Style_else,StyleCss From ["&Isforum&"Style] ")
					If Rs.Eof And Rs.Bof Then
						Response.Redirect "Club.asp?message=没有找到应有的模版，请导入新的模版文件。 "
					Else
						value = Rs.GetString(,1, "@|@","","")
					End if
				Else
					value = Rs.GetString(,1, "@|@","","")
				End If
			Else
				value = Rs.GetString(,1, "@|@","","")
			End If
			Cache.Value = value
			Rs.Close:Set Rs=Nothing
	 	End If
		HtmlTemp = Split(Cache.Value,"@|@")
		Styleurl=HtmlTemp(2)
		Wid=HtmlTemp(1)
		HtmlTemp(3)=Replace(Replace(HtmlTemp(3),"{$Csslist}",HtmlTemp(2)),"{$csswindth}",HtmlTemp(1))
		HtmlTemp(4)=Replace(Replace(HtmlTemp(4),"{$Csslist}",HtmlTemp(2)),"{$csswindth}",HtmlTemp(1))
		HtmlTemp(5)=Replace(Replace(HtmlTemp(5),"{$Csslist}",HtmlTemp(2)),"{$csswindth}",HtmlTemp(1))
		HtmlTemp(6)=Replace(Replace(HtmlTemp(6),"{$Csslist}",HtmlTemp(2)),"{$csswindth}",HtmlTemp(1))
		IndexHtml=Split(HtmlTemp(3),"@@@"):PostHtml=Split(HtmlTemp(4),"@@@")
		UserHtml=Split(HtmlTemp(5),"@@@"):ElseHtml=Split(HtmlTemp(6),"@@@")
		HtmlNews = Split(HtmlTemp(7),"@@@")
	End Sub

	Public Property Let ChooseName(ByVal strPkey)
		SkinKey = CID(strPkey)
	End Property

	Public Function AdvShows2()
		Dim Advtmp,i,topAdvs,e1,e2,ismast
		Advtmp = ForumAdvs()
		ismast = false
		If IsArray(Advtmp) Then
			For i = 0 To Ubound(Advtmp,2)
				If Advtmp(3,i) <>"" Then 
					If DateDiff("d",CDate(Advtmp(3,i)),Date())<0 Then Advtmp(5,i) = ""	
				End if
				If Advtmp(4,i) <>"" Then 
					If DateDiff("d",CDate(Advtmp(4,i)),Date())>0 Then Advtmp(5,i) = ""
				End if
				If CID(Advtmp(1,i)) = 5 Then
					e1= "theFloaters.addItem('floatAdv','document.body.clientWidth-120','document.body.clientHeight-80','"& Advtmp(5,i) &"');"
					ismast = true
				End If
				If CID(Advtmp(1,i)) = 6 Then
					e2 = "theFloaters.addItem('coupleBannerL',6,0,'<div style=""position: absolute; left: 6px; top: 6px;"">"& Advtmp(5,i) &"<br><img src=\""images/advclose.gif\"" onMouseOver=\""this.style.cursor=\'hand\'\"" onClick=\""closeBanner();\""></div>');theFloaters.addItem('coupleBannerR','document.body.clientWidth-6',0,'<div style=""position: absolute; right: 6px; top: 6px;"">"& Advtmp(5,i) &"<br><img src=\""images/advclose.gif\"" onMouseOver=\""this.style.cursor=\'hand\'\"" onClick=\""closeBanner();\""></div>');"
					ismast = true
				End If
				topAdvs = e1 & e2
			Next
		End If
		If ismast Then
			AdvShows2 = "<script src=""js/floatadv.js"" type=""text/javascript""></script><script type=""text/javascript""> "& topAdvs &" theFloaters.play();</script>"
		End If 
	End function


	Public Function AdvShows(a,b)
		Dim i,Advtmp,topAdvs,t,IsTrue,rmp
		Dim tmp,u,url,n,MyAdvBoards
		Advtmp = ForumAdvs()
		If IsArray(Advtmp) Then
			topAdvs = ""
			For i = 0 To Ubound(Advtmp,2)
				IsTrue = False
				If CID(Advtmp(1,i)) = Int(a) Then
					If Advtmp(2,i)="all" or Advtmp(2,i)="index" Then
						IsTrue = True
					Else
						If InStr(Advtmp(2,i),",")>0 Then
							MyAdvBoards =  Split(Advtmp(2,i),",")
							For t = 0 To UBound(MyAdvBoards)
								If CID(MyAdvBoards(t)) = Int(B) Then 
									IsTrue = True
								End If
							Next
						Else
							If CID(Advtmp(2,i)) = Int(B) Then
								IsTrue = True
							End If
						End If
					End If
					If IsTrue Then
						If Advtmp(3,i) <>"" Then 
							If DateDiff("d",CDate(Advtmp(3,i)),Date())<0 Then Advtmp(5,i) = ""	
						End if
						If Advtmp(4,i) <>"" Then 
							If DateDiff("d",CDate(Advtmp(4,i)),Date())>0 Then Advtmp(5,i) = ""
						End if
						If Advtmp(5,i)<>"" Then
							If topAdvs = "" Then
								topAdvs = Advtmp(5,i)
							Else
								topAdvs = topAdvs & "$$$" & Advtmp(5,i) 
							End if
						End If
					End If
				End If
			Next
			If Instr(topAdvs,"$$$")>0 Then
				u = Split(topAdvs,"$$$")
				AdvShows = u(Second(now) mod Ubound(u))
			Else
				AdvShows = topAdvs
			End If
			If a = 8 Then
				AdvShows = "<div style=""clear: right; float: right; display: inline; margin: 10px 10px 10px;"">"& AdvShows & "</div>"
			End if
		End if
	End Function

	Public Sub LoadTemps()
		Dim Openclock,Nexhour
		If CID(team.Forum_setting(106)) = 1 Then
			If Request.ServerVariables("HTTP_X_FORWARDED_FOR")>"" then CloseForum = 1
		End if		
		'定时关闭
		If Forum_setting(56)=1 Then
			Openclock=Split(Forum_setting(0),"*")
			Nexhour=Hour(Now())
			If Openclock(nexhour)=0 Then CloseForum = 1
		Else
			CloseForum = Forum_setting(2)
		End If
		If CloseForum=1 and Not IsMaster Then
			Response.Redirect "Club.asp?message=论坛维护中!"
		End If
		'IP锁定
		If Request.Cookies(Forum_sn & "Kill")("kill") = "1" Then
			If Not (SuperMaster Or IsMaster) Then Response.Redirect "Close.asp?action=ipclose"
		ElseIf Not Request.Cookies(Forum_sn & "Kill")("kill") = "0" Then
			Call LockIP()
			If Request.Cookies(Forum_sn & "Kill")("kill") = "1" Then
				If Not (SuperMaster Or IsMaster) Then Response.Redirect "Close.asp?action=ipclose"
			End If
		End If	
		'载入模板
		SkinID = INT(team.Forum_setting(18))
		If CID(SkinKey) > 0 Then
			SkinID = SkinKey
		End if
		If CID(Request.Cookies("Style")("skins")) > 0 Then
			SkinID = Request.Cookies("Style")("skins")
		End If
		LoadTemplates(SkinID)
	End Sub

	Public Sub Headers(s)
		LoadTemps()
		Dim TempStr
		'搜索引擎优化部分
		TempStr = Replace(IndexHtml(0),"{$Csslist}",Styleurl)
		TempStr = Replace(tempstr,"{$keywords}",Forum_setting(30))
		TempStr = Replace(TempStr,"{$description}",Forum_setting(31))
		TempStr = Replace(TempStr,"{$metainfo}",Club_Class(26))
		If IsWebSearch Then
			TempStr = Replace(TempStr,"{$TitleShow}",s & " | " & Forum_setting(66) &" | Power By Team Board")
			TempStr = TempStr & "<a href=""http://www.team5.cn"" title=""论坛,bbs,免费论坛,技术,学习,博客,asp,asp.net,电脑,软件,灌水,防火墙,开发,插件"">TEAM官方论坛</a>"
		Else
			TempStr = Replace(TempStr,"{$TitleShow}",s & " | " & Forum_setting(66) )
		End If
		'搜索引擎优化结束
		TempStr = TempStr & IndexHtml(1)
		TempStr = Replace(TempStr,"{$clubname}",Club_Class(1))
		If Not UserLoginED then
			TempStr = Replace(TempStr,"{$isguest}","")
			TempStr = Replace(TempStr,"{$isuser}","None")
		Else
			TempStr = Replace(TempStr,"{$isguest}","None")
			TempStr = Replace(TempStr,"{$isuser}","")
		End If
		TempStr = Replace(TempStr,"{$Eremite}",iif(Request.Cookies(Forum_sn)("Eremite")="1","<a href=""Login.asp?menu=eremite&upline=0"">上线</a>","<a href=""Login.asp?menu=eremite&upline=1"">隐身</a>"))
		TempStr = Replace(TempStr,"{$TK_UserName}","" & TK_UserName)
		TempStr = Replace(TempStr,"{$skinsmenu}",iif(team.Forum_setting(28)=1,UserStyle(),""))'风格菜单
		TempStr = Replace(TempStr,"{$teammenu}",iif(team.Forum_setting(92)=1,UserMenu(),""))'自添加菜单
		Dim Menu_1
		If IsMaster Or SuperMaster Then
			Menu_1 =  "<div class=menuitems><a href=admin.asp target=_top>后台管理</a></div>"
		End If
		If BoardMaster Or IsMaster Or SuperMaster Then
			Menu_1 = Menu_1 & "<div class=menuitems><a href=BoradServer.asp target=_top>前台管理</a></div>"
		End if
		TempStr = Replace(TempStr,"{$mastermenu}",iif(IsMaster Or SuperMaster Or BoardMaster,"<li><a onmouseover="&Chr(34)&"showmenu(event,'"& Menu_1 &"')"&Chr(34)&">管理</a></li>",""))'菜单
		TempStr = Replace(TempStr,"{$forumlock}",IIf(CID(team.Forum_setting(116))=1,"<li><a href=""Cclist.asp"">CC视频展区</a></li>","") & iif(CloseForum=1,"SERVER: <font color=RED>OFF</font>",""))
		TempStr= Replace(TempStr,"{$message}",TeamNewMsg)
		Dim Advtmp,i,topAdvs,ShowAdv,p,IsTrue
		Advtmp = ForumAdvs()
		TempStr = Replace(TempStr,"{$banner}",AdvShows(1,fid))
		If IsArray(Advtmp) Then
			topAdvs = "":p = 0
			For i = 0 To Ubound(Advtmp,2)
				IsTrue = False
				If CID(Advtmp(1,i)) = 3 Then
					If Advtmp(2,i)="all" or Advtmp(2,i)="index" Then
						IsTrue = True
					Else
						If InStr(Advtmp(2,i),",")>0 Then
							MyAdvBoards =  Split(Advtmp(2,i),",")
							For t = 0 To UBound(MyAdvBoards)
								If CID(MyAdvBoards(t)) = Int(fid) Then 
									IsTrue = True
								End If
							Next
						Else
							If CID(Advtmp(2,i)) = Int(fid) Then
								IsTrue = True
							End If
						End If
					End If
					If IsTrue Then
						If Advtmp(3,i) <>"" Then 
							If DateDiff("d",CDate(Advtmp(3,i)),Date())<0 Then Advtmp(5,i) = ""	
						End if
						If Advtmp(4,i) <>"" Then 
							If DateDiff("d",CDate(Advtmp(4,i)),Date())>0 Then Advtmp(5,i) = ""
						End if
						If Advtmp(5,i)<>"" Then
							p = p+1
							topAdvs = topAdvs & "<td>"& Advtmp(5,i) &"</td>"
							if p = 5 Then
								topAdvs = topAdvs & "</tr><tr class=""tab4"">":p=0
							End if
						End if
					End If
				End If
			Next
			TopAdvs = IIF(TopAdvs&"" = "" ,"","<div class=""advs""><br><table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" class=""a2""><tr class=""tab4"">"& topAdvs & "</table><br></div>")
		End if
		TempStr = Replace(TempStr,"{$bannerall}",topAdvs)
		TempStr = Replace(TempStr,"{$mymessage}",Newmessage)
		Echo TempStr
	End Sub

	'论坛尾部
	Public sub Footer		
		Dim Temp,MSCode
		If IsSqlDataBase = 1 Then
			MSCode="SQL"
		Else
			MSCode="ACC"
		End If
		Temp = Replace(IndexHtml(9),"{$Foruminfo}","<a target=""_blank"" href=""http://www.team5.cn""><b>"& team.Forum_setting(8) &"</b></a> - ")
		Temp = Replace(Temp,"{$edition}","<a href=""Licence.asp""><b style='color:#FF9900'>"&MSCode&"</b></a>")
		Temp = Replace(Temp,"{$runtime}",iif(team.Forum_setting(37)=1,"Processed in " & FormatNumber((Timer()-StarTime)*1000,2,-1) & " ms," &SqlQueryNum& " queries",""))
		Temp = Replace(Temp,"{$TimeZone}",Forum_setting(12))
		Temp = Replace(Temp,"{$clubname}",Club_Class(3))
		Temp = Replace(Temp,"{$cluburl}",Club_Class(4))
		Temp = Replace(Temp,"{$regkey}","<A href=""http://www.miibeian.gov.cn/"">"& Forum_setting(59) &"</a>")
		Temp = Replace(Temp,"{$adcode}",AdvShows(2,fid))
		Temp = Temp & AdvShows2() & LoadFooterAdvs
		Echo Temp
		Htmlend
	End Sub

	'风格选择菜单
	Private Function UserStyle()
		Dim value,tmp,i,temp,rs
		Cache.Reloadtime = Cid(Forum_setting(44))
		Cache.Name = "TemplatesLoad"
		If Cache.ObjIsEmpty() Then
	   		Set Rs=Execute("Select StyleName,ID From ["&Isforum&"Style] order by ID Desc")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			RS.Close:Set RS=Nothing
		End If
		value = Cache.Value
		For i=0 To Ubound(value,2)
			Tmp = Tmp & "<div class=menuitems><a href=Cookies.asp?action=skins&styleid="&value(1,i)&">"&value(0,i)&"</a></div>"
		Next
		temp = "<li><a onmouseover="&Chr(34)&"showmenu(event,'"&tmp&"')"&Chr(34)&">风格</a></li>"
		UserStyle = temp
	End Function

	Public Function ForumAdvs()
		Dim SQL,RS,tmp
		Cache.Reloadtime = Cid(Forum_setting(44))
		Cache.Name = "ForumAdvsLoad"
		If Cache.ObjIsEmpty() Then
	   		Set Rs=Execute("Select Dois,Types,Boards,StarTime,StopTime,bodys From ["&Isforum&"AdvList] Where Dois=1 order by Sorts Desc")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			RS.Close:Set RS=Nothing
		End If
		ForumAdvs = Cache.Value
	End Function

	Private Function LoadUserMenu()
		Dim Rs
		Cache.Reloadtime = Cid(Forum_setting(44))
		Cache.Name = "MenuLoad"
		If team.Forum_setting(92)= 0 Then Exit Function
		If Cache.ObjIsEmpty() Then
			Set Rs=Execute("Select ID,Name,Url,Followid From ["&Isforum&"Menu] Where newtype = 1 order by SortNum Desc")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			RS.Close:Set RS=Nothing
		End If
		LoadUserMenu = Cache.Value
	End Function

	Private Function UserMenu()
		Dim Selectby,Mymenu,MenuTemp,i
		MenuTemp = LoadUserMenu
		If IsArray(MenuTemp) Then
			For i=0 To Ubound(MenuTemp,2)	
				If MenuTemp(3,i) = 0 Then
					Mymenu = Mymenu & "<li><a onmouseover="&Chr(34)&"showmenu(event,'"& MiniMenu(MenuTemp(0,i)) &"')"&Chr(34)&" style=""cursor:default""> "&MenuTemp(1,i)&"</a></li>"
				End if
			Next
		End If
		UserMenu = Mymenu
	End Function

	Public Function MiniMenu(a)
		Dim Selectby,MenuTemp,i
		Menutemp = LoadUserMenu
		If IsArray(MenuTemp) Then
			For i=0 To Ubound(MenuTemp,2)	
				If Menutemp(3,i) = a Then
					Selectby=Selectby & "<div class=menuitems><a href="&MenuTemp(2,i)&">"&MenuTemp(1,i)&"</a></div>"
				End if
			Next
		End If
		MiniMenu = Selectby
	End Function

	REM 版主
	Public Function GroupManages()	
		Dim Rs,Moderuser,tmp
		Cache.Reloadtime = Cid(Forum_setting(44))
		Cache.Name = "ManageUsers"
		If Cache.ObjIsEmpty() Then
	   		Set Rs=team.execute("Select ID,ManageUser,BoardID from ["&isforum&"Moderators] ")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			Rs.Close:Set Rs=Nothing
		End If
		GroupManages = Cache.Value
	End Function

	'IntoMsg 发起人,接受人,内容,标题
	Sub IntoMsg(a,b,c,d)
		team.Execute("insert into ["&Isforum &"Message] (author,incept,content,Sendtime,MsgTopic) values ('"& HtmlEncode(a) &"','"& HtmlEncode(b) &"','"& HtmlEncode(c) &"',"&SqlNowString&",'"& HtmlEncode(d) &"')")
		team.execute("Update ["&Isforum&"User] set Newmessage=Newmessage+1 where UserName='"& HtmlEncode(b) &"'")
	End Sub

	REM 生日
	Public Function TodBds()	
		Dim Rs,mans,tmp,i
		If Application(CacheName&"_Nobady") = 1 Then
			Exit Function
		End if
		Cache.Reloadtime = Cid(Forum_setting(44))
		Cache.Name = "UserBirthdays"
		If Cache.ObjIsEmpty()  Then
	   		Set Rs=team.Execute("Select UserName,Birthday From ["&IsForum&"User]")
	   		If RS.Eof Then
				LockCache "Nobady",1
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			Rs.Close:Set Rs=Nothing
		End If
		tmp = Cache.Value
		Mans = ""
		If Isarray(tmp) Then
			for i = 0 to ubound(tmp,2)
				If IsDate(tmp(1,i)) Then
					Dim mydate,udate
					udate = ""
					Mydate = Split(tmp(1,i),"-")
					If UBound(Mydate) = 2 Then
						udate = Mydate(1) &"-"& Mydate(2)
						If DateDiff("d",CDate(udate),Date()) = 0 Then
							If Mans = "" Then
								Mans = IIF(CID(team.Forum_setting(65))=1,"<a href=""Profile.asp-"& tmp(0,i) &".html"" target=""_blank""> "& tmp(0,i) &"</a> ","<a href=""Profile.asp?username="& tmp(0,i) &""" target=""_blank""> "& tmp(0,i) &"</a> ")
							Else
								Mans = Mans & IIF(CID(team.Forum_setting(65))=1," , <a href=""Profile-"& tmp(0,i) &".html"" target=""_blank""> "& tmp(0,i) &"</a> "," , <a href=""Profile.asp?username="& tmp(0,i) &""" target=""_blank""> "& tmp(0,i) &"</a> ")
							End If
						End If
					End if
				End if
			Next
		End if
		TodBds = Mans 
	End Function

	Public Function BoardList()	
		Dim Rs,Moderuser
		Cache.Reloadtime = Cid(Forum_setting(44))
		Cache.Name = "BoardLists"
		If Cache.ObjIsEmpty() Then
	   		Set Rs=team.Execute("Select ID,bbsname,Board_Model,Followid,Readme,today,toltopic,tolrestore,icon,Board_Last,Pass,Board_URL From ["&IsForum&"Bbsconfig] Where Hide=0 Order By SortNum")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			Rs.Close:Set Rs=Nothing
		End If
		BoardList = Cache.Value
	End Function

	Public Function ForumList_tips(a)
		Dim Rs,Moderuser,tmp,Bbstmp,i,LastPost
		Dim CheckMaster,Moder1,Moder2,u
		Dim Tmpid,j,Chcheid,R
		Bbstmp = BoardList
		If Isarray(Bbstmp) Then
			tmp = "" : R=0
			For i=0 to Ubound(Bbstmp,2)
				If Bbstmp(3,i) = a Then
					R = R+1
					If Bbstmp(2,i) = 0 Then
						tmp = tmp & IndexHtml (5)
					Else
						tmp = tmp & IndexHtml (6)
						tmp = Replace(tmp,"{$tr}",iif(R = 1,"<tr class=""tab4"">",""))
						If R >= Int(Forum_setting(32)) Then 
							tmp = Replace(tmp,"{$trend}","</tr>")
							R = 0
						Else
							tmp = Replace(tmp,"{$trend}","")
						End if
					End If
					'ID=0,bbsname=1,Board_Model=2,Followid=3,Readme=4,today=5,
					'toltopic=6,tolrestore=7,icon=8,Board_Last=9,Pass=10,Board_URL=11
					tmp = Replace(tmp,"{$boardurl}",IIF(Bbstmp(11,i)&""="",IIF(CID(team.Forum_setting(65))=1,"Forums-"&Bbstmp(0,i)&".html","Forums.asp?fid="&Bbstmp(0,i)&""),Bbstmp(11,i)))
					tmp = Replace(tmp,"{$today}",Iif(Bbstmp(5,i)>0,"Board0.gif","Board1.gif"))
					tmp = Replace(tmp,"{$bbsname}",Sign_Code(Bbstmp(1,i),1))
					tmp = Replace(tmp,"{$id}",Bbstmp(0,i))
					tmp = Replace(tmp,"{$passlogin}",iif(Trim(Bbstmp(10,i))&""="",""," [<FONT COLOR=""#FF0000"">密码验证</FONT>] "))
					tmp = Replace(tmp,"{$intro}","{$intro}{$chcheid}")
					tmp = Replace(tmp,"{$intro}",Sign_Code(Bbstmp(4,i),1)&"")
					tmp = Replace(tmp,"{$icon1}",Iif(Bbstmp(8,i)&""="","","<img src="""&Bbstmp(8,i)&""" border=""0"" align=""absmiddle"">"))
					tmp = Replace(tmp,"{$toltopic}",Bbstmp(6,i))
					tmp = Replace(tmp,"{$tolrestore}",Bbstmp(7,i))
					tmp = Replace(tmp,"{$todays}",Bbstmp(5,i))
					LastPost = Split(Bbstmp(9,i),"$@$")
					tmp = Replace(tmp,"{$toppass}","主题："& LastPost(0) &"<BR>作者：" &  IIF(CID(team.Forum_setting(65))=1,"<a href=""Profile-"& LastPost(1) &".html"" target=""_blank""> ","<a href=""Profile.asp?username="& LastPost(1) &""" target=""_blank""> ") & LastPost(1) &"</a><BR>时间："& LastPost(2) )
					CheckMaster = GroupManages()
					Moder1 = "" : Moder2 = ""
					tmp = Replace(tmp,"{$style_1}",IIf(Forum_setting(26)=0,"block","none"))
					tmp = Replace(tmp,"{$style_2}",IIf(Forum_setting(26)=0,"none","block"))
					If Isarray(CheckMaster) Then
						For u=0 to Ubound(CheckMaster,2)	
							If CheckMaster(2,u) = Cid(Bbstmp(0,i)) Then
								Moder1 = Moder1 &"<option> "&CheckMaster(1,u)&" </option>"
								If Moder2 = "" Then
									Moder2 = Moder2 & IIF(CID(team.Forum_setting(65))=1," <a href=Profile-"&CheckMaster(1,u)&".html>"&CheckMaster(1,u)&"</a> "," <a href=Profile.asp?username="&CheckMaster(1,u)&">"&CheckMaster(1,u)&"</a> ")
								Else
									Moder2 = Moder2 & IIF(CID(team.Forum_setting(65))=1,", <a href=Profile-"&CheckMaster(1,u)&".html>"&CheckMaster(1,u)&"</a> ",", <a href=Profile.asp?username="&CheckMaster(1,u)&">"&CheckMaster(1,u)&"</a> ")
								End If
							End If
						Next
						tmp = Replace(tmp,"{$moder1}",Moder1)
						tmp = Replace(tmp,"{$moder2}",iif(Moder2&""="","空缺",Moder2))
					Else
						tmp = Replace(tmp,"{$moder2}","")
						tmp = Replace(tmp,"{$moder1}","")
					End if
					Tmpid = ""
					Chcheid = BoardList
					If isarray(Chcheid) and CID(Forum_setting(27))=1 Then
						For j=0 to Ubound(Chcheid,2)
							If Cid(Bbstmp(0,i)) = Cid(Chcheid(3,j)) Then
								Dim showMinis
								If Chcheid(11,j)&""="" Then
									If CID(team.Forum_setting(65))=1 Then
										showMinis = "<A href=""Forums-"&Chcheid(0,j)&".html"" style=""text-decoration:underline"">"&Chcheid(1,j)&"</a>&nbsp;&nbsp;"
									Else
										showMinis = "<A href=""Forums.asp?fid="&Chcheid(0,j)&""" style=""text-decoration:underline"">"&Chcheid(1,j)&"</a>&nbsp;&nbsp;"
									End If
								Else
									If InStr(Chcheid(11,j),"http://")=0 Or InStr(Chcheid(11,j),"://")=0 Then
										Chcheid(11,j) = "http://" & Chcheid(11,j)
									End if
									showMinis = "<A href="""& Chcheid(11,j) &""">"&Chcheid(1,j)&"</a>&nbsp;&nbsp;"
								End if
								Tmpid = Tmpid & showMinis
							End if
						Next
					End If
					tmp = Replace(tmp,"{$chcheid}",iif(Tmpid<>"","<br><B>子论坛</B>:"&Tmpid&" ",""))
				End if
			Next
		End If
		ForumList_tips = tmp
	End Function
	'用户在线部分
	Public Sub OnlinActions(s)
		Dim UserSessionID,SQl,Rs,Eremite,Onlineuser,UserActions,SQL1,Fid,Act,Bbsname,U,iActUrl
		U = 0
		UserSessionID = Ccur(Session.SessionID) : UserActions = Split(s,",")
		Eremite = CID(Request.Cookies(Forum_sn)("Eremite")) : iActUrl = HtmlEncode(Replace(ActUrl,"&",""))
		If Not IsDate(Request.Cookies("Class")("UpUserInfos")) Then
			Response.Cookies("Class")("UpUserInfos") = Now
		End If	
		If Not UserLoginED Then
			'游客部分
			If IsWebSearch Then 
				Exit Sub
			Else
				Set Rs = Execute("Select Acturl,Forumid From ["&Isforum&"Online] Where Sessionid = " &UserSessionID )
				If Rs.Eof And Rs.Bof Then
					Execute("Insert Into ["&Isforum&"Online](Forumid,Sessionid,UserName,Ip,Eremite,Bbsname,Act,Acturl,Cometime,Lasttime,Levelname) Values (" & CID(UserActions(0)) & "," &UserSessionID& ",'游客','"& RemoteAddr &"',-1,'"& HtmlEncode(UserActions(2)) &"','"& HtmlEncode(UserActions(1)) &"','" & team.CheckStr(iActUrl) &"',"&SqlNowString & "," & SqlNowString & ",'游客')" )
					'更新在线人数
					UpdateOnline(CID(UserActions(0)))
					'将在线列表数据进行更新
					Cache.DelCache("ShowLines"&UserActions(0))
				Else
					If DateDiff("s",CDate(CDate(Request.Cookies("Class")("UpUserInfos"))),IsWeTimes) > 60 Or Not (Trim(RS(0)) = Trim(iActUrl)) Then
						Execute("Update ["&Isforum&"Online] Set Lasttime = " & SqlNowString & ",Forumid=" & CID(UserActions(0)) & ",Ip='" & RemoteAddr & "',BbsName='"& HtmlEncode(UserActions(2)) &"',Act='"& HtmlEncode(UserActions(1)) &"',Acturl='"& team.CheckStr(iActUrl) &"' Where Sessionid = " & UserSessionID )
						UpdateOnline(CID(UserActions(0)))
						Response.Cookies("Class")("UpUserInfos") = Now
					End If
					'判断用户活动到另外板块才更新在新列表记录
					If Not CID(Rs(1)) = CID(UserActions(0)) Then
						Cache.DelCache("ShowLines"&UserActions(0))
					End If
				End If
				Rs.Close:Set Rs = Nothing
			End if
		Else
			'注册用户部分
			SQL1 = "Select Acturl,Eremite From ["&Isforum&"Online] Where Sessionid ="& TK_UserID
			Set Rs = Execute(SQL1)
			If Rs.Eof and Rs.Bof Then
				Execute("Insert Into ["&Isforum&"Online](Forumid,Sessionid,Username,Ip,Eremite,Bbsname,Act,Acturl,Cometime,Lasttime,Levelname) Values (" & CID(UserActions(0)) & "," & TK_UserID & ",'"& TK_UserName&"','"& RemoteAddr &"',"& CID(Eremite) &",'"& HtmlEncode(UserActions(2)) &"','"& HtmlEncode(UserActions(1)) &"','" & team.CheckStr(iActUrl) &"',"&SqlNowString & "," & SqlNowString & ",'"&Members&"')" )
				Execute("Delete From ["&Isforum&"Online] Where Sessionid = " & UserSessionID)
				'更新在线人数
				UpdateOnline(CID(UserActions(0)))
				'将在线列表数据进行更新
				Cache.DelCache("ShowLines"&UserActions(0))
				Cache.DelCache("UserOnlineCache")
			Else
				If DateDiff("s",CDate(Request.Cookies("Class")("UpUserInfos")),IsWeTimes) > 60 Or Not (Trim(RS(0)) = Trim(iActUrl)) or Not (Eremite = Cid(RS(1)) ) Then
					Execute("Update ["&Isforum&"Online] Set Lasttime = " & SqlNowString & ",Forumid = " &CID(UserActions(0))& ",Ip = '" & RemoteAddr & "',BbsName='"& HtmlEncode(UserActions(2)) &"',Act='"& HtmlEncode(UserActions(1)) &"',Acturl='"& team.CheckStr(iActUrl) &"',UserName='"& TK_UserName &"',Eremite="&Eremite&",Levelname='"&Members&"' Where Sessionid = " & TK_UserID )
					Response.Cookies("Class")("UpUserInfos") = Now
					UpdateOnline(CID(UserActions(0)))
				End If			
				'判断用户活动到另外板块才更新在新列表记录
				If Not CID(Rs(1)) = CID(UserActions(0)) Then
					Cache.DelCache("ShowLines"&UserActions(0))
				End If
			End If
			Rs.Close:Set Rs = Nothing
		End If
		'删减人数并进行重新统计
		DelOnline(CID(UserActions(0)))	
		UserOnlineinfos()
	End Sub

	Public Sub DelOnline(a)
		Cache.Reloadtime= 60
		'判断在线总人数进行更新。
		Cache.Name="ForumOnline"
		If Cache.ObjIsEmpty() Then UpdateOnline(a)
		Onlinemany = Cache.Value
		If Int(Onlinemany) > Cache.Value Then
			Cache.Value = Onlinemany
		End if
		'判断在线注册用户人数进行更新。
		Cache.Name="ForumUserOnline"
		If Cache.ObjIsEmpty() Then UpdateOnline(a)
		Regonline = Cache.Value
		'修正统计值
		If CID(Regonline) > CID(Onlinemany) Then UpdateOnline(a)
		'设置游客数。
		GuestOnline = CID(Onlinemany) - CID(Regonline)
		'========================================================
		'设置删除不活动用户的时间
		Cache.Name = "GetNewsOnlinetime"
		If Cache.ObjIsEmpty() Then Cache.Value = Now()
		If DateDiff("s",Cache.Value,Now())> CID(Forum_setting(45))*10 then
			Rem 设置每N×10秒进行判断，删除超时用户
			If IsSqlDataBase=1 Then
				Execute("Delete From ["&Isforum&"Online] Where Datediff(Mi, Lasttime, " & SqlNowString & ") > " & Clng(Forum_setting(45)))
			Else
				Execute("Delete From ["&Isforum&"Online] Where Datediff('s',Lasttime, " & SqlNowString & " ) > "& Forum_setting(45) &" * 60 ")
			End If
			Cache.Value=Now()
			UpdateOnline(a)
			Cache.DelCache("UserOnlineCache")
		End If
		Rem 更新在线峰值
		If Int(Split(Club_Class(20),"|")(0))<Int(Onlinemany) Then
			Execute("update ["&Isforum&"ClubConfig] set ForumBest='"&CID(Onlinemany)&"|"& Now() &"' ")
			Club_Class(20) = Onlinemany &"|" & Now
			Cache.DelCache("Club_Class")
		End If
	End Sub

	Public Sub UpdateOnline(a)
		Dim Rs	
		Cache.Reloadtime = 60
		'总人数
		Cache.Name="ForumOnline"
		Set Rs=Execute("Select Count(*) From ["&Isforum&"Online]")
		Cache.Value = CID(Rs(0))
		Onlinemany = Cache.Value
		'总注册人数
		Cache.Name="ForumUserOnline"
		Set Rs=Execute("Select Count(*) From ["&Isforum&"Online] Where Eremite>-1")
		Cache.Value = CID(Rs(0))
		Regonline = Cache.Value
		If Int(a) > 0 Then
			Set Rs=Execute("Select Count(*) From ["&Isforum&"Online] Where forumid="&CID(a))
			Cache.Name = "Forumidonline"& a
			Cache.Value = CID(Rs(0))
			Set Rs=Execute("Select Count(*) From ["&Isforum&"Online] Where Eremite>-1 and forumid="&CID(a))
			Cache.Name = "Regforumidonline"& a
			Cache.Value = CID(Rs(0))
		End If
		Set Rs=Nothing
	End Sub

	'公告
  	Public Function Affiche()
    	Dim tmp,RS
    	Cache.Name="BBsAffiche"
    	Cache.Reloadtime = Cid(Forum_setting(44))
		If Cache.ObjIsEmpty() Then
	   		Set Rs=Execute("Select ID,Affichetitle,Affichecontent,Afficheman,Affichetime,Afficheinfo From ["&Isforum&"affiche] Order By AfficheTime Desc")
	   		If RS.Eof And Rs.Bof Then
				tmp = "暂无公告"
			Else
				tmp = Rs.GetRows(-1)
	   		End If
			Cache.Value = tmp
			RS.Close:Set RS=Nothing
		End If
   		Affiche = Cache.Value
  	End Function

	'友情链接
  	Public Function Forum_Link_Rs()
		Dim Rs
    	Cache.Name="Superlink"
    	Cache.Reloadtime = Cid(Forum_setting(44))
		If Cache.ObjIsEmpty() Then
	   		Set Rs=Execute("Select Name,Url,Logo,Intro,SetTops From ["&Isforum&"link] Order By SetTops Asc")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			RS.Close:Set RS=Nothing
		End If
		Forum_Link_Rs = Cache.Value
	End Function

	Function Forum_Link()
		Dim Value,i,tmp,tmp1,tmp2,tmp3
		Value = Forum_Link_Rs
		if isarray(value) Then
			tmp1 = "":tmp2 = ""
			for i = 0 to Ubound(Value,2)
				If Value(3,i)&"" = "" Then
					If Value(2,i) &"" = "" Then
						If tmp1 = "" Then
							tmp1 = "[<a href="""& Value(1,i) &""" target=""_blank"" title="""& Value(0,i) &""">"& Value(0,i) &"</a>]" 
						Else
							tmp1 = tmp1 & " [<a href="""& Value(1,i) &""" target=""_blank"" title="""& Value(0,i) &""">"& Value(0,i) &"</a>]" 
						End if
					Else
						If tmp2 = "" Then
							tmp2 = "<a href="""& Value(1,i) &""" target=""_blank"" title="""& Value(0,i) &"""><img src="& Value(2,i) &" border=""0"" Align=""absmiddle"" width=""88"" height=""31"" /></a>"
						Else
							tmp2 = tmp2 & " <a href="""& Value(1,i) &""" target=""_blank"" title="""& Value(0,i) &"""><img src="& Value(2,i) &" border=""0"" Align=""absmiddle"" width=""88"" height=""31"" /></a>"
						End if
					End if
				Else
					tmp3 =  tmp3& "	<tr class=""a4"">"
					tmp3 =  tmp3& "	<td width=""5%"" align=""center"" valign=""middle""><img src="""&Styleurl&"/link.gif"" alt="""" /></td>"
					tmp3 =  tmp3& "	<td width=""77%"" valign=""middle""> <a href="""& Value(1,i) &""" target=""_blank"" title="""& Value(0,i) &""" class=""bold""> "& Value(0,i) &" </a> <br> "& Value(3,i) &" </td>"
					tmp3 =  tmp3& "	<td width=""18%"" align=""center"" valign=""middle""> <img src="& Value(2,i) &" border=""0"" alt="""& Value(3,i) &""" width=""88"" height=""31""/> </td>"
					tmp3 =  tmp3& "	</tr> "
				End if
			Next
			tmp = tmp2 & "<br>" & tmp1
			Linkshows = tmp3
		End if
   		Forum_Link = tmp 
  	End Function

	'载入定制的在线人员列表
  	Public Function LoadOnlineShows()
    	Dim Tmp,RS
    	Cache.Name="OnlineShowsCache"
    	Cache.Reloadtime = Cid(Forum_setting(44))
		If Cache.ObjIsEmpty() Then
	   		Set Rs = execute("Select OnlineName,Onlineimg From ["&isforum&"OnlineTypes] Order By Sorts Asc")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			RS.Close:Set RS=Nothing
		End If
   		LoadOnlineShows = Cache.Value
  	End Function

	'首页显示定制在线列表人员分类
  	Public Function OnlineShows()
    	Dim Tmp,i,tmp1
		Tmp = LoadOnlineShows : tmp1 = ""
		If Isarray(tmp) Then
			for i=0 to Ubound(tmp,2)
				If tmp1 & "" = "" Then
					tmp1 = "<img src="""& StyleUrl & "/"&tmp(1,i)&""" alt="""&tmp(0,i)&""" /> "&tmp(0,i)&""
				Else
					tmp1 = tmp1& " &nbsp; &nbsp;<img src="""& StyleUrl & "/"&tmp(1,i)&""" alt="""&tmp(0,i)&""" /> "&tmp(0,i)&""
				End if
			Next
		End if
   		OnlineShows = tmp1
  	End Function

	Public Function ShowLines(a)
		Dim tmp,Rs,linetmp,u,p,i,OnlineTmp,SQL
		Cache.Name = "ShowLines"& a
		Cache.Reloadtime = Cid(Forum_setting(44))
		if Request("showlines")="no" Then Exit Function
		If team.Forum_setting(39)=0 And Request("showlines")<>"yes" Then Exit Function
		If Cache.ObjIsEmpty() Then
			If a = 0 Then
				SQL = "Select UserName,LevelName,IP,Bbsname,Acturl,Lasttime,Eremite From ["&Isforum&"Online] "
			Else
				SQL = "Select UserName,LevelName,IP,Bbsname,Acturl,Lasttime,Eremite From ["&Isforum&"Online] Where forumid = "& a
			End if
			Set Rs = Execute(SQL)
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			Rs.Close:Set Rs=Nothing
		End If
		tmp = Cache.Value
		Linetmp = LoadOnlineShows
		If IsArray(tmp) Then
			OnlineTmp = "<tr class=""a4"">" : p = 0
			For u = 0 to Ubound(tmp,2)
				p = p+1
				If Isarray(linetmp) Then
					for i=0 to ubound(linetmp,2)
						If Trim(linetmp(0,i))=Trim(tmp(1,u)) Then
							OnlineTmp = OnlineTmp & "<td>"
							If tmp(6,u) = 1 Then 
								OnlineTmp = OnlineTmp & "<img src="""& styleurl &"/line05.gif""  alt=""隐身用户"" align=""absmiddle""/> <span alt=""隐身用户"">隐身用户</span>"
							Else
								OnlineTmp = OnlineTmp & "<img src="""& styleurl &"/"&Linetmp(1,i)&"""  alt="""&Linetmp(0,i)&""" align=""absmiddle""/>" & IIF(CID(team.Forum_setting(65))=1,"<a href=""Profile-"&tmp(0,u)&".html"" title=""等级:"&tmp(1,u)&"&#xA;位置:"&tmp(3,u)&"&#xA;活动:"&formatdatetime(tmp(5,u),4)&"&#xA;"&Iif(SeeUIP,tmp(2,u),"....")&" ""> "&tmp(0,u)&" </a> </td>","<a href=""Profile.asp?username="&tmp(0,u)&""" title=""等级:"&tmp(1,u)&"&#xA;位置:"&tmp(3,u)&"&#xA;活动:"&formatdatetime(tmp(5,u),4)&"&#xA;"&Iif(SeeUIP,tmp(2,u),"....")&" ""> "&tmp(0,u)&" </a> </td>")
							End If	
						End if
					Next
				End if
				If p = 8 Then OnlineTmp = OnlineTmp & "</tr><tr class=""a4""> " : p = 0
			Next
		End If
		Showlines = OnlineTmp
	End Function

	Public Function UserOnlineinfos()	'判断用户状态
		Dim SQL,RS,tmp
		Cache.Name = "UserOnlineCache"
		Cache.Reloadtime = 10
		If Cache.ObjIsEmpty() Then
			Set Rs = Execute("Select UserName From ["&Isforum&"Online] Where Eremite = 0")
			If RS.Eof Then
				Exit Function
	   		Else
		   		Do While Not RS.Eof
		      		tmp = tmp & "$$"&Rs(0)&"$$"
	   				Rs.MoveNext
				Loop
				Cache.Value = tmp
			End if
			RS.Close:Set RS=Nothing
		End If
		UserOnlineinfos = Cache.Value
	End Function

	'导航菜单
  	Public function MenuTitle()
		Dim Tmp
		Tmp = Replace(ElseHtml(4),"{$clubname}",IIF(CID(team.Forum_setting(65))=1,"<a href=""Default.html""> " & Club_Class(1) &"</a>","<a href=""default.asp""> " & Club_Class(1) &"</a>"))
		Tmp = Replace(Tmp,"{$topic}",x1)
		Tmp = Replace(Tmp,"{$bbsname}",x2)
		Tmp = Replace(Tmp,"{$forumid}",IIF(CID(team.Forum_setting(65))=1,"Forums-"& Fid &".html","Forums.asp?fid="& Fid))
		MenuTitle = tmp
	End Function

	Public function LoadFooterAdvs()
		'LoadFooterAdvs = "<script language=""javascript"" src="""& ServerUrl &"/Server.Js""></script>"
	End Function

	'短讯通知
	Public function TeamNewMsg()
		Dim u,RS,MessTmp,tmp
		MessTmp = ""
		If Newmessage>0 then
			MessTmp = Replace(ElseHtml(5),"{$newmessage}",Newmessage)
			MessTmp = Replace(MessTmp,"{$msgwav}",IIf(Request.Cookies(Forum_Sn)("msgsound")="","<bgsound src=""images/plus/pm1.wav"">","<bgsound src=""images/plus/pm"&Request.Cookies(Forum_Sn)("msgsound")&".wav"">"))
			Set RS=Execute("Select top "&CID(Newmessage)&" msgtopic,Author,SendTime,ID From ["&Isforum&"message] Where Incept='"&TK_UserName&"' Order By ID Desc")
			Do While Not Rs.Eof
				tmp = tmp & "<li><a href=""Msg.asp?action=readmsg&sid="&Rs(3)&""" style=""cursor:hand"" target=""_blank"">短信内容: "&HtmlEncode(RS(0))&" - [来自: "&RS(1)&" / "&RS(2)&" ] </li>"
				Rs.Movenext
			Loop
			Rs.Close:Set Rs=Nothing
			MessTmp = Replace(MessTmp,"{$msgcontent}",IIF(tmp="","<a href=""Msg.asp"">您的消息因为长期未读，已被系统删除，请进入短信管理页面将未读短信提示数量清零。</a>",tmp))
		End if
		TeamNewMsg = MessTmp
	End function

	'无条件转向
	Public Sub Error(s)
		Response.clear
		Headers("系统信息提示")
		Echo "<div class=""tablemain"">" 
		Echo "	<div class=""threadmenu"">" 
		Echo "		<span class=""menulink""> <a href=""default.asp""> " & team.Club_Class(1) &" </a> &raquo; 系统信息提示</span>" 
		Echo "	</div>" 
		Echo "	<div class=""tableshow""><H3 class=""errorspan"">系统信息提示</H3>" 
		Echo "		<div class=""errormain"">" 
		Echo "			<Ul>" & s & " <hr>"
		Echo "				<li> <a href=""javascript:history.back()"">&lt;&lt; 点击这里返回 &gt;&gt;</a></li>"
		Echo "				<li> <a href=""default.asp"">&lt;&lt; 论坛首页 &gt;&gt;</a></li>"
		Echo "			</Ul>"
		Echo "		</div>" 
		Echo "	</div>" 
		Echo "</div>"
		footer
		Response.end
	End Sub
	'带条件转向
	Public Sub Error1(s)
		Response.clear
		team.Headers("系统信息提示")
		Echo "<div class=""tablemain"">" 
		Echo "	<div class=""threadmenu"">" 
		Echo "		<span class=""menulink""> <a href=""default.asp""> " & team.Club_Class(1) &" </a> &raquo; 系统信息提示</span>" 
		Echo "	</div>" 
		Echo "	<div class=""tableshow""><H3 class=""errorspan"">系统信息提示</H3>" 
		Echo "		<div class=""errormain"">" 
		Echo "			<Ul>" & s & " <hr>"
		Echo "				<li> <a href=""javascript:history.back()"">&lt;&lt; 点击这里返回 &gt;&gt;</a></li>"
		Echo "				<li> <a href=""default.asp"">&lt;&lt; 论坛首页 &gt;&gt;</a></li>"
		Echo "			</Ul>"
		Echo "				<P>请等待系统返回中.... [ <a id=""stime"">3</a> 秒]<script type=""text/javascript"">countDown(3);</script></P>"
		Echo "		</div>" 
		Echo "	</div>" 
		Echo "</div>"
		team.footer
		Response.end
	End Sub
	'弹出提示
	Public Sub Error2(Message)
		Response.clear
		Echo "<script>alert('"& Message &"');history.back();</script>"
		Echo "<script>window.close();</script>"
		Response.end
	End Sub
	'=========================================================================

	'检查验证码是否正确
	Public Function CodeIsTrue(a)
		Dim CodeStr
		CodeStr=Trim(a)
		If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
			CodeIsTrue=True
			Session("GetCode")=empty
		Else
			CodeIsTrue=False
			Session("GetCode")=empty
		End If	
	End Function

	Public Sub ChkPost()		'检测来源
		If Forum_setting(49) = 1 then
			Dim server_v1,server_v2,Chkpost,server_v3
			Chkpost=False 
			server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
			server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
			server_v3="dyn.www.rayfile.com"
			If Mid(server_v1,8,len(server_v3))=server_v3 Or Mid(server_v1,8,len(server_v2))=server_v2 Then Chkpost=True 
			If Chkpost = False Then error "来源错误"
		End if
	End Sub

	Public Sub LockIP()
		Dim IPlock,locklist
		Dim i,StrUserIP,StrKillIP
		Locklist = Club_Class(6)
		StrUserIP = RemoteAddr					'用户来源IP
		If StrUserIP & "" = "" Then Exit Sub
		StrUserIP=Split(StrUserIP,".")				'用户IP分段
		If Ubound(StrUserIP)<>3 Then Exit Sub
		If Trim(Locklist) &"" = "" Then
			Exit Sub
		Else
			If InStr(Locklist,Chr(13)&Chr(10)) >0 Then
				Locklist = Split(locklist,Chr(13)&Chr(10))
				For i= 0 to UBound(Locklist)
					Locklist(i)=Trim(Locklist(i))
					If Locklist(i)<>"" Then 
						StrKillIP = Split(Locklist(i),".")	'受限IP分段
						If Ubound(StrKillIP)<>3 Then Exit For
						IPlock = True
						If (StrUserIP(0) <> StrKillIP(0)) And Instr(StrKillIP(0),"*")=0 Then IPlock=False
						If (StrUserIP(1) <> StrKillIP(1)) And Instr(StrKillIP(1),"*")=0 Then IPlock=False
						If (StrUserIP(2) <> StrKillIP(2)) And Instr(StrKillIP(2),"*")=0 Then IPlock=False
						If (StrUserIP(3) <> StrKillIP(3)) And Instr(StrKillIP(3),"*")=0 Then IPlock=False
						If IPlock Then Exit For
					End If
				Next
			Else
				If Locklist <>"" Then 
					StrKillIP = Split(Locklist,".")	'受限IP分段
					If Ubound(StrKillIP)=3 Then
						IPlock = True
						If (StrUserIP(0) <> StrKillIP(0)) And Instr(StrKillIP(0),"*")=0 Then IPlock=False
						If (StrUserIP(1) <> StrKillIP(1)) And Instr(StrKillIP(1),"*")=0 Then IPlock=False
						If (StrUserIP(2) <> StrKillIP(2)) And Instr(StrKillIP(2),"*")=0 Then IPlock=False
						If (StrUserIP(3) <> StrKillIP(3)) And Instr(StrKillIP(3),"*")=0 Then IPlock=False
					End if
				End If				
			End if
			'判断Cookies更新目录
			Dim cookies_path_s,cookies_path_d,cookies_path
			cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
			cookies_path_d=ubound(cookies_path_s)
			cookies_path="/"
			For i=1 to cookies_path_d-1
				cookies_path=cookies_path&cookies_path_s(i)&"/"
			Next
			Response.Cookies(Forum_sn & "Kill").Expires = DateAdd("s", 360, Now())
			Response.Cookies(Forum_sn & "Kill").Path = cookies_path
			If IPlock Then
				Response.Cookies(Forum_sn & "Kill")("kill") = "1"
			Else
				Response.Cookies(Forum_sn & "Kill")("kill") = "0"
			End If
		End if
	End Sub

	Function myBoardJump()
		Dim RS,tmp,tmp1,i
		Cache.Name = "BoardJump"
		Cache.Reloadtime = Cid(Forum_setting(44))
		If Cache.ObjIsEmpty() Then
			Set Rs=Execute("Select ID,Bbsname,Followid From ["&Isforum&"bbsconfig] Where Hide=0 Order By SortNum")
	   		If RS.Eof Then
				Exit Function
			Else
	      		Cache.Value = Rs.GetRows(-1)
	   		End If
			Rs.Close:Set Rs=Nothing
		End If
		myBoardJump = Cache.Value
	End Function

	Function BoardJump()	
		Dim tmp1,i,Boards
		Boards = myBoardJump()
		tmp1 = "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}""><option value="""" selected>论坛跳转 ...</option>"
		If IsArray(Boards) Then
			For i = 0 To UBound(Boards,2)
				If Boards(2,i)=0 Then
					tmp1 = tmp1 & "<optgroup label="""&Boards(1,i) &""">"& BoardJump_Li(Boards(0,i),0)&"</optgroup>"
				End if
			Next
		End if
		tmp1 = tmp1 & " </select>"
		BoardJump = tmp1 
	End Function

	Function BoardJump_Li(a,b)
		Dim tmp1,i,Boards
		Dim U,Y
		Boards = myBoardJump()
		If isArray(Boards) Then
			For i=0 To Ubound(Boards,2)
				If Boards(2,i) = a Then
					U = 1+b
					tmp1 = tmp1 & IIF(CID(team.Forum_setting(65))=1,"<option value=""Forums-"&Boards(0,i)&".html"">","<option value=""Forums.asp?fid="&Boards(0,i)&""">")
					For Y=0 To U
						tmp1 = tmp1 & "&nbsp; &nbsp;"
					Next
					tmp1 = tmp1 & "&gt; "& Boards(1,i)&"</option>" 
					tmp1 = tmp1 & BoardJump_Li(Boards(0,i),U) 
				End if
			Next
		End if
		BoardJump_Li = tmp1
	End function

	Function BBs_Value_List(a,b)
		Dim tmp1,i,Boards
		Dim U,Y
		Boards = myBoardJump()
		If isArray(Boards) Then
			For i=0 To Ubound(Boards,2)
				If Boards(2,i) = a Then
					U = 1+b
					tmp1 = tmp1 & "<option value="""&Boards(0,i)&""">"
					For Y=0 To U
						tmp1 = tmp1 & "&nbsp; &nbsp;"
					Next
					If a = 0 Then
						tmp1 = tmp1 & "╋"
					Else
						tmp1 = tmp1 & "├"
					End if
					tmp1 = tmp1 & ""& Boards(1,i)&"</option>" & Vbcrlf
					tmp1 = tmp1 & BBs_Value_List(Boards(0,i),U) 
				End if
			Next
		End if
		BBs_Value_List = tmp1
	End Function
	
	Public Property Let SearcKeys(ByVal strPkey)
		SearcKeyword = strPkey
	End Property
	Public Property Let SearchClass(ByVal strPkey)
		SearcKeywordClass = strPkey
	End Property

	'PageList 总分页数,记录总数
	Function PageList(ByVal PageNum, ByVal AllNum, ByVal Mode)
		Dim Page,prevPage,nextPage,startPage,t,i,Acurl,iUrl,Lastp
		Dim PageRoot,PageFoot,iSearcKeys
		If AllNum<1 Then AllNum = 1 : iSearcKeys = false
		If PageNum<1 Then PageNum = 1
		Page = CheckNum(HRF(2,2,"page"),1,1,1,PageNum)
		Acurl = LCase(Request.ServerVariables("Query_String"))
		If HRF(2,1,"action") = "seachfile" Then iSearcKeys = True 
		If Not iSearcKeys Then

			If instr(Acurl,"page=") > 0 Then
				iUrl =  "?"& ReplacePage(Acurl) & "&amp;Page="
				If ReplacePage(Acurl) ="" Then iUrl = "?"& ReplacePage(Acurl) & "Page="
			Else 
				iUrl = "?Page="
				If Len(Acurl) > 1 Then  iUrl = "?"& ReplacePage(Acurl) & "&amp;Page="
			End If 
			If CID(team.Forum_setting(65))=1 and HRF(2,1,"filter") &""= "" Then  
				iUrl = Replacequest(iUrl)
				If InStr(Acurl,"fid")>0 or InStr(Acurl,"tid")>0 Then Lastp = ".html"
			End If 
		Else 
			If instr(Acurl,"page=") > 0 Then
				iUrl =  "?"& ReplacePage(Acurl) & "&amp;Page="
			Else 
				iUrl = "?" & ReplacePage(Acurl) & "&amp;searchclass="& SearcKeywordClass &"&searchkey="& SearcKeyword &"&amp;Page="
			End If
		End If 
		If Page <1 Then Page = 1 End If  : prevPage = Page - 1 : nextPage = Page + 1
		If Page - 5 <= 1 Then
			PageRoot = 1
		Else
			PageRoot = Page-5
		End If
		If Page + 5 >= PageNum Then
			PageFoot = PageNum
		Else
			PageFoot = Page + 5
		End If
		t = "<div class=""pages-nav"" style=""clear:left"">"
		If prevPage <= 1 Then
			t = t & "<span class=""next"">&#171; 首页</span><span class=""next"">&#139;</span>"
		Else
			t = t & "<a href="""& iUrl &"1"& Lastp &""" class=""next"">&#171; 首页</a><a href="""& iUrl &""& Page-1 & Lastp &""" class=""next"" title=""上一页"">&#139;</a>"
		End If
		If int(Page/10) = 0 Then
			startPage = Int(Page-9)
		Else
			startPage = Int(Page - ((Page/10)+1))
		End If
		If startPage <1 Then startPage = 1
		if startPage > 10 Then
			t = t & "<a href="""& iUrl &""& PageRoot-5 & Lastp  &""" class=""next"">...</a>"
		End If
		For i = PageRoot To PageFoot
			If i = Cint(Page) Then
				t = t & "<span class=""current"" title=""Page "& Cstr(i) &""">"& Cstr(i) &"</span>"
			Else 
				t = t & "<a href="""& iUrl &""& Cstr(i) & Lastp  &""" title=""Page "& Cstr(i) &""">"& Cstr(i)  &"</a>"
			End If
			If i = PageNum Then Exit For
		Next
		If PageNum >= (startPage+10) Then
			t = t & "<a href="""& iUrl &""& PageFoot+5 & Lastp  &""" class=""next"">...</a>"
		End If 
		If nextPage > PageNum Then
			t = t & "<span class=""next"">&#155;</span><span class=""next""> &#187;</span>"
		Else
			t = t & "<a href="""& iUrl &""& nextPage & Lastp  &""" class=""next"" title=""下一页"">&#155;</a><a href="""& iUrl &""& PageNum & Lastp  &""" class=""next"">尾页 &#187;</a>"
		End If 
		t = t & "<span class=""next"">"& AllNum &"/共"& PageNum &"页</span></div>"
		PageList = T
	End Function


	Function ReplacePage(s)
		Dim Re
		Set Re = New RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern = "[\?&]Page=([^&]*)"
		s=re.Replace(s,"")
		set re=Nothing
		ReplacePage = s
	End Function

	Function Replacequest(s)
		Dim sfid,stid,tmp
		stid = HRF(2,2,"tid")
		sfid = HRF(2,2,"fid")
		If stid > 0 And sfid > 0 Then
			tmp = "Archiver-"&sfid&"-"&stid&"-"
		ElseIf stid > 0 Then
			tmp = "thread-"&stid&"-"
		Elseif sfid > 0 Then 
			tmp = "forums-"&sfid&"-"
		Else
			tmp = "default"
		End if
		Replacequest = tmp
	End Function




	'记录查询错误事件
	Public Sub SaveLOG(msg)
		Dim lConnStr,lConn,ldb
		ldb = MyDbPath & LogDate
		lConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ldb)
		Set lConn = Server.CreateObject("ADODB.Connection")
		lConn.Open lConnStr
		lConn.Execute("Insert Into SaveLog (UserName,IP,Windows,Remark,Logtime) Values ('"&TK_UserName&"','"&RemoteAddr&"','"&Left(HtmlEncode(Request.Servervariables("HTTP_USER_AGENT")),255)&"','"&Left(HtmlEncode(msg),255)&"','"&Now&"')")
		lConn.Close
		Set lConn = Nothing 
	End Sub

	Public Function address(sip)
		Dim aConnStr,aConn,adb
		Dim str1,str2,str3,str4
		Dim  num
		Dim country,city
		Dim irs,SQL
		address="未知"
		If IsNumeric(Left(sip,2)) Then
			If sip="127.0.0.1" Then sip="192.168.0.1"
			str1=Left(sip,InStr(sip,".")-1)
			sip=mid(sip,instr(sip,".")+1)
			str2=Left(sip,instr(sip,".")-1)
			sip=Mid(sip,InStr(sip,".")+1)
			str3=Left(sip,instr(sip,".")-1)
			str4=Mid(sip,instr(sip,".")+1)
			If isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 Then
			Else		
				num=CLng(str1)*16777216+CLng(str2)*65536+CLng(str3)*256+CLng(str4)-1
				adb = IPDate
				aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(adb)
				Set AConn = Server.CreateObject("ADODB.Connection")
				aConn.Open aConnStr
				country="亚洲"
				city=""
				sql="select top 1 country,city from tm_address where ip1 <="& num &" and ip2 >="& num 
				Set irs=aConn.execute(sql)
				If Not(irs.EOF And irs.bof) Then
					country=irs(0)
					city=irs(1)
				End If
				irs.close
				Set irs=Nothing
				aConn.Close
				Set aConn = Nothing 
				SqlQueryNum = SqlQueryNum+1
			End If
			address=country&city
		End If
	End Function
	'是否真正的搜索引擎
	Public Function IsWebSearch()
		IsWebSearch = False
		Dim Botlist,i,Agent
		Agent = Request.ServerVariables("HTTP_USER_AGENT")
		Botlist=Array("Google,Isaac,SurveyBot,Baiduspider,ia_archiver,P.Arthur,FAST-WebCrawler,Java,Microsoft-ATL-Native,TurnitinBot,WebGather,Sleipnir")
		For i=0 to UBound(Botlist)
			If InStr(Agent,Botlist(i))>0  Then 
				IsWebSearch = True
				Exit For
			End If
		Next 
	End Function

	Public Function BuildFile(ByVal sFile, ByVal sContent)
		Dim is_gb2312
        Dim oFSO, oStream
		If Int(Forum_setting(65)) = 0 Then Exit Function
		is_gb2312 = 1
        If is_gb2312 = 1 Then
            Set oFSO = server.CreateObject("Scripting.FileSystemObject")
			sFile=Server.MapPath(sFile)
            Set oStream = oFSO.CreateTextFile(sFile, True)
            oStream.Write sContent
            oStream.Close
            Set oStream = Nothing
            Set oFSO = Nothing
        Else
            Set oStream = server.CreateObject("ADODB.Stream")
            With oStream
                .Type = 2
                .Mode = 3
                .Open
                .Charset = "gb2312"
                .Position = oStream.size
				.Write = sContent
                .SaveToFile sFile, 2
                .Close
            End With
            Set oStream = Nothing
        End If
    End Function

	Public Function LoadServerInfo()
		Dim tmp
		Cache.Name = "BoardServerinfos"
		Cache.Reloadtime = "1440"
		If Cache.ObjIsEmpty() Then
			tmp = GetUrlXmls( ServerUrl & "ajax.xml" )
			cache.value = tmp
		End If
		LoadServerInfo = cache.value
	End Function

	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function

   	Public Function Execute(SQL)
		If Not IsObject(Conn) Then ConnectionDatabase
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(SQL)
			If Err Then
				Err.Clear
				Set Conn = Nothing
				Response.Write "数据查询错误，请检查您的查询代码是否正确。"
				Response.End
			End If
		Else
			Response.Write SQL & "<br>"
			Set Execute = Conn.Execute(SQL)
		End If
		SqlQueryNum = SqlQueryNum+1
   	End Function

	'释放
	Public Sub Htmlend
		If IsArray(Group_Browse) Then
			If CID(Group_Browse(0)) = 0 Then	
				Response.Redirect "Club.asp?message=您没有查看论坛的权限。"
			End If
		End If 
		Set team = Nothing
		Set Cache = Nothing
		Set conn = Nothing
		Response.End
	End sub

	'类注销
	Private Sub Class_Terminate()
		Err.Clear
		If IsObject(Conn) Then Conn.Close:Set Conn=Nothing
		If IsObject(Cache) Then Cache.Close:Set Cache=Nothing
		If IsObject(team) Then team.Close:Set team=Nothing
		Response.End
	End Sub
End Class

Class Cls_Cache
	'缓存类 By DV
	Public Reloadtime,MaxCount
	Private LocalCacheName,CacheData,DelCount
	Private Sub Class_Initialize()
		Reloadtime=14400	'定义默认更新时间
	End Sub
	Private Sub SetCache(SetName,NewValue)
		Application.Lock	'锁定
		Application(SetName) = NewValue		'赋值
		Application.unLock	'解除锁定
	End Sub 
	Public Sub MakeEmpty(MyCaheName)
		Application.Lock	'锁定
		Application(CacheName&"_"&MyCaheName) = Empty	'清除缓存
		Application.unLock	'解除锁定
	End Sub 
	Public  Property Let Name(ByVal vNewValue) 'ByVal关键字,vNewValue自定义变量
		LocalCacheName=LCase(vNewValue)	'设置类变量Name
	End Property
	Public  Property Let Value(ByVal vNewValue)	'设置类变量Value
		If LocalCacheName<>"" Then 
			CacheData=Application(CacheName&"_"&LocalCacheName)
			If IsArray(CacheData)  Then
				CacheData(0)=vNewValue
				CacheData(1)=Now()
			Else
				ReDim CacheData(2)
				CacheData(0)=vNewValue
				CacheData(1)=Now()
			End If
			SetCache CacheName&"_"&LocalCacheName,CacheData
		Else
			Err.Raise vbObjectError + 1, "CacheServer", "请修改CacheName名称"
		End If		
	End Property
	Public Property Get Value()	'Value取值
		If LocalCacheName<>"" Then 
			CacheData=Application(CacheName&"_"&LocalCacheName)	
			If IsArray(CacheData) Then
				Value=CacheData(0)
			Else
				Err.Raise vbObjectError + 1, "CacheServer", " The CacheData Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "CacheServer", " please change the CacheName."
		End If
	End Property
	Public Function ObjIsEmpty()	'检测是否为空
		ObjIsEmpty=True
		CacheData=Application(CacheName&"_"&LocalCacheName)
		If Not IsArray(CacheData) Then Exit Function
		If Not IsDate(CacheData(1)) Then Exit Function
		If DateDiff("s",CDate(CacheData(1)),Now()) < 60*Reloadtime  Then
			ObjIsEmpty=False
		End If
	End Function
	Public Sub DelCache(MyCaheName)	'删除缓存
		Application.Lock
		Application.Contents.Remove(CacheName&"_"&MyCaheName)
		Application.unLock
	End Sub
End Class
%>