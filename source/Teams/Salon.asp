<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Public mRs
Dim UserName,fID,x1,x2,tmp,SQL,i,USign
Dim IsPage,Page,RS,Maxpage,PageNum,mmp
If HRF(2,1,"username")=Empty Then
	Call Main()
Else
	Call showMain()
End If 

Sub Main()
	Dim log_Year,log_Month,log_Day,Userface,UserOnlineinfos
	UserOnlineinfos = team.UserOnlineinfos
	log_Year = Trim(Request.QueryString("log_Year"))
	log_Month = Trim(Request.QueryString("log_Month"))
	log_Day = Trim(Request.QueryString("log_Day"))
	UserName = HRF(2,1,"username")
	team.Headers(Team.Club_Class(1) & "个人文集")
	X1="<a href=""Salon.asp"">个人文集</a>"
	X2=""
	tmp = team.MenuTitle
	IsPage = team.execute("Select Count(ID) From ["&IsForum&"Forum] Where Creatdiary = 1")(0)
	SQL="Select ID,Topic,UserName,Views,Replies,Posttime,Icon,Content From ["&IsForum&"Forum] Where Creatdiary = 1 Order By PostTime Desc"
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum=SqlQueryNum+1
		Maxpage = 20								'每页分页数
		PageNum = Abs(int(-Abs(IsPage/Maxpage)))	'页数
		Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		mRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
	tmp = tmp & "<table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""98%"" align=""center"" class=""a2"">"
	tmp = tmp & "<tr><td class=""a1"">个人文集</td></tr>"
	If Not Isarray(mRs) Then
		tmp = tmp & "<tr><td class=""a4"">暂无个人文集</td></tr>"
	Else
		For i=0 To Ubound(mRs,2)
			mmp = mmp & "<tr class=""a4"">"
			mmp = mmp & " <td> <a href=""salon.asp?username="&mRS(2,i)&""" title=""查看"&mRS(2,i)&"的文集"">[来自："&mRS(2,i)&"的文集] <a href=""Thread.asp?tid="&mRS(0,i)&""" target=""_blank""><B>"& mRS(1,i) &"</B></a> <BR><div align=""right""> [ 发表时间："&mRS(5,i)&" | 查看："&mRS(3,i)&" | 评论："&mRS(4,i)&""
			If Cid(mRs(6,i))>0 Then mmp = mmp & " | <img src=""images/brow/icon"&mRs(6,i)&".gif"" align=""absmiddle"">"
			mmp = mmp & " ]</div> </td></tr>"

		Next
		tmp = tmp & mmp
	End If
	tmp = tmp & "</table><br><table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""98%"" align=""center"">"
	tmp = tmp & "<tr><td>"
    tmp = tmp & "  <script language=""JavaScript"">"
	tmp = tmp & "       var pg = new showPages('pg'); "
    tmp = tmp & "		pg.pageCount ="&PageNum&";  "
	tmp = tmp & "		pg.dispCount ="&IsPage&";  "
    tmp = tmp & "		pg.printHtml(1); "
    tmp = tmp & "  </script>"
	tmp = tmp & "</td></tr></table>"
	Echo tmp
End Sub

Sub showMain()
	Dim log_Year,log_Month,log_Day,Userface,UserOnlineinfos
	UserOnlineinfos = team.UserOnlineinfos
	log_Year = Trim(Request.QueryString("log_Year"))
	log_Month = Trim(Request.QueryString("log_Month"))
	log_Day = Trim(Request.QueryString("log_Day"))
	UserName = HRF(2,1,"username")
	team.Headers(Team.Club_Class(1) & " - " & UserName & "的个人文集")
	IsPage = team.execute("Select Count(ID) From ["&IsForum&"Forum] Where Creatdiary = 1 and UserName='"&UserName&"'")(0)
	SQL="Select ID,Topic,UserName,Views,Replies,Posttime,Icon,Content From ["&IsForum&"Forum] Where Creatdiary = 1 and UserName='"&UserName&"' Order By PostTime Desc"
	Set Rs = Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1,&H0001
	If Not (Rs.Eof and Rs.Bof) Then 
		SqlQueryNum=SqlQueryNum+1
		Maxpage = 20								'每页分页数
		PageNum = Abs(int(-Abs(IsPage/Maxpage)))	'页数
		Page = CheckNum(Request.QueryString("page"),1,1,1,PageNum)	'当前页
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		mRs=Rs.GetRows(Maxpage)
	End if
	RS.Close:Set Rs=Nothing
	Set Rs = team.execute("Select Sign,Userface From ["&IsForum&"User] Where UserName='"&UserName&"'")
	If Rs.Eof And Rs.bof Then
		team.Error "系统不存在此用户的个人文集。"
	Else
		USign = Rs(0) : Userface = RS(1)
	End If 
	Rs.close:Set Rs=Nothing 
	x1 = "<a href=""Salon.asp?username="&UserName&"""> 查看 " & UserName & " 的个人文集 </a>"
	x2 = "<a href=""Salon.asp""> 查看所有文集 </a>"
	tmp = Replace(Team.UserHtml (3),"{$username}",UserName)
	tmp = Replace(tmp,"{$weburl}",team.MenuTitle)
	tmp = Replace(tmp,"{$onlinfo}",Iif(InStr(UserOnlineinfos,"$$"&UserName&"$$")>0, "[<FONT COLOR=""red"">在线</FONT>]","[离线]"))
	tmp = Replace(tmp,"{$imgface}",iif(Userface&""="","","<img src="""&Userface&""" border=""0"" onload='javascript:if(this.width>"&CID(team.Forum_setting(108))&")this.width="&CID(team.Forum_setting(108))&";if(this.height>"&CID(team.Forum_setting(109))&")this.height="&CID(team.Forum_setting(109))&";'onerror='javascript:this.src=""images/face/error.gif""'><br>"))
	tmp = Replace(tmp,"{$mybookdate}",Calendar(log_Year,log_Month,log_Day))
	tmp = Replace(tmp,"{$myinfos}",Sign_Code(USign,1))
	If Not Isarray(mRs) Then
		tmp=Replace(tmp,"{$showlives}","<tr class=""tab3""><td> <h4> 此用户暂无撰写文集 。</h4> </td></tr>")
	Else
		For i=0 To Ubound(mRs,2)
			mmp = mmp & "<tr class=""popupmenu_option"">"
			mmp = mmp & " <td> <a href=""Thread.asp?tid="&mRS(0,i)&""" target=""_blank""><B>"& mRS(1,i) &"</B></a><br> [ 时间："&mRS(5,i)&" | 查看："&mRS(3,i)&" | 评论："&mRS(4,i)&""
			If Cid(mRs(6,i))>0 Then mmp = mmp & " | <img src=""images/brow/icon"&mRs(6,i)&".gif"" align=""absmiddle"">"
			mmp = mmp & " ] </td></tr>"
			mmp = mmp & " <tr class=""a4""><td height=""50"" valign=""top""> "& Cutstr(mRS(7,i),200)&" </td>"
			mmp = mmp & "</tr>"
		Next
		tmp=Replace(tmp,"{$showlives}",mmp)
	End if
	tmp = Replace(tmp,"{$TotalPage}",IsPage)
	tmp = Replace(tmp,"{$allpage}",PageNum)	
	Echo tmp
End Sub



Function Calendar(C_Year,C_Month,C_Day)  'BLOG日历
	Dim tmp
	ReDim Link_Days(2,0)
	Dim Link_Count
	Link_Count=0
	Dim This_Year,This_Month,This_Day,RS_Month,Link_TF
	IF C_Year=Empty Then C_Year=Year(Now())
	IF C_Month=Empty Then C_Month=Month(Now())
	IF C_Day=Empty Then C_Day=0
	C_Year=Cint(C_Year)
	C_Month=Cint(C_Month)
	C_Day=Cint(C_Day)
	This_Year=C_Year
	This_Month=C_Month
	This_Day=C_Day
	Dim To_Day,To_Month,To_Year
	To_Day=Cint(Day(Now()))
	To_Month=Cint(Month(Now()))
	To_Year=Cint(Year(Now()))
	Dim the_Day,ismytime
	the_Day=0
	If IsArray(mRs) Then
		For i=0 To Ubound(mRs,2)
			ismytime= ""
			ismytime = Split(FormatDateTime(mRs(5,i),2),"-")
			IF ismytime(2)<>the_Day Then
				the_Day=ismytime(2)
				ReDim PreServe Link_Days(2,Link_Count)
				Link_Days(0,Link_Count)=Cint (ismytime(1))
				Link_Days(1,Link_Count)=Cint (ismytime(2))
				Link_Days(2,Link_Count)="Salon.asp?username="&UserName&"&log_Year="&ismytime(0)&"&log_Month="&ismytime(1)&"&log_Day="&ismytime(2)
				Link_Count=Link_Count+1
			End If
		Next
	End if

	Dim Month_Name(12)
	Month_Name(0)=""
	Month_Name(1)="一"
	Month_Name(2)="二"
	Month_Name(3)="三"
	Month_Name(4)="四"
	Month_Name(5)="五"
	Month_Name(6)="六"
	Month_Name(7)="七"
	Month_Name(8)="八"
	Month_Name(9)="九"
	Month_Name(10)="十"
	Month_Name(11)="十一"
	Month_Name(12)="十二"
	
	Dim Month_Days(12)
	Month_Days(0)=""
	Month_Days(1)=31
	Month_Days(2)=28
	Month_Days(3)=31
	Month_Days(4)=30
	Month_Days(5)=31
	Month_Days(6)=30
	Month_Days(7)=31
	Month_Days(8)=31
	Month_Days(9)=30
	Month_Days(10)=31
	Month_Days(11)=30
	Month_Days(12)=31
	If IsDate("February 29, " & This_Year) Then Month_Days(2)=29
	Dim Start_Week
	Start_Week=WeekDay(C_Month&"-1-"&C_Year)-1
	Dim Next_Month,Next_Year,Pro_Month,Pro_Year
	Next_Month=C_Month+1
	Next_Year=C_Year
	IF Next_Month>12 then 
		Next_Month=1
		Next_Year=Next_Year+1
	End IF
	Pro_Month=C_Month-1
	Pro_Year=C_Year
	IF Pro_Month<1 then 
		Pro_Month=12
		Pro_Year=Pro_Year-1
	End IF
	tmp = "<div id=""panelCalendar""><table width=""98%"" id=""calendar"" cellspacing=""1"" cellpadding=""4"" class=""a2"" align=""center""><tr><td colspan=""7"" class=""a1"" align=""center""> 日历 </td></tr><tr><td colspan=""7"" class=""a3"" align=""center""><a href=""Salon.asp?username="&UserName&"&log_Year="&C_Year-1&""" title=""上一年"">&laquo;</a><span>&nbsp;"&C_Year&"&nbsp;</span><a href=""Salon.asp?username="&UserName&"&log_Year="&C_Year+1&""" title=""下一年"">&raquo;</a>&nbsp;&nbsp;<a href=""Salon.asp?username="&UserName&"&log_Year="&Pro_Year&"&log_Month="&Pro_Month&""" title=""上一月"">&laquo;</a><span>&nbsp;"&Month_Name(C_Month)&"月&nbsp;</span><a href=""Salon.asp?username="&UserName&"&log_Year="&Next_Year&"&log_Month="&Next_Month&""" title=""下一月"">&raquo;</a></td></tr><tr>"
	tmp = tmp & "<td  class=""blog_a1"">Su</td><td class=""a4"">Mo</td><td class=""a3"">Tu</td><td class=""a4"">We</td><td class=""a3"">Th</td><td class=""a4"">Fr</td><td class=""blog_a2"">Sa</td></tr><tr>"
	Dim i,j,k,l,m
	For  i=0 TO Start_Week-1
		tmp = tmp & "<td class=""a4"">&nbsp;</td>"
	Next
	Dim This_BGColor
	j=1
	While j<=month_Days(This_Month)
	 	For k=start_Week To 6
			This_BGColor="a4"
			'双修日加字体粗 2005-09-05 0为星期天 6为星期六
			If k=0 Then This_BGColor="blog_a1"
			If k=6 Then This_BGColor="blog_a2"
			IF j=To_Day AND This_Year=To_Year AND This_Month=To_Month Then This_BGColor="blog_a4"
			IF j=This_Day Then This_BGColor="blog_a3"
			tmp = tmp & "<td class="""&This_BGColor&""">"
			Link_TF="Flase"
			For l=0 TO Ubound(Link_Days,2)
				IF Link_Days(0,l)<>"" Then
					IF Link_Days(0,l)=This_Month AND Link_Days(1,l)=j Then
						tmp = tmp & "<b><a href="""&Link_Days(2,l)&""">"
						Link_TF="True"
					End IF
				End IF
			Next
		IF j<=Month_Days(This_Month) Then tmp = tmp & (j)
		IF Link_TF="True" Then tmp = tmp & "</a></b>"
        tmp = tmp & "</td>"
		j=j+1
	Next
	Start_Week=0
	tmp = tmp & "</tr>"
	Wend
	tmp = tmp & "</table></div><BR>"
	Calendar = tmp
End Function


team.footer
%>
