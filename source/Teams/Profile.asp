<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim UserName,x1,x2,Fid
TestUser()
UserName = HRF(2,1,"username")
Call Userinfos
team.Footer()

Sub Userinfos
	Dim Temp,Rs,tmp,Ump
	Set Rs = Team.Execute("Select ID,UserName,Degree,UserGroupID,Levelname,Usermail,Userhome,Userface,UserCity,UserSex,Question,Answer,Honor,Birthday,Sign,Medals,UserInfo,Posttopic,Postrevert,Deltopic,Goodtopic,Regtime,Landtime,Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7,Members From ["&Isforum&"User] Where UserName='"&username&"'")
	'ID=0,UserName=1,Degree=2,UserGroupID=3,Levelname=4,Usermail=5,Userhome=6,Userface=7,UserCity=8,UserSex=9,Question=10,Answer=11,Honor=12,Birthday=13,Sign=14,Medals=15,UserInfo=16,Posttopic=17,Postrevert=18,Deltopic=19,Goodtopic=20,Regtime=21,Landtime=22,Extcredits0=23,Extcredits1=24,Extcredits2=25,Extcredits3=25,Extcredits4=26,Extcredits5=27,Extcredits6=28,Extcredits7=28,Members=30	
	If Rs.Eof and RS.Bof Then
		Team.Error "您查询的用户不存在"
	Else
		Temp = Rs.GetString(,1, "$#$","","")
	End If
	Rs.Close:Set Rs=Nothing
	tmp = Split(Temp,"$#$")
	team.Headers(tmp(1)&"用户的资料")
	X1 = "<a href=""Profile.asp?username="&UserName&""">查看用户"&UserName&"的资料</a>"
	Ump = Replace(Team.UserHtml(0),"{$weburl}",team.MenuTitle)
	Ump = Replace(Ump,"{$username}",tmp(1))
	Ump = Replace(Ump,"{$uid}",tmp(0))
	Ump = Replace(Ump,"{$regtime}",tmp(21))
	Ump = Replace(Ump,"{$landtime}",tmp(22))
	Ump = Replace(Ump,"{$posttopic}",tmp(17))
	Ump = Replace(Ump,"{$postrevert}",tmp(18))
	Ump = Replace(Ump,"{$userface}",Fixjs(tmp(7)))
	Ump = Replace(Ump,"{$goodtopic}",tmp(20))
	Ump = Replace(Ump,"{$degree}",CID(tmp(2)))
	Ump = Replace(Ump,"{$showdegree}",UserOnlinetimes(CID(tmp(2))))
	Ump = Replace(Ump,"{$mypostcont}",tmp(17)+CID(tmp(18)))
	If Not IsNumeric(Application(CacheName&"_ConverPostNum")) Or Application(CacheName&"_ConverPostNum")<1 Then
		Application(CacheName&"_ConverPostNum") = 1
	End If
	Dim MyTimes
	Mytimes = Datediff("d",CDaTe(tmp(21)),Now())
	If Mytimes <=0 Then Mytimes = 1
	Ump = Replace(Ump,"{$mypostcontall}",FormatNumber((tmp(17)+CID(tmp(18)))/CID(Application(CacheName&"_ConverPostNum")),3))
	Ump = Replace(Ump,"{$daypost}",iif(FormatNumber((CID(tmp(17))+CID(tmp(18)))/Mytimes,3)<1,"0"& FormatNumber((CID(tmp(17))+CID(tmp(18)))/Mytimes,3),FormatNumber((CID(tmp(17))+CID(tmp(18)))/Mytimes,3)))
	Ump = Replace(Ump,"{$levelname}",Split(tmp(4),"||")(0) &"  " & UserStar(Split(tmp(4),"||")(3)))
	If tmp(9) = 1 Then
		Ump = Replace(Ump,"{$sex}","男性")
	Elseif tmp(9) = 2 Then
		Ump = Replace(Ump,"{$sex}","女性")
	Else
		Ump = Replace(Ump,"{$sex}","未知")
	End if
	Ump = Replace(Ump,"{$city}",tmp(8))
	Ump = Replace(Ump,"{$birthday}",tmp(13))
	Ump = Replace(Ump,"{$userhome}",tmp(6))
	Ump = Replace(Ump,"{$usermail}",tmp(5))
	Ump = Replace(Ump,"{$userqq}",Split(tmp(16),"|")(0))
	Ump = Replace(Ump,"{$icq}",Split(tmp(16),"|")(1))
	Ump = Replace(Ump,"{$yahoo}",Split(tmp(16),"|")(2))
	Ump = Replace(Ump,"{$msn}",Split(tmp(16),"|")(3))
	Ump = Replace(Ump,"{$taobao}",Split(tmp(16),"|")(4))
	Ump = Replace(Ump,"{$isbuy}",Split(tmp(16),"|")(5))
	Ump = Replace(Ump,"{$sign}",Sign_Code(tmp(14),CID(Split(tmp(4),"||")(4))))
	Dim U,Emp,ExtCredits
	ExtCredits = Split(team.Club_Class(21),"|")
	for u = 0 to ubound(ExtCredits)
		If Split(ExtCredits(u),",")(4) =1 Then
			emp = emp & " <tr><td> "& Split(ExtCredits(u),",")(0) & ":</td><td>"& tmp(23 + u) &"&nbsp;"& Split(ExtCredits(u),",")(1) &" </td></tr>"
		End if
	Next
	Ump = Replace(Ump,"{$userext}",emp)
	Dim Us,Post
	Set Us = Team.Execute("select top 5 ID,topic,Posttime,replies,Views,Goodtopic,Lasttime from ["&IsForum&"forum] Where Deltopic=0 and UserName = '"&UserName&"' Order By PostTime Desc")
	Do While Not Us.Eof
		Post = Post & "<tr class=""tab4""><td width=""40%"" align=""left"">"
		If Us(5)=1 Then
			Post = Post & " [精华] "
		End if
		Post = Post & " <a href=""thread.asp?tid="&Us(0)&""" target=""_blank"">"&Us(1)&"</a></td><td> "&Us(4)&" / "&Us(3)&"</td><td>"&Us(2)&"</td><td>"&Us(6)&"</td></tr>"
		Us.MoveNext	
	Loop
	Us.Close:Set Us=Nothing
	Ump = Replace(Ump,"{$todaypost}",Post)
	Echo Ump
End Sub 

%>