<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim Fid,x1,x2
team.Headers("�ļ�����!")
X1="<a href=Search.asp>�ļ�����</A>"
X2=""
echo team.MenuTitle
Call TestUser()
Select Case Request("action")
	Case "seachfile"
		Call seachfile
	Case Else
		Call Main
End Select
Team.footer

Sub seachfile
	Call Main()
	Dim searchkey,Rs,StrSQL,SClass,AllCount,PageNum
	Dim IsWhere,nTop,Page,Shows,Maxpage,i
	Page = HRF(2,2,"Page")
	Sclass = CID(Trim(Request("searchclass")))
	searchkey = HtmlEncode(Trim(Request("searchkey")))
	team.SearcKeys = searchkey
	team.SearchClass = Sclass
	if team.Group_Browse(4) < 1 then 
		team.Error(" �����ڵ��� "&team.levelname(0)&" û��������Ȩ��.")
	End If
	If Sclass = 0 Or Not IsNumeric(Sclass) Then
		team.Error "������������"
	End If

	If Sclass = 1 Then
		If Not IstrueName(searchkey) Then
			team.Error "�������������"
		End If
		If searchkey &"" = "" Then
			team.Error "�������ݲ���Ϊ��"
		End If
	ElseIf Sclass = 2 Then
		If searchkey &"" = "" Then
			team.Error "�������ݲ���Ϊ��"
		End If
	ElseIf Sclass = 3 Then
	Else
		team.Error "�������������"
	End If 
	If Page=0 Then
		Cache.Reloadtime = 1
		Cache.Name="Usearch"
		If Cache.ObjIsEmpty() Then
			Cache.value = 1
		Else
			Cache.value = Cache.value +1
		End if
		If Cache.value > team.Forum_setting(52) Then
			team.Error "������ͬʱ����������ϵͳ���ã����Ժ���������������"
		End If
		Call CheckpostTime
		Session(CacheName &"searchtime") = Now()
	End If
	If Sclass = 1 Then
		IsWhere = " UserName like '%" & searchkey & "%' " 
	ElseIf Sclass = 2 Then
		If IsSqlDataBase = 1 Then
			IsWhere = " Topic like '%" & searchkey & "%' "
		Else
			IsWhere = " InStr(1,LCase(Topic),LCase('"&searchkey&"'),0)<>0 "
		End if
	ElseIf Sclass = 3 and team.Group_Browse(4) = 2 Then
		If IsSqlDataBase = 1 Then
			IsWhere = " Content like '%" & searchkey & "%' "
		Else
			IsWhere = " InStr(1,LCase(Content),LCase('"&searchkey&"'),0)<>0 "
		End If
	ElseIf Sclass = 4 Then
		IsWhere = " Goodtopic = 1 "
	End If
	If Sclass = 3 Then
		nTop = " Top 10 "
	Else
		nTop = ""
	End If
	AllCount = team.Execute("Select Count(ID) from ["&IsForum&"forum] where deltopic=0 and Locktopic=0 and "&IsWhere&" ")(0)
	If AllCount = "" Or Not IsNumEric(AllCount) Then 
		AllCount = 0
	End if
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	UpdateUserpostExc()
	StrSQL = "Select "&nTop&" ID,forumid,Topic,Username,Views,Icon,Replies,lasttime,Goodtopic,Createpoll,Creatdiary,Creatactivity,Rewardprice,Readperm,Rewardpricetype from ["&IsForum&"forum] where deltopic=0 and Locktopic=0 and "&IsWhere&" Order By Lasttime DESC"
	Rs.Open StrSQL,Conn,1,1,&H0001
	Response.Write "<table cellspacing=1 cellpadding=10 width=98% align=center border=0><tr class=a3><td colspan=2 align=center>�����������ҵ� <Font color=red>"&AllCount&"</Font> ��������Ӽ�¼</td></tr>"
	If Rs.Eof And Rs.Bof Then
		Echo " <tr class=a4><td colspan=2 align=center> �Բ��𣬱�վû���ҵ���Ҫ��ѯ�����ݣ����뵽�ٶ�ȥ�������� <B>["&searchkey&"]</B> ����Ϣô�� <a href=""http://www.baidu.com/baidu?tn=team5&word="&searchkey&""" target=""_blank"">�������������硿</a> </td></tr></table>"
	Else
		Maxpage = 20
		PageNum = Abs(int(-Abs(AllCount/Maxpage)))	'ҳ��
		Page = CheckNum(Page,1,1,1,PageNum)	'��ǰҳ
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	Dim Un,tmp,ExtCredits,bbsname,Chcheid,j
	ExtCredits = Split(team.Club_Class(21),"|")
	Chcheid = team.BoardList
	For i=0 To Ubound(shows,2)
		If Request("Page")<2 Then
			Un=i+1
		Else
			Un=Int(Request("Page"))*Maxpage-Maxpage+i+1
		End If
		For j=0 to Ubound(Chcheid,2)
			If Cid(Shows(1,i)) = Cid(Chcheid(0,j)) Then
				BBsName = Chcheid(1,j)
			End if
		Next
		Echo "<tr class=a4><td height=50>"&Un&"."&iif(Shows(8,i)=1,"<img src="""&team.styleurl&"/f_good.gif"" border=""0"" align=""absmiddle"" alt=""����"" >","")&" "
		Echo "<a Href=Thread.asp?tid="&Shows(0,i)&" target=""_blank""> "&GetColor(Shows(2,i),searchkey)&" "&iif(Cid(Shows(13,i))>0,"- [<b>�Ķ�Ȩ��</b> "&Shows(13,i)&"]","")&" "&iif(Cid(Shows(10,i))>0,"- [<b>�û��ļ�</b>]","")&" "&iif(Cid(Shows(11,i))>0,"- [<b>��ټ�</b>]","")&" "&IIf(Cid(Shows(14,i))=0,iif(Cid(Shows(14,i))>0,"- [<b>���� </b> "&IIF(Split(ExtCredits(Cid(team.Forum_setting(99))),",")(3)=1,  "  "& Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)&" "&Shows(14,i)&" "," ������δ���� ")&"]",""),"[�ѽ��]")&" "&iif(DateDiff("d",Shows(7,i),date())=0,"  <img src="""&team.styleurl&"/new.gif"" border=""0"" align=""absmiddle"">","")&" </a> "
		Echo " <BR /> <font color=""green"">���ߣ�"&GetColor(Shows(3,i),searchkey)&" �����"&Shows(4,i)&"  �ظ���"&Shows(6,i)&"  ��   "
		Echo " <a Href=Thread.asp?tid="&shows(0,i)&" target=""_blank"">"&BBsName&"</a> "&Shows(7,i)&" </Font> "
		Echo " <hr style=""border-top:1px #B3B3B3 dashed;border-bottom:0px;height:0px;width:98%;""></hr></td> </tr>"
	Next
	Echo "</table><BR>"
	Echo "<div id=""pagediv"">"& team.PageList(PageNum,AllCount,6) &"</div><BR><div id=""rsspage"">"&Iif(team.Forum_setting(42)=1,team.BoardJump,"")&"</div>"
	Echo tmp
End Sub

Sub UpdateUserpostExc()
	'�û����ֲ���
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
				UExt = UExt &"Extcredits0=Extcredits0-"&MustSort(6)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"-"&MustSort(6)&""
			End If
			If (team.User_SysTem(14+U)-MustSort(6))-MustSort(8)<0 Then
				team.Error "����"&ExtSort(0)&" ["& team.User_SysTem(14+U) - MustSort(4) &"] ���ڻ��ֲ�������ֵ ["& MustSort(8)&"] �������޷����д˲�����"
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where ID = "& team.TK_UserID)
End Sub

Sub Main
	Echo "<script type=""text/javascript""> "
	Echo "function validate(theform) { "
	Echo " var searchkey = theform.searchkey.value; "
	Echo " if ( searchkey ==  '') { "
    Echo "		alert('�������������'); "
    Echo "		return false; "
	Echo "	}	"
	Echo " this.document.myform.submit.disabled = true; "
	Echo "}"
	Echo "</script>"
	Echo "<form name=""myform"" method=""post"" action=""?action=seachfile"" onSubmit=""return validate(this)"">"
	Echo "<table cellSpacing=""0"" cellPadding=""10"" width=""98%"" border=""0"" align=""center"">"
	Echo "<tr class=""a3""><td width=""30%""></td>"
	Echo "	<td class=""bold""><IMG SRC="""&team.styleurl&"/bin.GIF"" BORDER=""0"" ALT=""����"" align=""absmiddle"">  TEAM��������ϵͳ , ������Ҫ�����Ĺؼ��� </td>"
	Echo "</tr>"
	Echo "<tr class=""a4"">"
	Echo "	<td></td><td>"
	Echo "	<input value="""" name=""searchkey"" size=""50"" onmouseover=""this.focus()"" maxlength=""50"" type=""text"" title=""����������ʼ -- >"" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.select();"" class=""colorblur""> "
	Echo "	<input type=""submit"" name=""submit"" value=""��ʼ����"">"
	Echo "	<BR><BR> "
	Echo "	<input type=""radio"" name=""searchclass"" value=""1"" class=""radio""> ���û������� "
	Echo "	<input type=""radio"" name=""searchclass"" value=""2"" class=""radio"" checked> ���������� "
	If team.Group_Browse(4) = 2 Then  
		Echo "	<input type=""radio"" name=""searchclass"" value=""3"" class=""radio""> ���������� "
	End if
	Echo "	</td>"
	Echo "</tr>"
	Echo "</table>"
	Echo "</form> "
End Sub

Sub CheckpostTime()
	If CID(team.Forum_setting(51))<=0 Then
		Exit Sub
	Else
		If IsDate(Session(CacheName &"searchtime")) Then
			If DateDiff("s",Session(CacheName &"searchtime"),Now())< CID(team.Forum_setting(51)) Then
				team.Error "Ϊ��ֹ���˶�������ϵͳ��Դ����̳���Ƶ����û�������������������"&team.Forum_setting(51)&"�룬������Ҫ�ȴ� "& CLng(team.Forum_setting(51))-DateDiff("s",Session(CacheName &"searchtime"),Now()) &" �롣"
			End If
		End If
	End if
End Sub
%>