<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim Admin_Class,Page
Call Master_Us()
Header()
Admin_Class=",10,"
Call Master_Se()

If Cid(Session("UserMember")) <> 1 Then 
	SuccessMsg "�Բ���ֻ�й���Ա���ɲ鿴�˰������ �� "
End if
team.SaveLog ("��̳ά��")
Page = HRF(2,2,"Page")
Select Case Request("action")	
	Case "updates"
		Call updates
	Case "runquery"
		Call runquery
	Case "reforums"
		Call reforums
	Case "updatestb"
		Call updatestb
	Case "creattable"
		Call creattable
	Case "reforumdel"
		Call reforumdel
	Case "upfiles"
		Call upfiles
	Case "attachments"
		Call attachments
	Case "deleattachments"
		Call deleattachments
	Case "BakUserbf"
		Call BakUserbf
	Case "SQLUserReadme"
		Call SQLUserReadme
	Case "rebakuserdata"
		Call rebakuserdata
	Case "compressdata"	
		Call compressdata
	Case "clearmsg"
		Call clearmsg
	Case "delmsgok"
		Call delmsgok
	Case "savelog"
		Call savelog
	Case "dellogok"
		Call dellogok
	Case "reforumdelpass"
		Call reforumdelpass
	Case Else
		Call Main
End Select

Sub dellogok
	Dim lConnStr,lConn,ldb,ho
	ldb = MyDbPath & LogDate
	lConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ldb)
	Set lConn = Server.CreateObject("ADODB.Connection")
	lConn.Open lConnStr
	for each ho in Request.form("deleteid")
		lConn.execute("Delete from [SaveLog] Where ID="&ho)
	Next
	SuccessMsg " ѡ�еĲ�����¼�Ѿ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=savelog>������¼����  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=savelog>�� "
End Sub

Sub savelog %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=dellogok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="6">������¼����</td>
    </tr>
    <tr class="tab3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio"> ɾ</td><td>������Ա</td><td>��½IP</td><td>��������</td><td>����ʱ��</td>
    </tr>
	<%
	Dim Rs,tocou,Maxpage,PageNum,Shows
	Dim SQL,i
	Dim lConnStr,lConn,ldb
	ldb = MyDbPath & LogDate
	lConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ldb)
	Set lConn = Server.CreateObject("ADODB.Connection")
	lConn.Open lConnStr
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	tocou = lConn.execute("Select Count(ID) From [SaveLog]")(0)
	SQL = "Select ID,UserName,IP,Windows,Remark,Logtime From [SaveLog] Order By ID DESC"
	Rs.Open SQL,lConn,1,1,&H0001
	If Rs.Eof And Rs.Bof Then
		Echo "<tr class=""a4""><td colspan=""6"" align=""center""> �������ݲ�����¼ </td></tr></table>"
	Else
		Maxpage = 50
		PageNum = Abs(int(-Abs(tocou/Maxpage)))	'ҳ��
		Page = CheckNum(Page,1,1,1,PageNum)	'��ǰҳ
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If	
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	For i=0 To Ubound(shows,2)
		Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""deleteid"" value="&Shows(0,i)&" class=""radio""></td><td> <a href=""../Profile.asp?username="& Shows(1,i) &""" target=""_blank"" alt=""����鿴"">"& Shows(1,i) &"</a> </td><td>  "& Shows(2,i) &" </td><td align=""left""> "& Shows(4,i) &" </td><td>"& Shows(5,i) &" </td></tr>"
	Next
	Echo "<tr class=""a4""><td colspan=""6"">"
	Echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td>"
	Echo "<script language=""JavaScript"">"
	Echo"		var pg = new showPages('pg');	"
	Echo"		pg.pageCount = "& PageNum &"	;	"
	Echo"		pg.dispCount = "& tocou &";	"
	Echo"		pg.argName = 'Page';"
	Echo"		pg.printHtml(1); "
	Echo "</script></td></tr></table></td></tr></table><BR/><center><input type=""submit"" name=""onlinesubmit"" value=""�� ��""></center></form>"
	Set lConn = Nothing 
End Sub


Sub delmsgok
	Dim ho
	If Request.Form("chkallmsg") = 1 Then
		team.execute("Delete from ["&IsForum&"Message] ")
		SuccessMsg " ���еĶ����Ѿ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=clearmsg>���Ź���  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=clearmsg>�� "
	Else
		for each ho in Request.form("deleteid")
			team.execute("Delete from ["&IsForum&"Message] Where ID="&ho)
		Next
		SuccessMsg " ѡ�еĶ����Ѿ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=clearmsg>���Ź���  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=clearmsg>�� "
	End if
End Sub

Sub clearmsg %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=delmsgok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="5">���Ź��� (ɾ�����ж���<input type="checkbox" name="chkallmsg" class="radio" value="1"> )</td>
    </tr>
    <tr class="tab3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)" class="radio"> ɾ </td><td>������</td><td>������</td><td>����</td><td>����ʱ��</td>
    </tr>
	<%
	Dim Rs,tocou,Maxpage,PageNum,Shows
	Dim SQL,i
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	tocou = team.execute("Select Count(ID) From ["&IsForum&"Message]")(0)
	SQL = "Select ID,author,incept,msgtopic,Sendtime,isbak From ["&IsForum&"Message] Order By Sendtime asc"
	Rs.Open SQL,Conn,1,1,&H0001
	If Rs.Eof And Rs.Bof Then
		Echo "<tr class=""a4""><td colspan=""5"" align=""center""> �������������� </td></tr></table>"
	Else
		Maxpage = 20
		PageNum = Abs(int(-Abs(tocou/Maxpage)))	'ҳ��
		Page = CheckNum(Page,1,1,1,PageNum)	'��ǰҳ
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If	
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	For i=0 To Ubound(shows,2)
		Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""deleteid"" value="&Shows(0,i)&" class=""radio""></td><td> <a href=""../Profile.asp?username="& Shows(1,i) &""" target=""_blank"" alt=""����鿴"">"& Shows(1,i) &"</a> </td><td> <a href=""../Profile.asp?username="& Shows(2,i) &""" target=""_blank"">"& Shows(2,i) &"</a> </td><td align=""left""> <a href=""../Msg.asp?action=readmsg&sid="& Shows(0,i) &""" target=""_blank"">"& Shows(3,i) &"</a> "
		If Shows(5,i) = 1 Then
			Echo " - [�ݸ�]"
		End if
		Echo "</td><td>"& Shows(4,i) &" </td></tr>"
	Next
	Echo "<tr class=""a4""><td colspan=""5"">"
	Echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td>"
	Echo "<script language=""JavaScript"">"
	Echo"		var pg = new showPages('pg');	"
	Echo"		pg.pageCount = "& PageNum &"	;	"
	Echo"		pg.dispCount = "& tocou &";	"
	Echo"		pg.argName = 'Page';"
	Echo"		pg.printHtml(1); "
	Echo "</script></td></tr></table></td></tr></table><BR/><center><input type=""submit"" name=""onlinesubmit"" value=""�� ��""></center></form>"	
End Sub

Sub deleattachments
	Dim ho,mFso,fPath,Rs,fName
	fPath = "../Images/Upfile/"
	Set mFso = Server.CreateOBject("Scripting.FileSystemObject")
	for each ho in Request.form("deleteid")
		Set Rs = team.execute("Select FileName From ["&IsForum&"Upfile] Where FILEID="&ho)
		If Not Rs.Eof Then
			fName = fPath & Rs(0)
			If  mFso.FileExists(Server.mappath(fName)) Then
				'On Error Resume Next
				mFso.deletefile(Server.mappath(fName))
			End  If
		End if
		team.execute("Delete from ["&IsForum&"Upfile] Where FILEID="&ho)
	Next
	SuccessMsg " ѡ�еĸ����Ѿ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=upfiles>��������  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=upfiles>�� "
End Sub

Sub attachments
	Dim inforum,dmincount,dmaxcount,upname,upsize
	Dim Twher,tocou,sql,Maxpage,PageNum,Rs,Shows
	Dim i,tids
	inforum = HRF(1,2,"inforum")
	tids = HRF(1,2,"tids")
	upname = HRF(1,1,"upname")
	upsize = HRF(1,1,"upsize")
	dmaxcount = HRF(1,2,"dmaxcount")
	dmincount = HRF(1,2,"dmincount")
	If upname&"" = "" Then
		Twher = " UserName <>'' "
	Else
		Twher = " UserName Like '% "& upname &" %' "
	End if
	If upsize <> "" Then
		Twher = Twher & " and FileName Like '% "& upsize &" %'"
	End if
	If dmaxcount > 0 Then
		Twher = Twher & " and Upcount>"& dmaxcount &" "
	End if
	If dmincount > 0 Then
		Twher = Twher & " and Upcount<"& dmincount &" "
	End If
	If inforum > 0 Then
		Twher = Twher & " and FID="& Int(inforum) &" "
	End If
	If tids > 0 Then
		Twher = Twher & " and ID="& Int(tids) &" "
	End if
	tocou = team.execute("Select Count(ID) From ["&IsForum&"Upfile] Where "&Twher&" ")(0)
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	SQL = "Select FILEID,ID,FID,UserName,FileName,Types,FileSize,Upcount,ByPowers,Lasttime From ["&IsForum&"Upfile] Where "&Twher&" Order By Lasttime Desc"
	Rs.Open SQL,Conn,1,1,&H0001
	Response.Write "<body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10""><br><br><form method=""post"" action=""?action=deleattachments""><table cellspacing=""1"" cellpadding=""5"" width=""95%"" align=""center"" border=""0"" class=""a2""><tr class=""a3""><td colspan=""8"" align=""center"">�����������ҵ� <Font color=""red"">"& tocou &"</Font> ����ظ�����¼</td></tr><tr class=""tab1""><td><input type=""checkbox"" name=""chkall"" onClick=""checkall(this.form)"" class=""radio""> ɾ</td><td> ��������</td><td>��������</td><td>�ϴ��û�</td><td>�ϴ�ʱ��</td><td>�Ķ�Ȩ��</td><td>���ش���</td><td>����״̬</td></tr>"
	If Rs.Eof And Rs.Bof Then
		Echo "<tr class=""a4""><td colspan=""8"" align=""center""> �Բ���û���ҵ���Ҫ��ѯ������ </td></tr></table>"
	Else
		Maxpage = 20
		PageNum = Abs(int(-Abs(tocou/Maxpage)))	'ҳ��
		Page = CheckNum(Page,1,1,1,PageNum)	'��ǰҳ
		Rs.AbsolutePosition=(Page-1)*Maxpage+1
		Shows = Rs.GetRows(Maxpage)
		Rs.Close:Set Rs=Nothing
	End If
	If Not IsArray(Shows) Then
		Exit Sub
	End If
	For i=0 To Ubound(shows,2)
		Echo "<tr class=""tab4""><td><input type=""checkbox"" name=""deleteid"" value="&Shows(0,i)&" class=""radio""></td><td> <a href=""../Images/Upfile/"& Shows(4,i) &""" target=""_blank"" alt=""����鿴"">"& Shows(4,i) &"</a> </td><td> <a href=""../Thread.asp?tid="& Shows(1,i) &""" target=""_blank"">��������</a> </td><td>"& Shows(3,i) &"</td><td>"& Shows(9,i) &" </td><td>"& Shows(8,i) &"</td><td>"& Shows(7,i) &"</td><td>"
		Set Rs = team.execute("Select Deltopic from ["&IsForum&"Forum] Where ID="& Cid(Shows(1,i)))
		If Rs.Eof And Rs.Bof Then
			Echo "��ɾ��"
		Else
			If CID(Rs(0)) = 1 Then
				Echo "��ɾ��"
			Else
				Echo "����"
			End If
		End if
		Echo "</td></tr>"
	Next
	Echo "<tr class=""a4""><td colspan=""8"">"
	Echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td>"
	Echo "<script language=""JavaScript"">"
	Echo"		var pg = new showPages('pg');	"
	Echo"		pg.pageCount = "& PageNum &"	;	"
	Echo"		pg.dispCount = "& tocou &";	"
	Echo"		pg.argName = 'inforum="&inforum&"&upname="&upname&"&upsize="&upsize&"&dmaxcount="&dmaxcount&"&dmincount="&dmincount&"&Page';"
	Echo"		pg.printHtml(1); "
	Echo "</script></td></tr></table></td></tr></table><BR/><center><input type=""submit"" name=""onlinesubmit"" value=""�� ��""></center></form>"
End Sub


Sub upfiles	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<form method="post" action="?action=attachments">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr>
      <td class="a1" colspan="2">�������� ��ģ��������</td>
    </tr>
    <tr>
      <td class="altbg1">������̳:</td>
      <td class="altbg2" align="right">
		<select name="inforum">
			<option value="all"	selected="selected">&nbsp;&nbsp;> ȫ��</option>
			<%Call BBsList(0)%>
        </select>
      </td>
    </tr>
    <tr>
      <td class="altbg1">����ID:</td>
      <td class="altbg2" align="right"><input type="text" name="tids" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">�ϴ��û���:</td>
      <td class="altbg2" align="right"><input type="text" name="upname" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">��������:</td>
      <td class="altbg2" align="right"><input type="text" name="upsize" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">�����ش�������:</td>
      <td class="altbg2" align="right"><input type="text" name="dmaxcount" size="40"></td>
    </tr>
    <tr>
      <td class="altbg1">�����ش���С��:</td>
      <td class="altbg2" align="right"><input type="text" name="dmincount" size="40"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="searchsubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub reforumdel
	Dim Rs,tablename
	Tablename = Replace(Request("tablename"),"'","''")
	If Tablename&"" = "" Then
		Successmsg " ��������������ơ�"
	Else
		If not team.Execute("Select ReList From ["&isforum&"forum] where ReList='"&Tablename&"'" ).eof Then
			SuccessMsg " �ñ����ж�Ӧ�����⣬��ȷ��Ҫɾ��ô�� <a href=""?action=reforumdelpass&pname="& Tablename &""">���ȷ���밴��һ��</a>"
		Else
			if Ucase(Trim(team.Club_Class(11))) = Ucase(Trim(Tablename)) then
				SuccessMsg("��ǰ����ʹ���е����ݿⲻ��ɾ����")
			End If
			team.execute " delete from ["&isforum&"TableList] where TableName='"&Tablename&"' "
			team.Execute " drop table "&Tablename&"  " 
			SuccessMsg " ѡ�еĻ������Ѿ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=reforums>����������  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>�� "
		End if
	End if
End Sub

Sub reforumdelpass
	Dim pname
	pname = Replace(Request("pname"),"'","''")
	if Ucase(Trim(team.Club_Class(11))) = Ucase(Trim(pname)) then
		SuccessMsg("��ǰ����ʹ���е����ݿⲻ��ɾ����")
	End If
	team.execute " delete from ["&isforum&"TableList] where TableName='"&pname&"' "
	team.Execute " drop table "& pname &"  " 
	SuccessMsg " ѡ�еĻ������Ѿ���ɾ������ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=reforums>����������  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>�� "
End Sub 

Sub creattable
	Dim SQL,tablename
	Tablename = Replace(Request.Form("tablename"),"'","''")
	If Tablename&"" = "" Then
		Successmsg " ��������������ơ�"
	Else
		Sql="CREATE TABLE "&isforum&""&tablename&" ("&_
			"id int IDENTITY (1, 1) NOT NULL ,"&_
			"topicid int NOT NULL ,"&_
			"username varchar(255) NOT NULL ,"&_
			"ReTopic varchar(255) NOT NULL ,"&_
			"content text NOT NULL ,"&_
			"posttime datetime Default "&SqlNowString&" NOT NULL ,"&_
			"postip varchar(255)  NOT NULL ,"&_
			"Reward int NOT NULL ,"&_
			"IsNoName int NOT NULL ,"&_
			"Auditing int NOT NULL ,"&_
			"lock int NULL"&_
			")"
		team.execute(sql)
		team.Execute("insert into ["&isforum&"TableList] (TableName) values ('"&tablename&"')" )
	End if
	SuccessMsg "�»��������ɹ�����ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=reforums>����������  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>�� "
End Sub

Sub updatestb
	Dim tablename
	Tablename = Replace(Request.Form("tablename"),"'","''")
	If Tablename&"" = "" Then
		Successmsg " ��������������ơ�"
	Else
		Cache.DelCache("club_class")
		team.execute("update ["&isforum&"Clubconfig] set ReForumName='"&Tablename&"'")
		Successmsg " ���������óɹ� ����ȴ�ϵͳ�Զ����ص� <a href=Admin_dbmake.asp?action=reforums>����������  </a> ҳ�� ��<meta http-equiv=refresh content=3;url=Admin_dbmake.asp?action=reforums>�� "
	End if
End Sub

Sub reforums
	If IsSqlDataBase = 1 then
		Successmsg " <BR><BR><BR><div class=""a2"" style='height:50;width:80%'> <ul><BR><li>SQL�汾�������û�����</li></ul></div>"
		Exit Sub
	End If
	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="altbg1">
    <td><br>
      <ul>
        <li>�����������ݴ�������ʱ���ᵼ�¶�ȡ���ݱ������������һ���µĻ�����������Ч�ӿ��ٶȡ�</li>
      </ul>
      <ul>
        <li> ʹ��ACCSEE���ݿ�ʱ�������ݿ����������100M�Ժ�����㷢�־�����Ӹ���Ļ�����Ҳ���������ı��ٶȣ���ô�Ƽ�������SQL���ݿ⡣</li>
      </ul></td>
  </tr>
</table>
<BR>
<form method="post" action="?action=updatestb">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan="4">�������ݱ���� </td>
    </tr>
    <tr class="a3"  align="center">
      <td> ��ǰ������ </td>
      <td> ������ </td>
      <td> ѡ�� </td>
      <td> ���� </td>
    </tr>
    <%
			Dim Rs
			Set Rs=team.execute("select TableName from ["& isforum &"TableList] ")
			Do While Not RS.EOF
				Echo " <tr class=""a4"" align=""center"">"
				Echo " <td bgcolor=""#FFFFFF""> "&RS(0)&"</td>"
				Echo " <td bgcolor=""#F8F8F8""> "&team.execute("Select count(id)from ["&RS(0)&"]")(0)&" </td>"
				Echo " <td bgcolor=""#FFFFFF""> <input type=""radio"" "
				if Ucase(Trim(team.Club_Class(11))) = Ucase(Trim(Rs(0))) then 
					Echo " CHECKED "
				End if
				Echo " value="&RS(0)&" name=""tablename""> </td><td bgcolor=""#F8F8F8""> "
				If Ucase(Trim(Rs(0)))=Ucase("Reforum") Then
					Echo "Ĭ�ϱ���ɾ��"
				Else
					Echo " <a href=""?action=reforumdel&tablename="&RS(0)&""">ɾ��</a> "
				End if
				Echo " </td></tr>"
			RS.MoveNext
		Loop
		Rs.close:Set Rs = Nothing
		%>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="�� ��">
  </center>
</form>
<form method="post" action="?action=creattable">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr class="a4">
      <td class="a1" colspan="2"> ����µĻ����� </td>
    </tr>
    <tr class="a4">
      <td width="60%"><B>����µ����ݱ�</B><br>
        ��д���µĻ��������ƣ�����ӵĻ��������Ʋ������Ѿ����ڵĻ�����������ͬ��������������Ƽ�ʹ��Ӣ����ĸ��</td>
      <td width="40%"><input type="text" size="30" name="tablename" value="newreforum"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="�� ��">
  </center>
</form>
<%
End Sub


Sub runquery
	Dim Sqlstr,Sqlstr1,i
	Sqlstr=Request.Form("queries")
	If Sqlstr="" Then
		Successmsg("������sqlִ�����!")
		Exit Sub
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	On Error Resume Next
	If InStr(Sqlstr,Chr(13)&Chr(10))>0 Then
		Sqlstr1 = Split(Sqlstr,Chr(13)&Chr(10))
		For i=0 To UBound(Sqlstr1)
			Conn.Execute(Sqlstr1(i))
		Next
	Else
		Conn.Execute(Sqlstr)
	End if
	If Err Then
		Err.Clear
		Successmsg "�������sql����д��� �� <blockquote> "&Sqlstr&" </blockquote>"
	Else
		Successmsg " �ɹ�ִ��SQL��� ��"
	End If
End Sub

Sub updates %>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<table cellspacing="1" cellpadding="4" width="60%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="altbg1">
    <td><br>
      <ul>
        <li> �������ݿ������һ����Σ���ԣ���С�Ĳ�����</li>
		<li> ÿ������һ��SQL��䣬����һ�����������SQL����������</li>
      </ul>
	  </td>
  </tr>
</table>
<BR>
<form method="post" action="?action=runquery">
  <table cellspacing="1" cellpadding="4" width="60%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">TEAM's ���ݿ����� - �뽫���ݿ��������ճ�������� </td>
    </tr>
    <tr class="altbg1" align="center">
      <td valign="top"><textarea cols="85" rows="10" name="queries"></textarea>
        <br>
        <br>
        ע��: Ϊȷ�������ɹ����벻Ҫ�޸� SQL �����κβ��֡�</td>
    </tr>
  </table>
  <br>
  <br>
  <center>
    <input type="submit" name="sqlsubmit" value="�� ��">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub Main	
	If IsSqlDataBase = 1 then
		Call SQLUserReadme()
		Exit Sub
	End If
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="altbg1">
    <td><br>
      <ul>
        <li> ���²�����Ҫ�ռ��FSO�����֧�֣���鿴<a href="Admin_Path.asp?action=discreteness"> <B>���֧�����</B> </a>ȷ�ϡ�</li>
      </ul>
      <ul>
        <li> �����������ݿ�Ĳ���ǰ�����ȹر���̳��</li>
      </ul>
      <ul>
        <li> �����Եı������ݿ������Ч��ֹ�����ݿ��𻵴�����Ӱ�죨����ÿ�����ڱ���һ�Σ����������ݿ�ʱ�����޸�Ĭ�ϵı���·���ͱ����ļ����ƣ����������Ĭ�����ݿ����ƣ����������ݿⱻ�ڿ����أ��Ӷ�����������ƽ��Σ�ա�</li>
      </ul>
      <ul>
        <li> �����ݿ������ԵĽ���ѹ������Ч�ӿ���̳�������ٶȣ�����ÿ����ѹ��һ�Σ� ����ȷ��ѹ������Ӧ�����Ƚ����ݿⱸ�ݣ�Ȼ��Ա��ݺõ����ݿ����ѹ����ѹ����ɺ��ٽ�ѹ�����ݿ⻹ԭΪ��ǰ���ݿ⡣���𽫵�ǰ�����ݿ����ѹ������Ϊ���������������ݿ��Σ�ա� </li>
      </ul></td>
  </tr>
</table>
<BR>
<table cellspacing="1" cellpadding="3" width="95%" align="center">
  <tr>
    <td class="a2"><BR>
      <ul>
        <li> <FONt  COLOR="red">���²��������ݿ�Ǳ��Σ���ԣ�����ʧ��������ݿ���𻵣���������������Ӧ�ļ��ɺ��ٶ����ݿ�������á�</FONt></li>
      </ul></td>
  </tr>
</table>
<BR>
<form method="post" action="?action=BakUserbf">
  <table cellspacing="1" cellpadding="3" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan="2">�������ݿ� ( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )</td>
    </tr>
    <tr class="a3">
      <td width="30%">��ǰ���ݿ�·��(���·��)�� </td>
      <td width="70%"><input type="text" size="60" name="DBpath" value="../<%=db%>"></td>
    </tr>
    <tr class="a4">
      <td width="30%">�������ݿ�Ŀ¼(���·��)��<br>
        ��Ŀ¼�����ڣ������Զ�����</td>
      <td width="70%"><input type="text" size="60" name="bkfolder" value="../Databackup"></td>
    </tr>
    <tr class="a4">
      <td width="30%">�������ݿ�����(��д����)��<br>
        �籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����</td>
      <td width="70%"><input type="text" size="60" name="bkDBname" value="teams_Backup.mdb"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="�� ��">
  </center>
</form>
<br>
<form method="post" action="?action=rebakuserdata">
  <table cellspacing="1" cellpadding="2" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan=2>�ָ����ݿ� ( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )</td>
    </tr>
    <tr class="a3">
      <td width="30%">�������ݿ�·��(���)�� </td>
      <td width="70%"><input type=text size="60" name="DBpath" value="../DataBackup/teams_Backup.MDB"></td>
    </tr>
    <tr class="a4">
      <td width="30%">Ŀ�����ݿ�·��(���)��</td>
      <td width="70%"><input type=text size="60" name="backpath" value="../<%=db%>"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="�� ��">
  </center>
</form>
<BR>
<form action="?action=compressdata" method="post">
  <table cellspacing="1" cellpadding="2" width="95%" border="0" class="a2" align="center">
    <tr>
      <td class="a1" colspan="2">ѹ�����ݿ�</td>
    </tr>
    <tr>
      <td class="a4" colspan="2"><b>ע�⣺</b><br>
        �������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ��������</td>
    </tr>
    <tr class="a3">
      <td width="30%">���ݿ�·���� </td>
      <td width="70%"><input size="60" value="../DataBackup/teams_Backup.MDB" name="dbpath"></td>
    </tr>
    <tr class="a4">
      <td width="30%">���ݿ��ʽ��</td>
      <td width="70%"><input type="radio" value="true" name="boolIs97" id="boolIs97">
        <label for="boolIs97">Access 97</label>
        <input type="radio" value="" name="boolIs97" checked id="boolIs97_1">
        <label for="boolIs97_1">Access 2000��2002��2003</label></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="exportsubmit" value="ѹ ��">
  </center>
</form>
<BR>
<%
End Sub

Sub compressdata()
	dim dbpath,boolIs97
	dbpath = request("dbpath")
	boolIs97 = request("boolIs97")
	If dbpath <> "" then
		dbpath = server.mappath(dbpath)
		response.write(CompactDB(dbpath,boolIs97))
	End If
End Sub

Sub BakUserbf
		Dim Dbpath,backpath,testConn,bkfolder,bkdbname,fso
		On error resume next
		Dim FileConnStr,Fileconn
		Dbpath=request.Form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.Form("bkfolder")
		bkdbname=request.Form("bkdbname")
		FileConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Dbpath
		Set Fileconn = Server.CreateObject("ADODB.Connection")
		Fileconn.open FileConnStr
		If Err then
			Response.Write Err.Description
			Err.Clear
			Set Fileconn = Nothing
			SuccessMsg("���ݵ��ļ����ǺϷ������ݿ�!")
			Exit Sub
		Else
			Set Fileconn = Nothing
		End If
		Set Fso=server.createobject("scripting.filesystemobject")
		If Fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = true then
				Fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
				MakeNewsDir bkfolder
				Fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			SuccessMsg("�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname &" ")
		Else
			SuccessMsg("�Ҳ���������Ҫ���ݵ��ļ�!")
		End if
End Sub
Sub rebakuserdata
	Dim Dbpath,backpath,testConn,fso
	Dbpath=request.form("Dbpath")
	backpath=request.form("backpath")
	if dbpath="" then
		SuccessMsg("��������Ҫ�ָ��ɵ����ݿ�ȫ��!")
	else
		Dbpath=server.mappath(Dbpath)
	end if
	backpath=server.mappath(backpath)
	Set testConn = Server.CreateObject("ADODB.Connection")
	On Error Resume Next
	testConn.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Dbpath
	If Err then
		Response.Write Err.Description
		Err.Clear
		Set testConn = Nothing
		SuccessMsg("���ݵ��ļ����ǺϷ������ݿ�!")
		Response.End 
	Else
		Set testConn = Nothing
	End If
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(dbpath) then  					
		fso.copyfile Dbpath,Backpath
		SuccessMsg("���ݿ�ָ��ɹ�!")
	else
		SuccessMsg("����Ŀ¼�²������ı����ļ�!")
	end if
End Sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	Dim fso1
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = true
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	dim f,fso1
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = true
    Set fso1 = nothing
End Function
'=====================ѹ������=========================
Function CompactDB(dbPath, boolIs97)
	Dim fso, Engine, strDBPath,JEt_3X
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(dbPath) then
		fso.CopyFile dbpath,strDBPath & "temp.mdb"
		Set Engine = CreateObject("JRO.JetEngine")
		If boolIs97 = "true" then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
			& "Jet OLEDB:Engine type=" & JEt_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		End If
		fso.CopyFile strDBPath & "temp1.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		fso.DeleteFile(strDBPath & "temp1.mdb")
		Set fso = nothing
		Set Engine = nothing
		SuccessMsg("������ݿ�, " & dbpath & ", �Ѿ�ѹ���ɹ�!") & vbCrLf
	Else
		SuccessMsg("���ݿ����ƻ�·������ȷ. ������!") & vbCrLf
	End If
End Function

Sub SQLUserReadme()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align="center" width="95%" class="a2">
  <tr>
    <td class="a1">&nbsp;&nbsp;SQL���ݿ����ݴ���˵��</td>
  </tr>
  <tr>
    <td class="a4"><blockquote> <B>һ���������ݿ�</B> <BR>
        <BR>
        1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
        2��SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼<BR>
        3��ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ�񱸷����ݿ�<BR>
        4������ѡ��ѡ����ȫ���ݣ�Ŀ���еı��ݵ����ԭ����·����������ѡ�����Ƶ�ɾ����Ȼ�����ӣ����ԭ��û��·����������ֱ��ѡ����ӣ�����ָ��·�����ļ�����ָ�����ȷ�����ر��ݴ��ڣ����ŵ�ȷ�����б��� <BR>
        <BR>
        <B>������ԭ���ݿ�</B><BR>
        <BR>
        1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
        2��SQL Server��-->˫������ķ�����-->��ͼ�������½����ݿ�ͼ�꣬�½����ݿ����������ȡ<BR>
        3������½��õ����ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ��ָ����ݿ�<BR>
        4���ڵ������Ĵ����еĻ�ԭѡ����ѡ����豸-->��ѡ���豸-->�����-->Ȼ��ѡ����ı����ļ���-->��Ӻ��ȷ�����أ���ʱ���豸��Ӧ�ó������ղ�ѡ������ݿⱸ���ļ��������ݺ�Ĭ��Ϊ1���������ͬһ���ļ�������α��ݣ����Ե�����ݺ��ԱߵĲ鿴���ݣ��ڸ�ѡ����ѡ�����µ�һ�α��ݺ��ȷ����-->Ȼ�����Ϸ������Աߵ�ѡ�ť<BR>
        5���ڳ��ֵĴ�����ѡ�����������ݿ���ǿ�ƻ�ԭ���Լ��ڻָ����״̬��ѡ��ʹ���ݿ���Լ������е��޷���ԭ����������־��ѡ��ڴ��ڵ��м䲿λ�Ľ����ݿ��ļ���ԭΪ����Ҫ������SQL�İ�װ�������ã�Ҳ����ָ���Լ���Ŀ¼�����߼��ļ�������Ҫ�Ķ������������ļ���Ҫ���������ָ��Ļ���������Ķ���������SQL���ݿ�װ��D:\Program Files\Microsoft SQL Server\MSSQL\Data����ô�Ͱ������ָ�������Ŀ¼������ظĶ��Ķ������������ļ�����øĳ�����ǰ�����ݿ�������ԭ����bbs_data.mdf�����ڵ����ݿ���forum���͸ĳ�forum_data.mdf������־�������ļ���Ҫ���������ķ�ʽ����صĸĶ�����־���ļ�����*_log.ldf��β�ģ�������Ļָ�Ŀ¼�������������ã�ǰ���Ǹ�Ŀ¼������ڣ���������ָ��d:\sqldata\bbs_data.mdf����d:\sqldata\bbs_log.ldf��������ָ�������<BR>
        6���޸���ɺ󣬵�������ȷ�����лָ�����ʱ�����һ������������ʾ�ָ��Ľ��ȣ��ָ���ɺ�ϵͳ���Զ���ʾ�ɹ������м���ʾ�������¼����صĴ������ݲ�ѯ�ʶ�SQL�����Ƚ���Ϥ����Ա��һ��Ĵ����޷���Ŀ¼��������ļ����ظ������ļ���������߿ռ䲻���������ݿ�����ʹ���еĴ������ݿ�����ʹ�õĴ��������Գ��Թر����й���SQL����Ȼ�����´򿪽��лָ��������������ʾ����ʹ�õĴ�����Խ�SQL����ֹͣȻ�����𿴿����������������Ĵ���һ�㶼�ܰ��մ�����������Ӧ�Ķ��󼴿ɻָ�<BR>
        <BR>
        <B>�����������ݿ�</B><BR>
        <BR>
        һ������£�SQL���ݿ�����������ܴܺ�̶��ϼ�С���ݿ��С������Ҫ������������־��С��Ӧ�����ڽ��д˲����������ݿ���־����<BR>
        1���������ݿ�ģʽΪ��ģʽ����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼-->ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�����Ҽ�ѡ������-->ѡ��ѡ��-->�ڹ��ϻ�ԭ��ģʽ��ѡ�񡰼򵥡���Ȼ��ȷ������<BR>
        2���ڵ�ǰ���ݿ��ϵ��Ҽ��������������е��������ݿ⣬һ�������Ĭ�����ò��õ�����ֱ�ӵ�ȷ��<BR>
        3��<font color="blue">�������ݿ���ɺ󣬽��齫�������ݿ�������������Ϊ��׼ģʽ����������ͬ��һ�㣬��Ϊ��־��һЩ�쳣����������ǻָ����ݿ����Ҫ����</font> <BR>
        <BR>
        <B>�ġ��趨ÿ���Զ��������ݿ�</B><BR>
        <BR>
        <font color="red">ǿ�ҽ������������û����д˲�����</font><BR>
        1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����<BR>
        2��Ȼ�������˵��еĹ���-->ѡ�����ݿ�ά���ƻ���<BR>
        3����һ��ѡ��Ҫ�����Զ����ݵ�����-->��һ�����������Ż���Ϣ������һ�㲻����ѡ��-->��һ��������������ԣ�Ҳһ�㲻ѡ��<BR>
        4����һ��ָ�����ݿ�ά���ƻ���Ĭ�ϵ���1�ܱ���һ�Σ��������ѡ��ÿ�챸�ݺ��ȷ��<BR>
        5����һ��ָ�����ݵĴ���Ŀ¼��ѡ��ָ��Ŀ¼������������D���½�һ��Ŀ¼�磺d:\databak��Ȼ��������ѡ��ʹ�ô�Ŀ¼������������ݿ�Ƚ϶����ѡ��Ϊÿ�����ݿ⽨����Ŀ¼��Ȼ��ѡ��ɾ�����ڶ�����ǰ�ı��ݣ�һ���趨4��7�죬�⿴���ľ��屸��Ҫ�󣬱����ļ���չ��һ�㶼��bak����Ĭ�ϵ�<BR>
        6����һ��ָ��������־���ݼƻ�����������Ҫ��ѡ��-->��һ��Ҫ���ɵı���һ�㲻��ѡ��-->��һ��ά���ƻ���ʷ��¼�������Ĭ�ϵ�ѡ��-->��һ�����<BR>
        7����ɺ�ϵͳ�ܿ��ܻ���ʾSql Server Agent����δ�������ȵ�ȷ����ɼƻ��趨��Ȼ���ҵ��������ұ�״̬���е�SQL��ɫͼ�꣬˫���㿪���ڷ�����ѡ��Sql Server Agent��Ȼ�������м�ͷ��ѡ���·��ĵ�����OSʱ�Զ���������<BR>
        8�����ʱ�����ݿ�ƻ��Ѿ��ɹ��������ˣ�������������������ý����Զ����� <BR>
        <BR>
        �޸ļƻ���<BR>
        1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->����-->���ݿ�ά���ƻ�-->�򿪺�ɿ������趨�ļƻ������Խ����޸Ļ���ɾ������ <BR>
        <BR>
        <B>�塢���ݵ�ת�ƣ��½����ݿ��ת�Ʒ�������</B><BR>
        <BR>
        һ������£����ʹ�ñ��ݺͻ�ԭ����������ת�����ݣ�����������£������õ��뵼���ķ�ʽ����ת�ƣ�������ܵľ��ǵ��뵼����ʽ�����뵼����ʽת������һ�����þ��ǿ������������ݿ���Ч�������������С�����������ݿ�Ĵ�С��������Ĭ��Ϊ����SQL�Ĳ�����һ�����˽⣬��������еĲ��ֲ�������⣬������ѯtEAM��̳�����Ա���߲�ѯ��������<BR>
        1����ԭ���ݿ�����б��洢���̵�����һ��SQL�ļ���������ʱ��ע����ѡ����ѡ���д�����ű��ͱ�д�����������Ĭ��ֵ�ͼ��Լ���ű�ѡ��<BR>
        2���½����ݿ⣬���½����ݿ�ִ�е�һ������������SQL�ļ�<BR>
        3����SQL�ĵ��뵼����ʽ���������ݿ⵼��ԭ���ݿ��е����б�����<BR>
      </blockquote></td>
  </tr>
</table>
<%
end sub

Sub BBsList(V)
	Dim SQL,ii,RS,i
	Set Rs=Team.Execute("Select ID,BBSname,Followid From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		If RS(2)=0 Then 
			Echo "<optgroup label="""&Rs(1)&""">"
		Else
			Echo "<option value="&RS(0)&">"&String(ii,"��") & RS(1)&"</option>"
		End if
		ii=ii+1
		BBsList RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub

Footer()
%>
