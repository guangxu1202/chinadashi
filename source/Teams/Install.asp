<!-- #include file="Conn.asp" -->
<!-- #include file="Inc/Md5.asp" -->
<link href="skins/teams/bbs.css" rel=stylesheet type=text/css>
<%
Dim mconn
Dim MyCode,IsWrite,CodeStr,NextAction
'====================ʹ��ǰ���޸Ĵ˴���==================
'�ϴ�ǰ�޸���֤���룬��ֹ���������á�
MyCode = "team_admin_key"
'�ļ���ʼֵΪ1����Ҫ�޸�Ϊ0�ſ������д��ļ���
IsWrite = 1
'========================================================
NextAction = 0
CodeStr = ""
If IsWrite = 1 Then
	Response.Write "<BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAM��̨����װ���� >>> </td></tr><tr class=a4><td height=100 valign=top  colspan=3> <li><b> ˵��: </b></li> <UL>��װ�ļ�δ���ã����Install.asp�ļ�������IsWrite = 0 ��Ȼ�������ϴ�ԭ�ļ���</UL></td></tr></table><br><center><input type=Submit value='��һ��' name=Submit></center>"
	Response.Write" <BR/><table border=""0"" cellspacing=""3"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4><td  colspan=3>��̳����: TEAM ��̳<BR/>�����д: DayMoon<BR/>��̳��ַ: <a href=http://WWW.TEAM5.CN>WWW.TEAM5.CN</a> </td></tr></table>"
	Response.End 
End if
Select Case Request("Menu")
	Case "Update"
		Update
	Case "Update1"
		Call Update1
	Case "upskins"
		Call upskins
	Case "upclubsys"
		Call upclubsys
	Case Else
		Main
End Select

Sub upclubsys
	ismyKey()
	Conn.execute("update [Clubconfig] set Allclass='0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0*0$$$0$$$0$$$��̳ά����,���Ժ����!$$$1$$$$$$1$$$0$$$TEAM 2.0.4 Release$$$$$$0$$$0.2$$$100$$$0$$$10$$$0$$$��ӭ����TEAM��̳!$$$1$$$1$$$20$$$10$$$20$$$15$$$3$$$0$$$1000$$$1$$$1$$$1$$$1$$$team��̳,asp,bbs,���,����$$$team��̳,asp,bbs,���,����$$$3$$$1$$$30$$$31$$$1$$$1$$$1$$$0$$$1$$$$$$1$$$0$$$60$$$60$$$0$$$3$$$0$$$1$$$10$$$30$$$10$$$500$$$5$$$$$$0$$$$$$$$$��ICP��05004532��$$$5$$$1$$$1$$$1$$$0$$$0$$$TEAM BOARD$$$10000$$$10$$$1$$$0$$$200$$$100$$$rar|jpg|txt|gif|zip$$$999$$$0$$$images/Mypic/logo.gif$$$88$$$1$$$$$$$$$$$$����$$$0$$$$$$1$$$$$$$$$40$$$80$$$0$$$20$$$0$$$0$$$0$$$1$$$1$$$$$$1$$$1$$$51$$$$$$$$$$$$100$$$100$$$0$$$$$$120$$$120$$$0$$$0$$$$$$1$$$$$$1$$$0$$$$$$$$$0$$$$$$'")
	FSOlinewrite "Install.asp",10,"IsWrite = 1"
	Response.Write"<body text=""#000000"" leftmargin=""10"" topmargin=""10""><Br><BR><div class=a3 style=""padding: 15px;width:600""><div class=a4><li>ģ�浼��ɹ������ڽ�ת����̳��ҳ<BR><meta http-equiv=refresh content=3;url=../></div></body>" 
End sub

Sub myConnectionDatabase
	Dim ConnStr
	ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & "skins/TM_Style.mdb")
	On Error Resume Next
	Set mconn = Server.CreateObject("ADODB.Connection")
	mconn.open ConnStr
	If Err Then
		err.Clear
		Set mconn = Nothing
		Response.Write "���ݿ����ӳ������������ִ���"
		Response.End
	End If
End Sub

Sub upskins
	Dim tRs
	ismyKey()
	On Error Resume Next
	If Not IsObject(Conn) Then ConnectionDatabase
	conn.execute("delete from [style]")
	If Not IsObject(mConn) Then myConnectionDatabase
	Dim i
	Dim InsertName,InsertValue
	Set TRs= mconn.Execute("select * from [Style] where id = 1 ")
	Do while not TRs.eof
	InsertName=""
	InsertValue=""
		For i = 1 to TRs.Fields.Count-1
			InsertName=InsertName & TRs(i).Name
			InsertValue=InsertValue & "'" &checkStr(TRs(i)) & "'"
			If i<> TRs.Fields.Count-1 Then 
				InsertName	= InsertName & ","
				InsertValue	= InsertValue & ","
			End If
		Next
	conn.Execute("insert into [Style] ("&InsertName&") values ("&InsertValue&") ")
	TRs.movenext
	loop
	TRs.close
	set TRs=Nothing
	Application.Contents.RemoveAll()
	FSOlinewrite "Install.asp",10,"IsWrite = 1"
	Response.Write"<body text=""#000000"" leftmargin=""10"" topmargin=""10""><Br><BR><div class=a3 style=""padding: 15px;width:600""><div class=a4><li>ģ�浼��ɹ������ڽ�ת����̳��ҳ<BR><meta http-equiv=refresh content=3;url=../></div></body>" 
End Sub

Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
End Function

Sub Update()
	Dim Myname,MyPass,Forumname,ForumPass
	If Request.Form("NextAction") = 1 Then
		Myname=HTMLEncode(Request.Form("Myname"))
		MyPass=HTMLEncode(Request.Form("MyPass"))
		Forumname=HTMLEncode(Request.Form("Forumname"))
		ForumPass=HTMLEncode(Request.Form("ForumPass"))
		Response.Write"<div align=left id=UpFile class=a1><div class=a4>"
		ismyKey()
		ConnectionDatabase
		Conn.Execute("Update [User] Set Userpass ='"&MD5(ForumPass,16)&"',UserGroupID=1 Where  UserName='"&Forumname&"'")
		Conn.Execute("insert into admin (adminname,adminpass,adminclass,forumname) values ('"&Myname&"','"&MD5(MyPass,16)&"','1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36','"&Forumname&"')")
		Response.Write"<li> TEAM���³ɹ� </li>"
		Response.Write"<li> TEAMǰ̨�û����ƣ�"&Forumname&"�����룺"&ForumPass&"</li>"
		Response.Write"<li> TEAM��̨�û����ƣ�"&Myname&"�����룺"&MyPass&"</li>"
		Response.Write"<li> ���³ɹ�,������̳��Ŀ¼ɾ�����ļ�! �ļ��� [Install.asp] </li></div></div>"
		FSOlinewrite "Install.asp",10,"IsWrite = 1"
	Else
		Response.Write "��������ȷ����֤���롣"
		Response.End 
	End if
End Sub

Sub Main
	Response.Write" <BR/><BR/><form name=myform method=post action='?Menu=Update1'><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAM��̨����װ���� >>> </td></tr><tr class=a4><td height=100 valign=top  colspan=3> <li><b> ˵��: </b></li> <UL>1. ���ļ����ڹ������Ǻ�̨��½�������������̨�û�Ȩ��ʱ,���¸��º�̨����ʹ��. </UL> <UL>2. ���޸Ĵ��ļ��Ĳ�����������ϴ�����̳��Ŀ¼!</UL><UL>3. ʹ�����ļ��Ժ��뼰ʱɾ�����ļ�! </UL><UL>4. ����ȷ����Ϊ�˹����õĴ���,���Install.ASP����˵������! </UL></td></tr> <tr class=a3><td colspan=3>���������ȷ����: <input size='15' name='CodeStr'></td></tr></table><br><center><input type=Submit value='��һ��' name=Submit></center>"
	Response.Write" <BR/><table border=""0"" cellspacing=""3"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4><td  colspan=3>��̳����: TEAM ��̳<BR/>�����д: DayMoon<BR/>��̳��ַ: <a href=http://WWW.TEAM5.CN>WWW.TEAM5.CN</a> </td></tr></table>"
End Sub

Sub Update1
	ismyKey()
	ConnectionDatabase
	Response.Write" <BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAMϵͳ�ָ�����ָ�� >>> </td></tr></table>"
	Response.Write" <BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>ģ��ָ�</td></tr>"
	Response.Write" <tr class=a4 align=center><td  colspan=3><li> 1. ����ȷ������̳skinsĿ¼�������TM_Style.MDB�ļ���<li> 2.  �鿴TM_Style.MDB�ļ��ĸ������ڡ����չٷ��ṩ��ѹ�������жԱȡ�<li>����ָ����ܽ���ɾ����̳����ģ�壬����ģ��ָ�Ϊ��̳�ٷ�Ĭ�ϵ�ģ�塣<li>ȷ��������ѡ����� <B><a href=""?Menu=upskins"">ģ��ָ�</a></B></td></tr>"
	Response.Write" </table>"
		
	Response.Write" <BR/><BR/><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>���������ָ�</td></tr>"
	Response.Write" <tr class=a4 align=center><td  colspan=3>�����ѡ�����̳�ָ���ԭʼ�������á� <B><a href=""?Menu=upclubsys"">���������ָ�</a></B></td></tr>"
	Response.Write" </table>"

	Response.Write" <BR/><BR/><form name=myform method=post action='?Menu=Update'><input type=""hidden"" value="""&CodeStr&""" name=""CodeStr""><input type=""hidden"" value=""1"" name=""NextAction""><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a1 align=center><td  colspan=3>TEAM��̨����װ���� >>> </td></tr><tr class=a4><td height=100 valign=top  colspan=3> <li><b> ˵��: </b></li> <UL>1. ���ļ����ڹ������Ǻ�̨��½�������������̨�û�Ȩ��ʱ,���¸��º�̨����ʹ��. </UL> <UL>2. ���޸Ĵ��ļ��Ĳ�����������ϴ�����̳��Ŀ¼!</UL><UL>3. ʹ�����ļ��Ժ��뼰ʱɾ�����ļ�! </UL><UL>4. ����ȷ����Ϊ�˹����õĴ���,���Install.ASP����˵������! </UL><UL>5. �˲�������������ǰ̨�û����룬�����һ���µĺ�̨��½�ʺţ���ע�ⲻҪʹ���ظ��ĺ�̨���룬��Ȼ�������û��޷���½��̨ . �������µĺ�̨��½�ʺ�ʱ����֤��½���Ʋ��ظ��� </UL></td></tr>"
	ConnectionDatabase
	Dim Rs
	Response.Write" <tr class=a1><td colspan=2>  Ŀǰ���ڵĺ�̨�û�  </td></tr><tr class=a4><td colspan=2>"
	Set Rs=Conn.Execute("Select adminname,forumname from [admin]" )
	Do While Not Rs.Eof
		Response.Write " <li>��̨�û����ƣ� "& RS(0)&" -  [ �󶨵�ǰ̨�û����� "& RS(1)&" ]</li> "
		Rs.MoveNext
	Loop
	Rs.close:Set Rs=nothing
	Response.Write" </td> </tr><tr class=a1><td colspan=2>  ��ӵĺ�̨�û�  </td></tr>"
	Response.Write" <tr class=a4><td>ǰ̨�û�����: <input size='15' name='Forumname'> </td><td>ǰ̨�û�����: <input size='15' name='ForumPass'>  </td> </tr>"
	Response.Write"  <tr class=a3><td>��̨�û�����: <input size='15' name='Myname'> </td><td>��̨�û�����: <input size='15' name='MyPass'>  </td></tr>"
	Response.Write" </table><br><center><input type=Submit value='��һ��' name=Submit></center> "
	Response.Write" <BR/><table border=""0"" cellspacing=""3"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4><td  colspan=3>��̳����: TEAM ��̳<BR/>�����д: DayMoon<BR/>��̳��ַ: <a href=http://WWW.TEAM5.CN>WWW.TEAM5.CN</a> </td></tr></table>"
End Sub

Sub ismyKey()
	Dim CodeStr
	CodeStr=Request.Form("CodeStr")
	If Trim(MyCode) <>  Trim(CodeStr) Then
		Response.Write "<BR/><BR/><BR><BR><table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""80%"" align=center class=a2><tr class=a4 align=center><td  colspan=3>��������ȷ����֤���롣 >>> </td></tr></table>"
		Response.End 
	End if
End Sub

Function FSOlinewrite(filename,lineNum,Linecontent)
	if linenum < 1 then exit function
	dim fso,f,temparray,tempCnt
	set fso = server.CreateObject("scripting.filesystemobject")
	if not fso.fileExists(server.mappath(filename)) then exit function
	set f = fso.opentextfile(server.mappath(filename),1)
	if not f.AtEndofStream then
		tempcnt = f.readall
		f.close
		temparray = split(tempcnt,chr(13)&chr(10))
		if lineNum>ubound(temparray)+1 then
			exit function
		else
		temparray(lineNum-1) = lineContent
		end if
		tempcnt = join(temparray,chr(13)&chr(10))
		set f = fso.createtextfile(server.mappath(filename),true)
		f.write tempcnt
	end if
	f.close
	set f = nothing
End Function

Function HTMLEncode(fString)
	If fString="" or IsNull(fString) Then 
		Exit Function
	Else
		If Instr(fString,"'")>0 Then 
			fString = replace(fString, "'","&#39;")
		End If
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32)," ")
		fString = Replace(fString, CHR(9)," ")
		fString = Replace(fString, CHR(34),"&quot;")
		fString = Replace(fString, CHR(13),"")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10),"<BR>")
		fString = Replace(fString, CHR(39),"&#39;")
		HTMLEncode = fString
	End If
End Function
%>
