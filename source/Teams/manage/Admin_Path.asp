<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Public boards
Dim Admin_Class,Uid
Call Master_Us()
Uid = Cid(Request("uid"))
Header()
Admin_Class=",11,"
Call Master_Se()
team.SaveLog ("  ͳ����Ϣ")
Select Case Request("action")
	Case "discreteness"
		Call discreteness
	Case "statroom"
		Call statroom
	Case Else
		Call Main
End Select
Sub statroom	
	Dim fso,upfacedir,d,upphotosize,tolsize,totalBytes,upfacesize,upphotodir
	dim upfiledir,upfilesize,toldir
	set fso=server.createobject("scripting.filesystemobject")
	upfacedir=server.mappath("../images/upface")
	set d=fso.getfolder(upfacedir)
	upfacesize=d.size
	upfiledir=server.mappath("../images/upfile")
	set d=fso.getfolder(upfiledir)
	upfilesize=d.size
	toldir=server.mappath("../")
	set d=fso.getfolder(toldir)
	tolsize=d.size
	totalBytes=upfacesize+upphotosize+upfilesize+tolsize
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="10" width="95%" align="center" class="a2">
  <tr class="a1">
    <td colspan="2">ͳ��ռ�ÿռ� [�� <%totalsize(totalBytes)%>]</td>
  </tr>
  <tr class="a3">
	<td width="20%"> �ϴ�ͷ��ռ�ÿռ� </td>
	<td>
		<div class="a2" Style="width:<%=Int(upfacesize/totalBytes*100)%>%;padding:3px" title="<%=Int(upfacesize/totalBytes*100)%>%"> <%totalsize(upfacesize)%> </div>
	 </td>
  </tr>
  <tr class="a4">
	<td> �ϴ�����ռ�ÿռ� </td>
	<td>
		<div class="a2" Style="width:<%=Int(upfilesize/totalBytes*100)%>%;padding:3px" title="<%=Int(upfilesize/totalBytes*100)%>%"> <%totalsize(upfilesize)%> </div>
	 </td>
  </tr>
  <tr class="a3">
	<td> Ŀ¼�ܹ�ռ�ÿռ� </td>
	<td>
		<div class="a2" Style="width:<%=Int(tolsize/totalBytes*100)%>%;padding:3px" title="<%=Int(tolsize/totalBytes*100)%>%"> <%totalsize(tolsize)%> </div>
	 </td>
  </tr>
</table>
<%
End Sub

Sub totalsize(size)
	if size<1024 then
		response.write size&" Bytes"
	elseif size<1048576 then
		response.write Int(size/1024)&" KB"
	else
		response.write Int(size/1024/1024)&" MB"
	end if
End sub
Sub discreteness	

	Echo " <body Style=""background-color:#8C8C8C"" text=""#000000"" leftmargin=""10"" topmargin=""10""> "
	Echo " <br>"
	Echo " <table cellspacing=""1"" cellpadding=""4"" width=""95%"" align=""center"" class=""a2""> "
	Echo "	<tr class=""a1"">"
    Echo "		<td width=""50%"">�������</td><td>����</td> "
	Echo "	</tr>"
	Dim theInstalledObjects(17)
	theInstalledObjects(0) = "MSWC.AdRotator"
	theInstalledObjects(1) = "MSWC.BrowserType"
	theInstalledObjects(2) = "MSWC.NextLink"
	theInstalledObjects(3) = "MSWC.Tools"
	theInstalledObjects(4) = "MSWC.Status"
	theInstalledObjects(5) = "MSWC.Counters"
	theInstalledObjects(6) = "MSWC.PermissionChecker"
	theInstalledObjects(7) = "Scripting.FileSystemObject"
	theInstalledObjects(8) = "adodb.connection"
	theInstalledObjects(9) = "SoftArtisans.FileUp"
	theInstalledObjects(10) = "SoftArtisans.FileManager"
	theInstalledObjects(11) = "JMail.Message"
	theInstalledObjects(12) = "CDONTS.NewMail"
	theInstalledObjects(13) = "Persits.MailSender"
	theInstalledObjects(14) = "LyfUpload.UploadFile"
	theInstalledObjects(15) = "Persits.Upload.1"
	theInstalledObjects(16) = "w3.upload"
	theInstalledObjects(17) = "Persits.Jpeg"
	For i=0 to 17
		Response.Write "<tr class=""a4""><td class=""a2"">&nbsp;" & theInstalledObjects(i) & " &nbsp;"
		Select case i
			case 7
				Response.Write "(FSO �ı��ļ���д)"
			case 8
				Response.Write "(ACCESS ���ݿ�)"
			case 9
				Response.Write "(SA-FileUp �ļ��ϴ�)"
			case 10
				Response.Write "(SA-FM �ļ�����)"
			case 11
				Response.Write "(JMail �ʼ�����)"
			case 12
				Response.Write "(WIN����SMTP ����)"
			case 13
				Response.Write "(ASPEmail �ʼ�����)"
			case 14
				Response.Write "(LyfUpload �ļ��ϴ�)"
			case 15
				Response.Write "(ASPUpload �ļ��ϴ�)"
			case 16
				Response.Write "(w3 upload �ļ��ϴ�)"
			case 17
				Response.Write "(ASPˮӡ)"
		end select
		Response.Write "  " & getver(theInstalledObjects(i)) & "  </td><td><div class=""a2"" style=""width:30;padding:3px;"">"
		If Not IsObjInstalled(theInstalledObjects(i)) Then
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<b>��</b> "
		End If
		Response.Write "</div></td></TR>" & vbCrLf
	Next
	Echo "</table>"
End Sub

Sub Main%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>������ʾ</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>��ҳ��ֻ����ο�ʹ�ã����������Ϊ�����������ܵĲ�ͬ������
      </ul>
	  </td>
  </tr>
</table>
<br>
<form method="post" action="?action=onlinelistok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
      <td>������Ŀ</td>
      <td>����ֵ</td>
    </tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ������������ </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("server_name")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ��������IP��ַ </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("LOCAL_ADDR")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ����������ϵͳ </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("OS")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> �������������� </td>
		<td bgcolor="#F8F8F8"> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ��������������Ƽ��汾 </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("SERVER_SOFTWARE")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> �������������еĶ˿� </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("server_port")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ������CPU���� </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ������Application���� </td>
		<td bgcolor="#F8F8F8"> <%=Application.Contents.Count%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ������Session���� </td>
		<td bgcolor="#F8F8F8"> <%=Session.Contents.Count%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ���������·�� </td>
		<td bgcolor="#F8F8F8"> <%=Request.ServerVariables("path_translated")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> �����URL </td>
		<td bgcolor="#F8F8F8"> http://<%=Request.ServerVariables("server_name")%><%=Request.ServerVariables("script_name")%> </td>
	</tr>
	<tr> 
		<td bgcolor="#FFFFFF"> ��������ǰʱ�� </td>
		<td bgcolor="#F8F8F8"> <%=Now()%> </td>
	</tr>

	<tr> 
		<td bgcolor="#FFFFFF"> �ű����ӳ�ʱʱ�� </td>
		<td bgcolor="#F8F8F8"> <%=Server.ScriptTimeout%> �� </td>
	</tr>
</table><BR>
 <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr align="left" class="a1"> 
		<td> �����ļ������ٶȲ��� </td>
	</tr>
	<tr class="a2"><td> <%
	Response.Flush

	If Not IsObjInstalled("Scripting.FileSystemObject") Then
		Response.Write "��֧��FSO,�޷������ļ���ȡ"
	Else
	Response.Write "�����ظ�������д���ɾ���ı��ļ�50��..."
	Dim thetime3,tempfile,iserr,t1,FsoObj,tempfileOBJ,t2
	Set FsoObj=Server.CreateObject("Scripting.FileSystemObject")
	iserr=False
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	For i=1 To 50
		Err.Clear

		Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		If Err <> 0 Then
			Response.Write "�����ļ�����"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		If Err <> 0 Then
			Response.Write "д���ļ�����"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		If Err <> 0 Then
			Response.Write "ɾ���ļ�����"
			iserr=True
			Err.Clear
			Exit For
		end if
		Set tempfileOBJ=Nothing
	Next
	t2=timer
	If Not iserr Then
		thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
		Response.Write "...����ɣ���������ִ�д˲�������ʱ <font color=red>" & thetime3 & " ����</font>"
	End If
	End If
%>
<BR>P4 2.4,2GddrEcc,SCSI36.4G*2 ִ�д˲�����Ҫ <font color=red>32��65</font> ����</a>
<%Response.Flush%>
</td>
	</tr>
	</table>
	<BR> 
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
	<tr align="left" class="a1"> 
		<td>  ASP�ű����ͺ������ٶȲ���  </td>
	</tr>
	<tr class="a2">
		<td>
<%
	Response.Write "����������ԣ����ڽ���50��μӷ�����..."
	dim lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime & " ����</font><br>"
	Response.Write "����������ԣ����ڽ���20��ο�������..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime2 & " ����</font><br>"
%>
	<BR>P4 2.4,2GddrEcc,SCSI36.4G*2 ����������Ҫ <font color=red>171��203</font> ����, ����������Ҫ <font color=red>156��171</font> ����
	</td></tr>
  </table>
<br>
<br>
<%
End Sub


Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
''''''''''''''''''''''''''''''
Function getver(Classstr)
	On Error Resume Next
	getver=""
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(Classstr)
	If 0 = Err Then getver=xtestobj.version
	Set xTestObj = Nothing
	Err = 0
End Function
footer()
%>
