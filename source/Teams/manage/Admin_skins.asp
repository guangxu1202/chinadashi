<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%

Call Master_Us()
Header()
Dim Admin_Class,StyleConn
Admin_Class=",7,"
Call Master_Se()
Dim LoadName,LoadInfo
team.SaveLog ("����ģ��Ĳ鿴���޸�")
Select Case Request("Menu")
	Case "Add"
		Show()
		Add
	Case "Addok"
		Addok
	Case "Edit"
		Show()
		Edit
	Case "Edithtml"
		Show()
		Edithtml()
	Case "Editok"
		Editok
	Case "Del"
		Dim id
		ID=Cid(Request("id"))
		If Request("model")=1 Then 
			SuccessMsg("�Բ���Ĭ��ģ�岻��ɾ��!")
		Else
			Team.execute("delete from ["&Isforum&"Style] where id="&ID)
			Cache.DelCache("TemplatesLoad")
			SuccessMsg "ģ��ɾ���ɹ���"
		End If
	Case "copy"
		Dim rs,Rsnew,sql
		ID=team.checkStr(request("id"))
		'ȡ���ƶ�����
		Set Rs=Team.execute("Select  StyleName,StyleWid,StyleCss,Style_index,Style_post,Style_user,Style_else,Styleurl From ["&Isforum&"Style] Where ID="&ID )
		'������ģ��
		Set Rsnew=Server.CreateObject("Adodb.RecordSet")
		SQL="Select StyleName,StyleWid,StyleCss,Style_index,Style_post,Style_user,Style_else,Styleurl From ["&Isforum&"Style]"
		If Not IsObject(Conn) Then ConnectionDatabase
		Rsnew.Open SQL,Conn,2,3
		Rsnew.addnew
		Rsnew(0)=Rs(0)
		Rsnew(1)=Rs(1)
		Rsnew(2)=Rs(2)
		Rsnew(3)=Rs(3)
		Rsnew(4)=Rs(4)
		Rsnew(5)=Rs(5)
		Rsnew(6)=Rs(6)
		Rsnew(7)=Rs(7)
		Rsnew.Update
		Rsnew.Close:Set Rsnew=Nothing
		Rs.Close:Set Rs=Nothing
		Cache.DelCache("TemplatesLoad")
		SuccessMsg("<li>ģ�帴�Ƴɹ�!<p><a href=javascript:history.back()>< �� �� ></a></P>")

	Case "loading"
		Show()
		LoadName="����"
		LoadInfo="styleload"
		loading
	Case "output"
		LoadName="����"
		LoadInfo="outputok"
		showstyle
	Case "styleload"
		LoadName="����"
		LoadInfo="loadok"
		showstyle
	Case "outputok"
		outputok
	Case "loadok"
		loadok
	Case Else
		Show()
End Select

Sub loading
%>
<form action="?menu=<%=LoadInfo%>" method="post">
<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center">
<tr><td align="center" class="a1" colspan="2"><%=LoadName%>ģ������</td></tr>
<tr class="a4">
   <td><%=LoadName%>ģ�����ݿ���</td>
   <td><input type="text" name="skinmdb" size="30" value="../skins/TM_Style.mdb"></td>
</tr>
<tr><td align="center" class="a3" colspan="2"><input type="submit" name="submit" value="��һ��"></td></tr>
</form>
<%
End Sub

Sub showstyle
	Dim skinmdb,rs1
	skinmdb=team.checkstr(trim(Request.form("skinmdb")))
	If Request("menu")="styleload" and skinmdb<>"" Then
		SkinConnection(skinmdb)
		If IsFoundTable("Style",1)=False Then
			SuccessMsg("<li>"&mdbname&"���ݿ����Ҳ���ָ�������ݱ����½�������ݱ�")
			Exit Sub
		End IF
		set RS1=StyleConn.Execute("select ID,StyleName,styleurl,StyleWid from ["&Isforum&"Style]")
	Else
		set RS1=team.Execute("select ID,StyleName,styleurl,StyleWid from ["&Isforum&"Style]")
	End If
	If skinmdb="" Then
		skinmdb="../skins/TM_Style.mdb"
	End If
%>
<form action="?menu=<%=LoadInfo%>" method="post">
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center">
<tr><td align="center" class="a1" colspan="5"><%=LoadName%>ģ������</td></tr>
<tr class="a3" align="center">
	<td width="5%">ģ��ID</td>
	<td width="20%">ģ������</td>
	<td width="30%">ģ��·��</td>
	<td width="20%">Ĭ�Ͽ��</td>
	<td>ѡ��</td>
</tr>
<%Do While Not Rs1.Eof%>
<tr class="a4" align="center">
	<td><%=rs1(0)%></td>
	<td><%=rs1(1)%></td>
	<td><%=rs1(2)%></td>
	<td><%=rs1(3)%></td>
	<td><input type="checkbox" name="skid" value="<%=Rs1(0)%>">ѡ��</td>
</tr>
<%	Rs1.MoveNext
	Loop
	%>
<tr class="a4" align="center">
   <td colspan=5><%=LoadName%>ģ�����ݿ���  ������<input type="text" name="skinmdb" size="30" value="<%=skinmdb%>"></td>
</tr>
<tr><td align="center" class=a3 colspan=5><input type="submit" name="submit" value="��һ��"></td></tr>
</form>
<%
End Sub

Sub Show()
%>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center">
	<tr class="a1"><td align="center" colspan="2">TEAM's ��ʾ</td></tr>
	<tr class="a4">
		<td colspan="2">
		<ul><li>�û�����ͨ����̳�ṩ�ķ��˵�ѡ���������л���ǰ�������� http://your.com/Cookies.asp?action=skins&styleid=<Font Color=red>1</font></ul>
		<ul><li>��̳��Ĭ��ģ�岻��ɾ����</ul>
		<ul><li>����޸ķ�ģ��ҳ�����ƻ�ɾ����ģ��ҳ�����ڹر���̳֮�������������ܻ�Ӱ����̳���ʡ�</ul>
		<ul><li>�½����޸�ģ�壬�밴�����ҳ����ʾ������д����Ϣ��</ul></td>
	</tr>
	<tr class="a4"><td colspan="2">
		<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center"><tr class="a4"><td>
		��ݷ�ʽ�� <a href="?menu=Add"> �����µķ�� </a> | <a href="admin_skins.asp">���ر༭ģ����ҳ</a> | <a href="Admincp.asp#��������ʾ��ʽ">����Ĭ����̳���</a></td></tr></table>
		</td></tr>
	</table>
	<br />
	<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center">
	<tr class="a1">
		<td align="center" width="25%">ģ������</td><td align="center" width="15%">ģ��ID</td><td align="center" width="15%">Ĭ��ģ��</td><td align="center" width="15%">�༭ģ��</td><td align="center" width="15%">����ģ��</td><td align="center" width="15%">ɾ��ģ��</td>
	</tr>
	<%
	Dim Rs
	Set Rs=team.Execute("Select ID,StyleName,StyleHid From ["&IsForum&"Style] order by Id Asc")
	Do While Not Rs.eof
		Response.Write "<tr align=center> "
		Response.Write "<td bgcolor=#FFFFFF><input type=text name=stylename value="&Rs(1)&" size=18></td>"
		Response.Write "<td bgcolor=#F8F8F8>"&Rs(0)&"</td>"
		Response.Write "<td bgcolor=#FFFFFF>"
		If Rs(2)=1 Then 
			Response.Write " �� " 
		Else 
			Response.Write " �� "
		End If 
		Response.Write "</td>"
		Response.Write "<td bgcolor=#F8F8F8><a href=admin_skins.asp?menu=Edit&id="&Rs(0)&">�༭ģ��</a></td>"
		Response.Write "<td bgcolor=#FFFFFF><a href=admin_skins.asp?menu=copy&id="&Rs(0)&">����ģ��</a></td>"

		If Rs(2)=1 Then
			Response.Write "<td bgcolor=#FFFFFF><a href=admin_skins.asp?menu=Del&id="&Rs(0)&"&model=1>ɾ��ģ��</a></td>"
		Else
			Response.Write "<td bgcolor=#FFFFFF><a href=admin_skins.asp?menu=Del&id="&Rs(0)&">ɾ��ģ��</a></td>"
		End if

		Response.Write "</tr>"
		Rs.Movenext
	Loop
	Rs.Close:Set Rs = Nothing
%>
	</table><br />
<%
End Sub

Sub Add()
	Dim StyleName,Styleurl,Stylewid,Stylecss
%>
	<form method="post" action="?menu=Addok" name=form>
	<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center">
	<tr class="a1">
		<td colspan="2">ģ�����ҳ��</td>
	</tr>
	<tr class="a3">
		<td width="25%">���༭��� </td><td> <input size="30" name="StyleName" value="<%=StyleName%>"> </td>
	</tr>
	<tr class="a4">
		<td>��������� </td><td> <input size="30" name="Styleurl" value="<%=Styleurl%>"></td>
	</tr>
	<tr class="a3">
		<td>��Ĭ�Ͽ�� </td><td><input size="30" name="Stylewid" value="<%=Stylewid%>"></td>
	</tr>
	<tr class=a3>
		<td align="center" width="100%" colspan="4"> 
		<input type="submit" value=" �� �� ">
		<input type="reset" value=" �� �� " name="Submit2"></td>
	</tr>
</table>
<%
End Sub

Sub Addok
	Dim StyleName,Stylewid,Styleurl,Stylecss,Msg
	StyleName=team.checkStr(Request.Form("StyleName"))
	Stylewid=team.checkStr(Request.Form("Stylewid"))
	Styleurl=team.checkStr(Request.Form("Styleurl"))
	Stylecss=team.checkStr(Request.Form("Editname"))
	If StyleName="" Or Stylewid="" or Styleurl="" Then SuccessMsg ("����д��������!")
	Set RS= Server.CreateObject("ADODB.Recordset")
	RS.Open "["&Isforum&"Style]",Conn,1,3
	RS.addnew
	RS("StyleName")=StyleName
	RS("Styleurl")=Styleurl
	RS("Stylewid")=Stylewid
	RS("Stylecss")=Stylecss
	RS("Styleurl")=Styleurl
	RS("Style_index")="@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
	RS("Style_post")="@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
	RS("Style_user")="@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
	RS("Style_else")="@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
	RS.Update
	RS.Close
	Msg="��ģ����ӳɹ�,��༭ģ������!"
	Cache.DelCache("TemplatesLoad")
	SuccessMsg MSG
End Sub

Sub Editok
	Dim id,name,StyleName,Stylewid,Stylecss,Styleurl,msg
	Dim TempStr,Editname
	ID=team.checkStr(request("Id"))
	Name=Int(request("Name"))
	Select Case Name
		Case 1
			StyleName=team.checkStr(Request.Form("StyleName"))
			Stylewid=team.checkStr(Request.Form("Stylewid"))
			Styleurl=team.checkStr(Request.Form("Styleurl"))
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#��β#@@@","#��β#"))
			team.Execute("update ["&Isforum&"Style] set StyleName='"&StyleName&"',Stylewid='"&Stylewid&"',Stylecss='"&Editname&"',Styleurl='"&Styleurl&"' Where ID="&ID)
			Msg="��̳ģ�������Ϣ�༭�ɹ�"
		Case 2
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#��β#@@@","#��β#"))
			team.Execute("update ["&Isforum&"Style] set Style_index='"&Editname&"' Where ID="&ID)
			Msg="��ҳģ��༭�ɹ� "
		Case 3
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#��β#@@@","#��β#"))
			team.Execute("update ["&Isforum&"Style] set Style_post='"&Editname&"' Where ID="&ID)
			Msg="�����б�ģ��༭�ɹ� "
		Case 4
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#��β#@@@","#��β#"))
			team.Execute("update ["&Isforum&"Style] set Style_user='"&Editname&"' Where ID="&ID)
			Msg="�û�����ģ��༭�ɹ� "
		Case 5
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#��β#@@@","#��β#"))
			team.Execute("update ["&Isforum&"Style] set Style_else='"&Editname&"' Where ID="&ID)
			Msg="�������ģ��༭�ɹ� "
	End Select
	Cache.DelCache("Templates_"&ID)
	SuccessMsg Msg
End Sub

Sub edit()
	Dim id,rs1
	ID=team.checkStr(request("Id"))
	Set Rs1=team.Execute("Select * From ["&Isforum&"Style] Where ID="&Id)
	Do while not rs1.eof
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""a2"" Width=""98%"">"
	Response.Write "<tr>"
	Response.Write "<td width=""100%"" class=""a1"" colspan=2 height=25>"&rs1("StyleName")&" -- ��̳ģ�����"
	Response.Write "</td>"
	Response.Write "</tr>"
	response.write "<tr Class=a4><td height=25 Width=""70%""><li>ģ��������ü�����ģ�� </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=1>�༭</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>��̳��ҳģ������ </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=2>�༭</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>�����б�ģ������ </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=3>�༭</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>�û�����ģ������ </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=4>�༭</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>�������ģ������ </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=5>�༭</a></td></tr>"
	rs1.movenext
	loop
	Rs1.Close:Set Rs1 = Nothing
	response.write "</table><BR><BR><BR><BR>"
End Sub

Sub Edithtml()
	Dim ID,Name,HtmlName,rs,Editname
	ID=team.checkStr(request("Id"))
	Name=Int(request("Name"))
	Set Rs=team.Execute("Select * From ["&Isforum&"Style] Where ID="&Id)
	If (Rs.eof or Rs.bof) Then SuccessMsg("ָ����ģ���ļ�������")
	Select Case Name
		Case 1
			Editname=Split(rs("Stylecss"),"@@@")
			HtmlName="HtmlNews"
		Case 2
			Editname=Split(rs("Style_index"),"@@@")
			HtmlName="IndexHtml"
		Case 3
			Editname=Split(rs("Style_post"),"@@@")
			HtmlName="PostHtml"
		Case 4
			Editname=Split(rs("Style_user"),"@@@")
			HtmlName="UserHtml"
		Case 5
			Editname=Split(rs("Style_else"),"@@@")
			HtmlName="ElseHtml"
	End Select%>
	<form method="post" action="?menu=Editok" name=form>
	<input type=hidden name=id value=<%=Request("id")%>>
	<input type=hidden name=Name value=<%=Request("Name")%>>
	<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center" style="word-break:break-all">
	<tr class=a1>
		<td>��̨���� --> ģ��༭ [<%=rs("StyleName")%>]</td>
	</tr>
	<%If Name = "1" Then%>
		<tr class="a3">
			<td>ģ������ <input size="30" name="StyleName" value="<%=Rs("StyleName")%>"></td>
		</tr>
		<tr class="a4">
		<td>ģ��·�� <input size="30" name="Styleurl" value="<%=Rs("Styleurl")%>"></td>
		</tr>
		<tr class="a3">
			<td>Ĭ�Ͽ�� <input size="30" name="Stylewid" value="<%=RS("StyleWid")%>"></td>
		</tr>
	<%	For i=0 To Ubound(Editname)	%>
			<tr class="a4">
				<td align="center" width="100%">Team.<%=HtmlName%> (<%=i%>)ģ���ļ�<br><textarea name="Editname" rows="10" cols="150" style="height:70;overflow-y:visible;"><%=server.htmlencode(Editname(i))%></textarea>
			</td>
			</tr>
	<%
		Next
	Else
		For i=0 To Ubound(Editname)%>
			<tr class="a4">
				<td align="center" width="100%">Team.<%=HtmlName%> (<%=i%>)ģ���ļ�<br><textarea name="Editname" rows="10" cols="150" style="height:70;overflow-y:visible;"><%=server.htmlencode(Editname(i))%></textarea>
			</td>
			</tr>
	<%
		Next
	End If
%>
	<tr class=a4>
		<td width="100%" colspan="4"> 
		<b>�ر�˵��: </b><BR>
		<li>ģ������һ���ļ����ݱ���Ϊ[   #��β#   ]. ������Ϊģ���ļ��Ľ�β����.</li>
		<li>�����Ҫ���ģ�����Ŀ,�뽫[   #��β#   ]�������޸�Ϊ����ֵ,�༭��ɺ�,������β�����[   #��β#   ],��Ϊģ���β.</li>
		</td>
	</tr>
	<tr class=a3>
		<td align="center" width="100%" colspan="4"> 
		<input type="submit" value=" �� �� ">
		<input type="reset" value=" �� �� " name="Submit2"></td>
	</tr>
</table>
</form>
<%
End Sub

Sub loadok
	Dim tRs,skid,mdbname
	skid=team.checkstr(Request("skid"))
	mdbname=team.Checkstr(trim(Request.form("skinmdb")))
	If skid="" or isnull(skid) or Not Isnumeric(Replace(Replace(skid,",","")," ","")) Then
		SuccessMsg("<li>����δѡȡҪ�����ģ��")
		Exit Sub
	End If
	If mdbname="" Then
		SuccessMsg("<BR><li>����д����ģ�����ݿ���")
		Exit Sub
	End If
	SkinConnection(mdbname)
	If IsFoundTable("Style",1)=False Then
		SuccessMsg("<li>"&mdbname&"���ݿ����Ҳ���ָ�������ݱ����½�������ݱ�")
		Exit Sub
	End IF
	Dim InsertName,InsertValue
	Set TRs=StyleConn.Execute("select * from ["&Isforum&"Style] where id in ("&skid&")  order by id ")
	Do while not TRs.eof
	InsertName=""
	InsertValue=""
		For i = 1 to TRs.Fields.Count-1
			InsertName=InsertName & TRs(i).Name
			InsertValue=InsertValue & "'" &team.checkStr(TRs(i)) & "'"
			If i<> TRs.Fields.Count-1 Then 
				InsertName	= InsertName & ","
				InsertValue	= InsertValue & ","
			End If
		Next
	team.Execute("insert into ["&Isforum&"Style] ("&InsertName&") values ("&InsertValue&") ")
	TRs.movenext
	loop
	TRs.close
	set Rs=nothing
	set TRs=nothing
	SuccessMsg("���ݵ���ɹ���")
End Sub

Sub outputok
	Dim TempRs,skid,mdbname
	skid=team.checkstr(Request("skid"))
	mdbname=team.Checkstr(Trim(Request.form("skinmdb")))
	If skid="" or Isnull(skid) or Not IsNumeric(Replace(Replace(skid,",","")," ","")) Then
		SuccessMsg("<li>����δѡȡҪ������ģ�棬������д���")
		Exit Sub
	End If
	If mdbname="" Then
		SuccessMsg("<li>������д����ģ�����ݿ���")
		Exit Sub
	End If
	SkinConnection(mdbname)
	If IsFoundTable("Style",1)=False Then
		SuccessMsg("<li>"&mdbname&"���ݿ����Ҳ���ָ�������ݱ����½�������ݱ�")
		Exit Sub
	End IF
	set Rs=team.Execute("select * from ["&Isforum&"Style] where id in ("&skid&") order by id ")
	If Rs.EOF Or Rs.BOF Then
		SuccessMsg("<BR><li>�޷�ȡ��Դģ������")
		Exit Sub
	End If
	Dim InsertName,InsertValue
	Do while not Rs.eof
	InsertName=""
	InsertValue=""
	For i = 1 to Rs.Fields.Count-1
		InsertName=InsertName & Rs(i).Name
		InsertValue=InsertValue & "'" & team.checkStr(Rs(i)) & "'"
		If i<> Rs.Fields.Count-1 Then 
			InsertName	= InsertName & ","
			InsertValue	= InsertValue & ","
		End If
	Next
	StyleConn.Execute("insert into ["&Isforum&"Style] ("&InsertName&") values ("&InsertValue&") ")
	Rs.movenext
	loop
	Rs.close
	set Rs=nothing
	SuccessMsg("<li>���ݵ����ɹ���")
End Sub

'У������Ƿ���ڡ�TableName=������str:0=Ĭ�Ͽ⣬1=����
Function IsFoundTable(TableName,Str)
	Dim ChkRs
	IsFoundTable=False
	If TableName<>"" Then 
	TableName=LCase(Trim(TableName))
		If Str=0 Then
		Set ChkRs=Conn.openSchema(20)
		Else
		Set ChkRs=StyleConn.openSchema(20)
		End If
		Do Until ChkRs.EOF
			If ChkRs("TABLE_TYPE")="TABLE" Then
				If Lcase(ChkRs("TABLE_NAME"))=TableName then
					IsFoundTable=True
					Exit Function
				End If
			End If
		ChkRs.movenext
		Loop
		ChkRs.close:Set ChkRs=Nothing
	End If
End Function

Sub SkinConnection(mdbname)
	On Error Resume Next 
	Set StyleConn = Server.CreateObject("ADODB.Connection")
	StyleConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	If Err.Number ="-2147467259"  Then 
		SuccessMsg("<li>"&Server.MapPath(mdbname)&"���ݿⲻ���ڡ�")
		Response.end
	End If
End Sub

If IsObject(StyleConn) Then
	StyleConn.close
	Set StyleConn=Nothing
End IF

%>