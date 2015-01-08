<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%

Call Master_Us()
Header()
Dim Admin_Class,StyleConn
Admin_Class=",7,"
Call Master_Se()
Dim LoadName,LoadInfo
team.SaveLog ("进行模板的查看和修改")
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
			SuccessMsg("对不起，默认模板不能删除!")
		Else
			Team.execute("delete from ["&Isforum&"Style] where id="&ID)
			Cache.DelCache("TemplatesLoad")
			SuccessMsg "模板删除成功。"
		End If
	Case "copy"
		Dim rs,Rsnew,sql
		ID=team.checkStr(request("id"))
		'取出制定数据
		Set Rs=Team.execute("Select  StyleName,StyleWid,StyleCss,Style_index,Style_post,Style_user,Style_else,Styleurl From ["&Isforum&"Style] Where ID="&ID )
		'拷入新模板
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
		SuccessMsg("<li>模板复制成功!<p><a href=javascript:history.back()>< 返 回 ></a></P>")

	Case "loading"
		Show()
		LoadName="导入"
		LoadInfo="styleload"
		loading
	Case "output"
		LoadName="导出"
		LoadInfo="outputok"
		showstyle
	Case "styleload"
		LoadName="导入"
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
<tr><td align="center" class="a1" colspan="2"><%=LoadName%>模版数据</td></tr>
<tr class="a4">
   <td><%=LoadName%>模版数据库名</td>
   <td><input type="text" name="skinmdb" size="30" value="../skins/TM_Style.mdb"></td>
</tr>
<tr><td align="center" class="a3" colspan="2"><input type="submit" name="submit" value="下一步"></td></tr>
</form>
<%
End Sub

Sub showstyle
	Dim skinmdb,rs1
	skinmdb=team.checkstr(trim(Request.form("skinmdb")))
	If Request("menu")="styleload" and skinmdb<>"" Then
		SkinConnection(skinmdb)
		If IsFoundTable("Style",1)=False Then
			SuccessMsg("<li>"&mdbname&"数据库中找不到指定的数据表，请新建风格数据表；")
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
<tr><td align="center" class="a1" colspan="5"><%=LoadName%>模版数据</td></tr>
<tr class="a3" align="center">
	<td width="5%">模板ID</td>
	<td width="20%">模板名称</td>
	<td width="30%">模板路径</td>
	<td width="20%">默认宽度</td>
	<td>选择</td>
</tr>
<%Do While Not Rs1.Eof%>
<tr class="a4" align="center">
	<td><%=rs1(0)%></td>
	<td><%=rs1(1)%></td>
	<td><%=rs1(2)%></td>
	<td><%=rs1(3)%></td>
	<td><input type="checkbox" name="skid" value="<%=Rs1(0)%>">选择</td>
</tr>
<%	Rs1.MoveNext
	Loop
	%>
<tr class="a4" align="center">
   <td colspan=5><%=LoadName%>模版数据库名  　　　<input type="text" name="skinmdb" size="30" value="<%=skinmdb%>"></td>
</tr>
<tr><td align="center" class=a3 colspan=5><input type="submit" name="submit" value="下一步"></td></tr>
</form>
<%
End Sub

Sub Show()
%>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center">
	<tr class="a1"><td align="center" colspan="2">TEAM's 提示</td></tr>
	<tr class="a4">
		<td colspan="2">
		<ul><li>用户可以通过论坛提供的风格菜单选项来轻松切换当前界面风格，如 http://your.com/Cookies.asp?action=skins&styleid=<Font Color=red>1</font></ul>
		<ul><li>论坛的默认模板不能删除。</ul>
		<ul><li>如果修改分模板页面名称或删除分模板页面请在关闭论坛之后操作，否则可能会影响论坛访问。</ul>
		<ul><li>新建和修改模板，请按照相关页面提示完整填写表单信息。</ul></td>
	</tr>
	<tr class="a4"><td colspan="2">
		<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center"><tr class="a4"><td>
		快捷方式： <a href="?menu=Add"> 创建新的风格 </a> | <a href="admin_skins.asp">返回编辑模板首页</a> | <a href="Admincp.asp#界面与显示方式">设置默认论坛风格</a></td></tr></table>
		</td></tr>
	</table>
	<br />
	<table cellspacing="1" cellpadding="3" width="98%" border="0" class="a2" align="center">
	<tr class="a1">
		<td align="center" width="25%">模版名称</td><td align="center" width="15%">模版ID</td><td align="center" width="15%">默认模版</td><td align="center" width="15%">编辑模版</td><td align="center" width="15%">复制模版</td><td align="center" width="15%">删除模版</td>
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
			Response.Write " √ " 
		Else 
			Response.Write " × "
		End If 
		Response.Write "</td>"
		Response.Write "<td bgcolor=#F8F8F8><a href=admin_skins.asp?menu=Edit&id="&Rs(0)&">编辑模板</a></td>"
		Response.Write "<td bgcolor=#FFFFFF><a href=admin_skins.asp?menu=copy&id="&Rs(0)&">复制模板</a></td>"

		If Rs(2)=1 Then
			Response.Write "<td bgcolor=#FFFFFF><a href=admin_skins.asp?menu=Del&id="&Rs(0)&"&model=1>删除模板</a></td>"
		Else
			Response.Write "<td bgcolor=#FFFFFF><a href=admin_skins.asp?menu=Del&id="&Rs(0)&">删除模板</a></td>"
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
		<td colspan="2">模板添加页面</td>
	</tr>
	<tr class="a3">
		<td width="25%">　编辑风格 </td><td> <input size="30" name="StyleName" value="<%=StyleName%>"> </td>
	</tr>
	<tr class="a4">
		<td>　风格排序 </td><td> <input size="30" name="Styleurl" value="<%=Styleurl%>"></td>
	</tr>
	<tr class="a3">
		<td>　默认宽度 </td><td><input size="30" name="Stylewid" value="<%=Stylewid%>"></td>
	</tr>
	<tr class=a3>
		<td align="center" width="100%" colspan="4"> 
		<input type="submit" value=" 编 辑 ">
		<input type="reset" value=" 重 填 " name="Submit2"></td>
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
	If StyleName="" Or Stylewid="" or Styleurl="" Then SuccessMsg ("请填写完整参数!")
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
	Msg="新模板添加成功,请编辑模板内容!"
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
			Editname=team.checkStr(Replace(Editname,"#结尾#@@@","#结尾#"))
			team.Execute("update ["&Isforum&"Style] set StyleName='"&StyleName&"',Stylewid='"&Stylewid&"',Stylecss='"&Editname&"',Styleurl='"&Styleurl&"' Where ID="&ID)
			Msg="论坛模板基本信息编辑成功"
		Case 2
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#结尾#@@@","#结尾#"))
			team.Execute("update ["&Isforum&"Style] set Style_index='"&Editname&"' Where ID="&ID)
			Msg="首页模板编辑成功 "
		Case 3
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#结尾#@@@","#结尾#"))
			team.Execute("update ["&Isforum&"Style] set Style_post='"&Editname&"' Where ID="&ID)
			Msg="帖子列表模板编辑成功 "
		Case 4
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#结尾#@@@","#结尾#"))
			team.Execute("update ["&Isforum&"Style] set Style_user='"&Editname&"' Where ID="&ID)
			Msg="用户属性模板编辑成功 "
		Case 5
			For Each TempStr in Request.Form("Editname")
				Editname=Editname & Replace(TempStr,"@@@","")&"@@@"
			Next
			Editname=team.checkStr(Replace(Editname,"#结尾#@@@","#结尾#"))
			team.Execute("update ["&Isforum&"Style] set Style_else='"&Editname&"' Where ID="&ID)
			Msg="其他版块模板编辑成功 "
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
	Response.Write "<td width=""100%"" class=""a1"" colspan=2 height=25>"&rs1("StyleName")&" -- 论坛模板管理"
	Response.Write "</td>"
	Response.Write "</tr>"
	response.write "<tr Class=a4><td height=25 Width=""70%""><li>模板基本设置及调用模版 </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=1>编辑</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>论坛首页模板设置 </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=2>编辑</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>帖子列表模板设置 </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=3>编辑</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>用户属性模板设置 </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=4>编辑</a></td></tr>"
	response.write "<tr Class=a4><td height=25><li>其他版块模板设置 </td><td height=25><a href=admin_skins.asp?menu=Edithtml&id="&rs1("id")&"&name=5>编辑</a></td></tr>"
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
	If (Rs.eof or Rs.bof) Then SuccessMsg("指定的模板文件不存在")
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
		<td>后台管理 --> 模板编辑 [<%=rs("StyleName")%>]</td>
	</tr>
	<%If Name = "1" Then%>
		<tr class="a3">
			<td>模板名称 <input size="30" name="StyleName" value="<%=Rs("StyleName")%>"></td>
		</tr>
		<tr class="a4">
		<td>模板路径 <input size="30" name="Styleurl" value="<%=Rs("Styleurl")%>"></td>
		</tr>
		<tr class="a3">
			<td>默认宽度 <input size="30" name="Stylewid" value="<%=RS("StyleWid")%>"></td>
		</tr>
	<%	For i=0 To Ubound(Editname)	%>
			<tr class="a4">
				<td align="center" width="100%">Team.<%=HtmlName%> (<%=i%>)模板文件<br><textarea name="Editname" rows="10" cols="150" style="height:70;overflow-y:visible;"><%=server.htmlencode(Editname(i))%></textarea>
			</td>
			</tr>
	<%
		Next
	Else
		For i=0 To Ubound(Editname)%>
			<tr class="a4">
				<td align="center" width="100%">Team.<%=HtmlName%> (<%=i%>)模板文件<br><textarea name="Editname" rows="10" cols="150" style="height:70;overflow-y:visible;"><%=server.htmlencode(Editname(i))%></textarea>
			</td>
			</tr>
	<%
		Next
	End If
%>
	<tr class=a4>
		<td width="100%" colspan="4"> 
		<b>特别说明: </b><BR>
		<li>模板的最后一个文件内容必须为[   #结尾#   ]. 用来作为模板文件的结尾部分.</li>
		<li>如果需要添加模板的栏目,请将[   #结尾#   ]的内容修改为其他值,编辑完成后,返回在尾部添加[   #结尾#   ],作为模板结尾.</li>
		</td>
	</tr>
	<tr class=a3>
		<td align="center" width="100%" colspan="4"> 
		<input type="submit" value=" 编 辑 ">
		<input type="reset" value=" 重 填 " name="Submit2"></td>
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
		SuccessMsg("<li>您还未选取要导入的模版")
		Exit Sub
	End If
	If mdbname="" Then
		SuccessMsg("<BR><li>请填写导入模版数据库名")
		Exit Sub
	End If
	SkinConnection(mdbname)
	If IsFoundTable("Style",1)=False Then
		SuccessMsg("<li>"&mdbname&"数据库中找不到指定的数据表，请新建风格数据表；")
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
	SuccessMsg("数据导入成功！")
End Sub

Sub outputok
	Dim TempRs,skid,mdbname
	skid=team.checkstr(Request("skid"))
	mdbname=team.Checkstr(Trim(Request.form("skinmdb")))
	If skid="" or Isnull(skid) or Not IsNumeric(Replace(Replace(skid,",","")," ","")) Then
		SuccessMsg("<li>您还未选取要导出的模版，或参数有错误！")
		Exit Sub
	End If
	If mdbname="" Then
		SuccessMsg("<li>请请填写导出模版数据库名")
		Exit Sub
	End If
	SkinConnection(mdbname)
	If IsFoundTable("Style",1)=False Then
		SuccessMsg("<li>"&mdbname&"数据库中找不到指定的数据表，请新建风格数据表；")
		Exit Sub
	End IF
	set Rs=team.Execute("select * from ["&Isforum&"Style] where id in ("&skid&") order by id ")
	If Rs.EOF Or Rs.BOF Then
		SuccessMsg("<BR><li>无法取出源模版数据")
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
	SuccessMsg("<li>数据导出成功！")
End Sub

'校验表名是否存在。TableName=表名，str:0=默认库，1=风格库
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
		SuccessMsg("<li>"&Server.MapPath(mdbname)&"数据库不存在。")
		Response.end
	End If
End Sub

If IsObject(StyleConn) Then
	StyleConn.close
	Set StyleConn=Nothing
End IF

%>