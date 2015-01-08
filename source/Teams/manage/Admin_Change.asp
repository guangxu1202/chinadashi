<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Public boards
Dim Admin_Class,Uid
Call Master_Us()
Uid = Cid(Request("uid"))
Header()
Admin_Class=",8,"
Call Master_Se()
Select Case Request("action")
	Case "medals"
		Call medals
	Case "medalsok"
		Call medalsok
	Case "announcements"
		Call announcements
	Case "newsannouncements"
		Call newsannouncements
	Case "announcementsok"
		Call announcementsok
	Case "forumlinks"
		Call forumlinks
	Case "forumlinksok"
		Call forumlinksok
	Case "adv"
		Call adv
	Case "advadd"
		Call advadd
	Case "advaddok"
		Call advaddok
	Case "advok"
		Call advok
	Case "advedit"
		Call advedit
	Case "onlinelist"
		Call onlinelist
	Case "onlinelistok"
		Call onlinelistok
End Select

Sub onlinelistok
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"OnlineTypes] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"OnlineTypes] set Sorts="&Cid(Request.Form("sorts"&i+1))&",OnlineName='"&Replace(Request.Form("titles"&i+1),"'","")&"',Onlineimg='"&Replace(Request.Form("urls"&i+1),"'","")&"' Where ID="&newid(i))
		Next
		if Request.Form("newMembers")<>"" and Request.Form("newurl")<>"" Then
			Dim mTitle
			mTitle = ""
			mTitle = Replace(Request.Form("newMembers"),"'","")
			If InStr(Trim(mTitle),"游客/未登陆")>0 Then mTitle ="游客"
			team.execute("insert into ["&Isforum&"OnlineTypes] (Sorts,OnlineName,Onlineimg) values ("&Cid(Request.Form("newsorts"))&",'"&mTitle&"','"&Replace(Request.Form("newurl"),"'","")&"' ) ")
		End if
	End If
	Application.Contents.RemoveAll()
	team.SaveLog ("在线图列设置完成")
	SuccessMsg " 在线图列设置完成，请等待系统自动返回到 <a href=Admin_Change.asp?action=onlinelist>在线列表定制 </a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=onlinelist>。 "
End Sub

Sub onlinelist%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>本功能用于自定义首页及主题列表页显示的在线会员分组及图例，只在在线列表功能打开时有效。
      </ul>
      <ul>
        <li>所有未添加显示的用户组成员将不显示在在线列表处。
      </ul>
      <ul>
        <li>用户组图例中请填写图片文件名，并将相应图片文件上传到 Skins/下面的相应风格目录中。
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=onlinelistok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
	  <td><input type="checkbox" name="chkall" onClick="checkall(this.form)"> 删</td>
      <td>显示顺序</td>
      <td>组头衔</td>
      <td>用户组图例</td>
    </tr>
	<%
	Dim Rs,Imgs,i,Styleurl
	Set Rs = team.execute("Select Styleurl From ["&Isforum&"Style] Where ID= "& INT(team.Forum_setting(18)))
	If Not Rs.Eof Then
		Styleurl = Rs(0)
	Else
		Styleurl = "skins/teams"
	End If
	Rs.close:Set Rs=Nothing
	Set Rs = team.execute("Select ID,Sorts,OnlineName,Onlineimg From ["&isforum&"OnlineTypes] Order By Sorts Asc")
	If Rs.Eof Then
		SuccessMsg " 未找到数据表，请确认数据库已经升级。"
	End if
	i=0
	Do While Not Rs.Eof
		i=i+1
			If Rs(3)<>"" Then 
				Imgs = "<img src=../"&Styleurl&"/"&RS(3)&" align=""absmiddle"">"
			Else
				Imgs = ""
			End if
			Echo " <tr align=""center""> <td bgcolor=""#FFFFFF""><input Name=""newid"" type=""hidden"" value="&RS(0)&"> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&"></td>	"
			Echo " <td bgcolor=""#F8F8F8""><input type=""text"" size=""3"" name=""sorts"&i&""" value="&Rs(1)&"></td>"
			Echo " <td bgcolor=""#FFFFFF""><input type=""text"" size=""15"" name=""titles"&i&""" value="&Rs(2)&"></td>"
			Echo " <td bgcolor=""#F8F8F8"" align=""left""><input type=""text"" size=""40"" name=""urls"&i&""" value="&Rs(3)&"> "
			Echo " "&Imgs&"</td></tr>"
		Rs.moveNext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
    <tr align="center" class="a1">
	  <td> 新增 </td>
      <td>显示顺序</td>
      <td>组头衔</td>
      <td>用户组图例</td>
    </tr>
    <tr align="center" class="a3">
	  <td> &nbsp; </td>
      <td> <input type="text" size="3" name="newsorts" value="0"> </td>
      <td><select name="newMembers" style="width:100%">
		<option value=""> 请选择用户组 </option>
		<%
		Dim Gs
		Set Gs = team.execute("Select ID,GroupName,Members From "&IsForum&"UserGroup Where MemberRank<=1 Order By ID ASC")
		Do While Not Gs.Eof
			If Gs(2)="嘉宾" Then
				Echo " <option value="""&Gs(1)&"""> "&Gs(1)&" </option> "
			Else
				Echo " <option value="""&Gs(2)&"""> "&Gs(2)&" </option> "
			End if
			Gs.MoveNext
		Loop
		Gs.Close:Set Gs=Nothing
		%>
		</select></td>
      <td align="left"> <input type="text" size="40" name="newurl" value=""> </td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="onlinesubmit" value="提 交">
  </center>
</form>
</td>
</tr>
<br>
<br>
<%
End Sub

Sub advedit
	Dim Rs
	If Uid = "" Or Not IsNumeric(Uid) Then
		SuccessMsg " 参数错误。 "
	Else
		Set Rs = team.execute("Select Titles,Types,Boards,StarTime,StopTime,bodys From ["&Isforum&"AdvList] Where ID="& uid)
		If Rs.Eof Then
			SuccessMsg " 参数错误。 "
		Else	
		%>
<br>
<Script>
function findobj(n, d) {
	var p, i, x;
	if(!d) d = document;
	if((p = n.indexOf("?"))>0 && parent.frames.length) {
		d = parent.frames[n.substring(p + 1)].document;
		n = n.substring(0, p);
	}
	if(x != d[n] && d.all) x = d.all[n];
	for(i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
	for(i = 0; !x && d.layers && i < d.layers.length; i++) x = findobj(n, d.layers[i].document);
	if(!x && document.getElementById) x = document.getElementById(n);
	return x;
}
</Script>
<BR>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" name="settings" action="?action=advaddok&edit=1">
  <input type="hidden" name="uid" value="<%=UID%>">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">编辑广告 - <%=RS(1)%></td>
    </tr>
    <tr>
      <td width="50%" bgcolor="#F8F8F8" ><b>广告标题(必填):</b><br>
        <span class="a3">注意: 广告标题只为识别辨认不同广告条目之用，并不在广告中显示</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="advtitle" value="<%=RS(0)%>">
      </td>
    </tr>
    <tr>
      <td width="50%" bgcolor="#F8F8F8" valign="top"><b>广告投放范围(必选):</b><br>
        <span class="a3">设置本广告投放的页面或论坛范围，可以按住 CTRL 多选，选择“全部”为不限制选择广告投放的范围</span></td>
      <td bgcolor="#FFFFFF"><select name="advtargets" size="10" multiple="multiple">
          <%
			Dim IsOk 
			If Instr(Rs(2),",")>0 or IsNumEric(Rs(2)) Then
				boards = Split(Rs(2),",")
			Else
				IsOk = Rs(2)
			End if
			Response.Write "<option value=""all"" "
			If IsOk ="all" Or Isok="index" Then Response.Write "selected=""selected"" "
			Response.Write " >&nbsp;&nbsp;> 全部</option> " 
			'Response.Write "<option value=""index"" "
			'If IsOk ="index" Then Response.Write "selected=""selected"" "
			'Response.Write " >&nbsp;&nbsp;> 首页</option> " 
			Call BBsList(0)
			%>
        </select>
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>生效时间:</b><br>
        <span class="a3">设置广告广告结束的时间，格式 yyyy-mm-dd，留空为不限制结束时间</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="startime" value="<%=RS(3)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>到期时间:</b><br>
        <span class="a3">设置广告广告结束的时间，格式 yyyy-mm-dd，留空为不限制结束时间</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="stoptime" value="<%=RS(4)%>">
      </td>
    </tr>
    <tr>
      <td width="50%" bgcolor="#F8F8F8" ><b>广告 html 代码:</b><br>
        <span class="a3">请直接输入需要展现的广告的 html 代码</span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="advcode" cols="40" style="height:70;overflow-y:visible;"><%=RS(5)%></textarea>
      </td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="advsubmit" value="提 交">
  </center>
</form>
<br>
<br>
<%		End If
	End if
End Sub
Sub advaddok
	Dim Bodys,textsize,imagewidth,imageheight,imagealt
	If Request.Form("advtitle")&""="" Then  SuccessMsg "请输入标题。"
	If Request("edit") = 1 Then
		If Uid = "" Or Not IsNumeric(Uid) Then
			SuccessMsg " 参数错误。 "
		Else
			team.execute("Update ["&Isforum&"AdvList] set Titles='"&Replace(Request.Form("advtitle"),"'","")&"',bodys='"&Replace(Request.Form("advcode"),"'","")&"',Boards='"&Replace(Request.form("advtargets")," ","")&"',StarTime='"&Replace(Request.form("startime")," ","")&"',StopTime='"&Replace(Request.form("stoptime")," ","")&"' Where ID="& UID )
		End if
	Else
		Select Case Request.Form("advnewstyle")
			Case "code"
				If Request.Form("advcode")&""="" Then  
					SuccessMsg "请输入内容。"
				Else
					Bodys = Request.Form("advcode")
				End If
			Case "text"
				If Request.Form("textlink")&""="" or Request.Form("texttitle")&""="" Then  
					SuccessMsg "请输入必填内容。"
				End if
				If Request.Form("textsize") <> "" Then
					textsize = " Style=""font-size:"& HtmlEncode(Request.Form("textsize")) &"""  "
				End if
				Bodys = "<a href="""& HtmlEncode(Request.Form("textlink")) &""" target=""_blank"" "&textsize&"> "& HtmlEncode(Request.Form("texttitle")) &" </a>"
			Case "image"
				If Request.Form("imageurl")&""="" or Request.Form("imagelink")&""="" Then  
					SuccessMsg "请输入必填内容。"
				End if
				If Request.Form("imagewidth")<>"" and Isnumeric(Request.Form("imagewidth")) Then
					imagewidth = " width="""& Request.Form("imagewidth")&""""
				End if
				If Request.Form("imageheight")<>"" and Isnumeric(Request.Form("imageheight")) Then
					imageheight = " height="""&Request.Form("imageheight")&""""
				End if
				if Request.Form("imagealt")<>"" Then
					imagealt = " alt="""& Request.Form("imagealt")&""" " 
				End if
				Bodys = "<a href="""& HtmlEncode(Request.Form("imagelink")) &""" target=""_blank""><img src="""& HtmlEncode(Request.Form("imageurl")) &""" Border=""0"" align=""absmiddle"" "&imagewidth&" "&imageheight&" "& imagealt&"> </a>"
			Case "flash"
				If Request.Form("flashurl")&""="" Then 
					SuccessMsg "请输入FLASH地址。"
				End If
				If Request.Form("flashwidth")="" or Not IsNumeric(Request.Form("flashwidth")) Then
					SuccessMsg " 宽度必须为数字 。"
				ElseIf Request.Form("flashheight")="" or Not IsNumeric(Request.Form("flashheight")) Then
					SuccessMsg " 高度必须为数字 。"
				Else
					Bodys = "<embed width="""&Request.Form("flashwidth") &""" height="""&Request.Form("flashheight") &""" src="""& HtmlEncode(Request.Form("flashurl")) &""" type=""application/x-shockwave-flash""></embed>"
				End If
		End Select
		team.execute("insert into ["&Isforum&"AdvList] (Dois,Sorts,Titles,Types,StarTime,StopTime,bodys,Boards) values (1,0,'"&Replace(Request.Form("advtitle"),"'","")&"','"&Replace(Request.Form("types"),"'","")&"','"&Replace(Request.Form("startime"),"'","")&"','"&Replace(Request.Form("stoptime"),"'","")&"','"&Replace(Bodys,"'","")&"','"&Replace(Request.form("advtargets")," ","")&"') ")
	End if
	Cache.DelCache("ForumAdvsLoad")
	team.SaveLog ("广告设置完成")
	SuccessMsg " 广告设置完成  ，请等待系统自动返回到 <a href=Admin_Change.asp?action=adv>广告设置</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=adv>。 " 
End Sub

Sub advadd
	Dim tmp,tmp1,tmp2,tmp3
	Select Case Request("type")
		Case "headerbanner"
			tmp = "头部横幅广告显示于论坛页面右上方，通常使用 468x60 图片或 Flash 的形式。当前页面有多个头部横幅广告时，系统会随机选取其中之一显示。"
			tmp1 = "由于能够在页面打开的第一时间将广告内容展现于最醒目的位置，因此成为了网页中价位最高、最适合进行商业宣传或品牌推广的广告类型之一。"
			tmp2 = "头部横幅广告"
			tmp3 = 1
		Case "footerbanner"
			tmp = "尾部横幅广告显示于论坛页面中下方，通常使用 468x60 或其他尺寸图片、Flash 的形式。当前页面有多个尾部横幅广告时，系统会随机选取其中之一显示。 "
			tmp1 = "与页面头部和中部相比，页面尾部的展现机率相对较低，通常不会引起访问者的反感，同时又基本能够覆盖所有对广告内容感兴趣的受众，因此适合中性而温和的推广。"
			tmp2 = "尾部横幅广告"
			tmp3 = 2
		Case "text"
			tmp = "页内文字广告以表格的形式，显示于首页、主题列表和帖子内容三个页面的中上方，通常使用文字的形式，也可使用小图片和 Flash。当前页面有多个文字广告时，系统会以表格的形式按照设定的显示顺序全部展现，同时能够对表格列数在 3~5 的范围内动态排布，以自动实现最佳的广告排列效果。"
			tmp1 = "由于此类广告通常以文字形式展现，但其所在的较靠上的页面位置，使得此类广告成为了访问者必读的内容之一。同一页面可以呈现多达十几条文字广告的特性，也决定了它是一种平民化但性价比较高的推广方式，同时还可用于论坛自身的宣传和公告之用。"
			tmp2 = "页内文字广告"
			tmp3 = 3
		Case "thread"
			tmp = "帖内广告显示于帖子标题的上方，通常使用文字的形式。当前页面有多个帖内广告时，系统会从中抽取与每页帖数相等的条目进行随机显示。"
			tmp1 = "由于帖子是论坛最核心的组成部分，位于帖子内容上方的帖内广告，便可在用户浏览帖子内容时自然的被接受，加上随机播放的特性，适合于特定内容的有效推广，也可用于论坛自身的宣传和公告之用。建议设置多条帖内广告以实现广告内容的差异化，从而吸引更多访问者的注意力。"
			tmp2 = "帖内广告"
			tmp3 = 4
		Case "float"
			tmp = "漂浮广告展现于页面左下角，当页面滚动时广告会自行移动以保持原来的位置，通常使用小图片或 Flash 的形式。当前页面有多个漂浮广告时，系统会随机选取其中之一显示。"
			tmp1 = " 漂浮广告是进行强力商业推广的有效手段，其在页面中的浮动性，使其与固定的图片和文字相比，更容易被关注，正因为如此，这种强制性的关注也可能招致对此广告内容不感兴趣的访问者的反感。请注意不要将过大的图片或 Flash 以漂浮广告的形式显示，以免影响页面阅读。"
			tmp2 = "漂浮广告"
			tmp3 = 5
		Case "couplebanner"
			tmp = "对联广告以长方形图片的形式显示于页面顶部两侧，形似一幅对联，通常使用宽小高大的长方形图片或 Flash 的形式。对联广告一般只在使用像素约定主表格宽度的情况下使用，如使用超过 90% 以上的百分比约定主表格宽度时，可能会影响访问者的正常流量。当访问者浏览器宽度小于 800 像素时，自动不显示此类广告。当前页面有多个对联广告时，系统会随机选取其中之一显示。"
			tmp1 = "对联广告由于只展现于高分辨率(1024x768 或更高)屏幕的两侧，只占用页面的空白区域，因此不会招致访问者反感，能够良好的突出推广内容。但由于对分辨率和主表格宽度的特殊要求，使得广告的受众比例无法达到 100%。"
			tmp2 = "对联广告"
			tmp3 = 6
		Case "affbanner"
			tmp = "公告位广告以长方形图片的形式显示于页面左侧，通常使用宽小高大的长方形图片或 Flash 的形式。公告位广告一般占用公告栏左侧。当前页面有多个公告位广告时，系统会随机选取其中之一显示。"
			tmp1 = "公告位广告由于只展现于公告栏处，有一定的片面性，但是由于寄存于公告栏，对点击率有一定影响。"
			tmp2 = "公告位广告"
			tmp3 = 7

		Case "threadleft"
			tmp = "主题帖子广告以长方形图片的形式显示于页面左侧，通常使用宽小高大的长方形图片或 Flash 的形式。主题帖子位广告一般占用主题帖子栏右侧。点击率比较大,所以适当的设置可以带来很大的流量."
			tmp1 = "主题帖子位广告由于只展现于主题帖子栏处，有一定的片面性，但是由于寄存于主题帖子栏，对点击率有一定影响。"
			tmp2 = "主题帖子位广告"
			tmp3 = 8
	End Select
%>
<br>
<Script>
function findobj(n, d) {
	var p, i, x;
	if(!d) d = document;
	if((p = n.indexOf("?"))>0 && parent.frames.length) {
		d = parent.frames[n.substring(p + 1)].document;
		n = n.substring(0, p);
	}
	if(x != d[n] && d.all) x = d.all[n];
	for(i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
	for(i = 0; !x && d.layers && i < d.layers.length; i++) x = findobj(n, d.layers[i].document);
	if(!x && document.getElementById) x = document.getElementById(n);
	return x;
}
</Script>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>展现方式: <%=tmp%>
      </ul>
      <ul>
        <li>价值分析: <%=tmp1%>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" name="settings" action="?action=advaddok">
  <input type="hidden" name="types" value="<%=tmp3%>">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">添加广告 - <%=tmp2%></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>广告标题(必填):</b><br>
        <span class="a3">注意: 广告标题只为识别辨认不同广告条目之用，并不在广告中显示</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="advtitle" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>广告投放范围(必选):</b><br>
        <span class="a3">设置本广告投放的页面或论坛范围，可以按住 CTRL 多选，选择“全部”为不限制选择广告投放的范围</span></td>
      <td bgcolor="#FFFFFF"><select name="advtargets" size="10" multiple="multiple">
          <option value="all" selected="selected">&nbsp;> 全部</option>
          <% BBsList(0) %>
        </select></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>生效时间:</b><br>
        <span class="a3">设置广告广告结束的时间，格式 yyyy-mm-dd，留空为不限制结束时间</span></td>
      <td bgcolor="#FFFFFF"> <input type="text" size="30" name="startime" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>到期时间:</b><br>
        <span class="a3">设置广告广告结束的时间，格式 yyyy-mm-dd，留空为不限制结束时间</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="stoptime" value="">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>选择广告模式:</b><br>
        <span class="a3">请选择所需的广告展现方式，可以方便的插入各类广告。</span></td>
      <td bgcolor="#FFFFFF"><select name="advnewstyle" onChange="var styles;var key;styles=new Array('code','text','image','flash'); for(key in styles) {var obj=findobj('style_'+styles[key]); obj.style.display=styles[key]==this.options[this.selectedIndex].value?'':'none';}">
          <option value="code"> 代码</option>
          <option value="text"> 文字</option>
          <option value="image"> 图片</option>
          <option value="flash"> Flash</option>
        </select></td>
    </tr>
  </table>
  <div id="style_code" style=""><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">Html 代码</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" valign="top"><b>广告 html 代码:</b><br>
          <span class="a3">请直接输入需要展现的广告的 html 代码</span></td>
        <td bgcolor="#FFFFFF"><textarea rows="5" name="advcode" cols="40" style="height:70;overflow-y:visible;"></textarea></td>
      </tr>
    </table>
  </div>
  <div id="style_text" style="display: none"><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">文字广告</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>文字内容(必填):</b><br>
          <span class="a3">请输入文字广告的显示内容</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="texttitle" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>文字链接(必填):</b><br>
          <span class="a3">请输入文字广告指向的 URL 链接地址,请以http://开头.</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="textlink" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>文字大小(选填):</b><br>
          <span class="a3">请输入文字广告的内容显示字体，可使用 pt、px、em 为单位</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="textsize" value="">
        </td>
      </tr>
    </table>
  </div>
  <div id="style_image" style="display:none"><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">图片广告</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>图片地址(必填):</b><br>
          <span class="a3">请输入图片广告的图片调用地址</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imageurl" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>图片链接(必填):</b><br>
          <span class="a3">请输入图片广告指向的 URL 链接地址</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imagelink" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>图片宽度(选填):</b><br>
          <span class="a3">请输入图片广告的宽度，单位为像素</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imagewidth" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>图片高度(选填):</b><br>
          <span class="a3">请输入图片广告的高度，单位为像素</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imageheight" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>图片替换文字(选填):</b><br>
          <span class="a3">请输入图片广告的鼠标悬停文字信息</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="imagealt" value="">
        </td>
      </tr>
    </table>
  </div>
  <div id="style_flash" style="display: none"><br>
    <br>
    <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
      <tr class="a1">
        <td colspan="2">Flash 广告</td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>Flash 地址(必填):</b><br>
          <span class="a3">请输入 Flash 广告的调用地址</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="flashurl" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>Flash 宽度(必填):</b><br>
          <span class="a3">请输入 Flash 广告的宽度，单位为像素</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="flashwidth" value="">
        </td>
      </tr>
      <tr>
        <td width="60%" bgcolor="#F8F8F8" ><b>Flash 高度(必填):</b><br>
          <span class="a3">请输入 Flash 广告的高度，单位为像素</span></td>
        <td bgcolor="#FFFFFF"><input type="text" size="30" name="flashheight" value="">
        </td>
      </tr>
    </table>
  </div>
  <br>
  <br>
  <center>
    <input type="submit" name="advsubmit" value="提 交">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub advok
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"AdvList] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"AdvList] set Dois="&Cid(Request.Form("availablenew"&i+1))&",Sorts="&CID(Request.Form("displayordernew"&i+1))&",Titles='"&Replace(Request.Form("titlenew"&i+1),"'","")&"' Where ID="&newid(i))
		Next
	End if
	Cache.DelCache("ForumAdvsLoad")
	team.SaveLog ("广告设置完成")
	SuccessMsg " 广告设置完成  ，请等待系统自动返回到 <a href=Admin_Change.asp?action=adv>广告设置</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=adv>。 " 
End Sub

Sub adv %>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>广告的类型决定广告所在的位置。</li>
      </ul></td>
  </tr>
</table>
<BR>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr>
    <td colspan="2" class="a1">添加广告</td>
  </tr>
  <tr>
    <td colspan="2" class="a4">
	  <input type="button" value="头部横幅广告" onClick="window.location='?action=advadd&type=headerbanner';">
      &nbsp;
      <input type="button" value="尾部横幅广告" onClick="window.location='?action=advadd&type=footerbanner';">
      &nbsp;
      <input type="button" value="页内文字广告" onClick="window.location='?action=advadd&type=text';">
      &nbsp;
      <input type="button" value="帖内广告" onClick="window.location='?action=advadd&type=thread';">
      &nbsp;
      <input type="button" value="漂浮广告" onClick="window.location='?action=advadd&type=float';">
      &nbsp;
      <input type="button" value="对联广告" onClick="window.location='?action=advadd&type=couplebanner';">
	  &nbsp;
      <input type="button" value="公告位广告" onClick="window.location='?action=advadd&type=affbanner';">
	  &nbsp;
	  <input type="button" value="主题帖子位广告" onClick="window.location='?action=advadd&type=threadleft';"></td>
	  
	  </td>
  </tr>
</table>
<BR>
<form method="post" action="?action=advok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr align="center" class="a1">
      <td width="48"><input type="checkbox" name="chkall" class="a1" onClick="checkall(this.form,'delete')">
        删?</td>
      <td width="5%">可用</td>
      <td width="8%">显示顺序</td>
      <td width="15%">标题</td>
      <td width="20%">类型</td>
      <td width="15%">起始时间</td>
      <td width="15%">终止时间</td>
      <td width="15%">投放范围</td>
      <td width="6%">编辑</td>
    </tr>
    <%
	dim Rs,i,tmp
	i = 0
	Set Rs=team.execute("Select ID,Dois,Sorts,Titles,Types,StarTime,StopTime,Boards From ["&Isforum&"AdvList] Order By Sorts Desc")
	Do While Not Rs.Eof
		i = i+1
		Select Case RS(4)
			Case 1
				tmp = "头部横幅广告"
			Case 2
				tmp = "尾部横幅广告"
			Case 3
				tmp = "页内文字广告"
			Case 4
				tmp = "帖内广告"
			Case 5
				tmp = "漂浮广告"
			Case 6
				tmp = "对联广告"
			Case 7
				tmp = "公告栏广告"
			Case 8
				tmp = "主题帖子栏广告"
		End Select				
	%>
    <tr align="center" class="a4">
      <input type="hidden" name="newid" value="<%=RS(0)%>">
      <td><input type="checkbox" name="deleteid" value="<%=RS(0)%>"></td>
      <td><input type="checkbox" name="availablenew<%=i%>" value="1" <%If Rs(1)=1 Then%>checked<%End if%>></td>
      <td><input type="text" size="2" name="displayordernew<%=i%>" value="<%=RS(2)%>"></td>
      <td><input type="text" size="15" name="titlenew<%=i%>" value="<%=RS(3)%>"></td>
      <td><%=tmp%></td>
      <td><%=IIF(RS(5)&""="","无限制",RS(5))%></td>
      <td><%=IIF(RS(6)&""="","无限制",RS(6))%></td>
      <td>
	  <%if Rs(7) = "all" Then 
			Echo "全部" 
		Else 
			Echo RS(7) 
		End if %></td>
      <td><a href="?action=advedit&uid=<%=RS(0)%>">[详情]</a></td>
    </tr>
    <%   Rs.Movenext
	Loop
	Rs.Close:Set RS=Nothing
	%>
  </table>
  <br>
  <center>
    <input type="submit" name="forumlinksubmit" value="提 交">
  </center>
</form>
<%
End Sub

Sub forumlinksok
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"Link] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"Link] set Name='"&Replace(Request.Form("name"&i+1),"'","")&"',Url='"&Replace(Request.Form("url"&i+1),"'","")&"',Intro='"&Replace(Request.Form("note"&i+1),"'","")&"',SetTops="&Cid(Request.Form("displayorder"&i+1))&",logo='"&Replace(Request.Form("logo"&i+1),"'","")&"' Where ID="&newid(i))
		Next
		If Request.Form("newname")<>"" and Request.Form("newurl")<>"" Then
			team.execute("insert into ["&Isforum&"Link] (Name,Url,Intro,SetTops,logo) values ('"&Replace(Request.Form("newname"),"'","")&"','"&Replace(Request.Form("newurl"),"'","")&"','"&Replace(Request.Form("newnote"),"'","")&"',"&CID(Request.Form("newdisplayorder"))&",'"&Replace(Request.Form("newlogo"),"'","")&"') ")
		End if
	End If
	Cache.DelCache("Superlink")
	team.SaveLog ("友情链接设置")
	SuccessMsg " 友情链接设置完成 ，请等待系统自动返回到 <a href=Admin_Change.asp?action=forumlinks>友情链接</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=forumlinks>。 "
End Sub

Sub forumlinks %>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>如果您不想在首页显示联盟论坛，请把已有各项删除即可。</li>
      </ul>
      <ul>
        <li>未填写文字说明的项目将以紧凑型显示。</li>
      </ul>
      <ul>
        <li>未填写logo 地址的项目将以文字排列显示。</li>
      </ul>
      <ul>
        <li>论坛 URL请以 http:// 开始，不然将出现链接无法访问的情况。</li>
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=forumlinksok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="6">友情链接编辑</td>
    </tr>
    <tr align="center" class="a3">
      <td><input type="checkbox" name="chkall" onClick="checkall(this.form)">
        删?</td>
      <td>显示顺序</td>
      <td>论坛名称</td>
      <td>论坛 URL</td>
      <td>文字说明</td>
      <td>logo 地址(可选)</td>
    </tr>
    <%Dim Rs,i
	i=0
	Set Rs=team.execute("Select ID,Name,Url,Intro,SetTops,logo From ["&Isforum&"Link] Order By SetTops Desc")
	Do While Not Rs.Eof
		i= i+1
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td bgcolor="#F8F8F8"><Input Name="newid" type="hidden" value="<%=RS(0)%>">
        <input type="checkbox" name="deleteid" value="<%=RS(0)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="3" name="displayorder<%=i%>" value="<%=RS(4)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="15" name="name<%=i%>" value="<%=RS(1)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="15" name="url<%=i%>" value="<%=RS(2)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="15" name="note<%=i%>" value="<%=RS(3)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="15" name="logo<%=i%>" value="<%=RS(5)%>"></td>
    </tr>
    <%
		Rs.Movenext
	Loop
	Rs.close:Set Rs=Nothing
	%>
    <tr>
      <td colspan="6" class="a4" height="5"></td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td>新增:</td>
      <td><input type="text" size="3"	name="newdisplayorder"></td>
      <td><input type="text" size="15" name="newname"></td>
      <td><input type="text" size="15" name="newurl"></td>
      <td><input type="text" size="15" name="newnote"></td>
      <td><input type="text" size="15" name="newlogo"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="forumlinksubmit" value="提 交">
  </center>
</form>
<%
End Sub


Sub newsannouncements
	Dim newsubject,newcss,newendtime,newmessage
	Newsubject = HtmlEncode(Trim(Request.Form("newsubject")))
	Newmessage = team.checkStr(Trim(Request.Form("newmessage")))
	If Newsubject &""="" Then 
		SuccessMsg "公告标题不能为空。"
	ElseIf Newmessage &""="" Then 
		SuccessMsg "公告内容不能为空。"	
	Else
		If Trim(Request.Form("newendtime"))<>"" Then
			If Not Isdate(Trim(Request.Form("newendtime"))) Then
				SuccessMsg "过期时间的格式不正确，请输入适当的日期格式。"
			End If
		End if
		If Request("edit") = 1 Then
			team.execute(" Update ["&Isforum&"Affiche] Set Affichetitle='"&Newsubject&"',Affichecontent='"&Newmessage&"',Afficheman='"&TK_UserName&"',Afficheinfo='"&Replace(Trim(Request.Form("newcss")),"'","")&"',Lifetime='"&Trim(Request.Form("newendtime"))&"',Affichetime="&SqlNowString&" Where ID="&UID)
			Cache.DelCache("BBsAffiche")
			SuccessMsg "公告编辑完成，请等待系统自动返回到 <a href=Admin_Change.asp?action=announcements>论坛公告</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=announcements>。"	
		Else
			team.execute("insert into ["&Isforum&"Affiche] (Affichetitle,Affichecontent,Afficheman,Afficheinfo,Lifetime,Affichetime) values ('"&Newsubject&"','"&Newmessage&"','"&TK_UserName&"','"&Replace(Trim(Request.Form("newcss")),"'","")&"','"&Trim(Request.Form("newendtime"))&"',"&SqlNowString&") ")
			Cache.DelCache("BBsAffiche")
			SuccessMsg "新的公告发布完成，请等待系统自动返回到 <a href=Admin_Change.asp?action=announcements>论坛公告</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=announcements>。"	
		End If
	End If
	team.SaveLog ("公告设置")
End Sub

Sub announcementsok
	Dim ho
	If request.form("deleteid") = "" Then
		SuccessMsg " 请选定需要删除的公告 "
	Else
		for each ho in request.form("deleteid")
			team.execute("Delete from ["&Isforum&"Affiche] Where ID="&ho)
		next
	End If
	Cache.DelCache("BBsAffiche")
	team.SaveLog ("公告删除")
	SuccessMsg " 公告删除完成，请等待系统自动返回到 <a href=Admin_Change.asp?action=announcements>论坛公告</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Change.asp?action=announcements>。"
End sub

Sub  announcements
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>点击公告的标题，即可对公告进行编辑。
        <li>CSS效果 ：font-weight: bold; 粗体文字。color: #FF0000; 文字颜色 。
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=announcementsok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="7">论坛公告编辑</td>
    </tr>
    <tr align="center" class="a3">
      <td width="48"><input type="checkbox" name="chkall" onClick="checkall(this.form)">
        删?</td>
      <td>作者</td>
      <td>标题</td>
      <td>内容</td>
      <td>起始时间</td>
      <td>终止时间</td>
    </tr>
    <%
	Dim Rs
	Set Rs=team.execute("Select ID,Affichetitle,Affichecontent,Afficheman,Afficheinfo,Lifetime,Affichetime From ["&Isforum&"Affiche] order By Id Asc")
	Do While Not Rs.Eof
	%>
    <tr align="center">
      <td bgcolor="#F8F8F8"><input type="checkbox" name="deleteid" value="<%=RS(0)%>"></td>
      <td bgcolor="#FFFFFF"><a href="./Profile.asp?username=<%=RS(3)%>" target="_blank"><%=RS(3)%></a></td>
      <td bgcolor="#F8F8F8"><a href="?action=announcements&uid=<%=rs(0)%>&edit=1" title="点击编辑此公告"><span Style="<%=Rs(4)%>"><%=RS(1)%></span></a></td>
      <td bgcolor="#FFFFFF"><a href="?action=announcements&uid=<%=rs(0)%>&edit=1" title="点击编辑此公告"><%=CutStr(RS(2),15)%></a></td>
      <td bgcolor="#F8F8F8"><%=RS(6)%></td>
      <td bgcolor="#FFFFFF"><% If RS(5)&"" = "" Then:Echo "无限制": Else:Echo RS(4):End If%></td>
    </tr>
    <%	Rs.MoveNext
	Loop
	Rs.close:set Rs=nothing%>
  </table>
  <br>
  <center>
    <input type="submit" name="announcesubmit" value="提 交">
  </center>
</form>
<br>
<%
If request("edit")=1 Then
	Dim Rs1
	If UID="" Or Not IsNumeric(UID) Then
		SuccessMsg " 参数错误。"
	Else
		Set Rs1=team.execute("Select ID,Affichetitle,Affichecontent,Afficheman,Afficheinfo,Lifetime,Affichetime From ["&Isforum&"Affiche] Where ID="& UID)
		If Rs1.eof Then 
			SuccessMsg " 参数错误。"
		Else
			Echo " <form method=""post"" action=""?action=newsannouncements&edit=1&uid="&UID&"""> "
		End If
	End if
Else
	Echo "<form method=""post"" action=""?action=newsannouncements"">"
End If
If request("edit")=1 Then
%>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td colspan="2">编辑论坛公告</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>标题:</b></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newsubject" Value = "<%=RS1(1)%>"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>标题颜色:</b><BR>
      <span class="a3">支持使用CSS效果</span></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newcss" Value = "<%=RS1(4)%>"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>过期时间:</b><br>
      格式: yyyy-mm-dd</td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newendtime" Value = "<%=RS1(5)%>">
      留空为不限制</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>内容:</b><br>
      公告内容支持UBB代码<BR>
      UBB代码使用请查看<a href="../Help.asp?page=mise#1"> <B>UBB指南</B> </a>
    <td width="60%" bgcolor="#FFFFFF"><textarea name="newmessage" cols="60" rows="10"><%=Server.htmlEncode(RS1(2))%></textarea></td>
  </tr>
</table>
<%Else%>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td colspan="2">添加论坛公告</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>标题:</b></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newsubject"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>标题颜色:</b><BR>
      <span class="a3">支持使用CSS效果</span></td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newcss"></td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8"><b>过期时间:</b><br>
      格式: yyyy-mm-dd</td>
    <td width="60%" bgcolor="#FFFFFF"><input type="text" size="45" name="newendtime">
      留空为不限制</td>
  </tr>
  <tr>
    <td width="40%" bgcolor="#F8F8F8" valign="top"><b>内容:</b><br>
      公告内容支持UBB代码<BR>
      UBB代码使用请查看<a href="../Help.asp?page=mise#1"> <B>UBB指南</B> </a>
    <td width="60%" bgcolor="#FFFFFF"><textarea name="newmessage" cols="60" rows="10"></textarea></td>
  </tr>
</table>
<%End if%>
<br>
<center>
<input type="submit" name="addsubmit" value="提 交">
</form>
<br>
<br>
<%
End Sub

Sub medalsok
	Dim MedalName,MedalSet,Medalimg
	Dim ho,newid,i
	for each ho in request.form("deleteid")
		team.execute("Delete from ["&Isforum&"Medals] Where ID="&ho)
	next
	If Request.form("deleteid")="" Then
		newid=Split(Replace(Request.Form("newid")," ",""),",")
		For i=0 To Ubound(newid)
			team.Execute("Update ["&Isforum&"Medals] set MedalName='"&Request.Form("MedalName"&i+1)&"',Medalimg='"&Request.Form("Medalimg"&i+1)&"',MedalSet="&Cid(Request.Form("MedalSet"&i+1))&" Where ID="&newid(i))
		Next
		If Request.Form("newname")<>"" and Request.Form("newimage")<>"" Then
			team.execute("insert into ["&Isforum&"Medals] (MedalName,Medalimg,MedalSet) values ('"&Request.Form("newname")&"','"&Request.Form("newimage")&"',"&CID(Request.Form("availablenew"))&" ) ")
		End if
	End If
	team.SaveLog ("勋章设置")
	SuccessMsg " 勋章设置完成 。 "
End Sub

Sub medals	
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a4">
    <td><br>
      <ul>
        <li>本功能用于设置可以颁发给用户的勋章信息，勋章图片中请填写图片文件名，并将相应图片文件上传到 ../images/plus 目录中。
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?action=medalsok">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="5">勋章编辑</td>
    </tr>
    <tr align="center" class="a3">
      <td><input type="checkbox" name="chkall" class="a4" onClick="checkall(this.form, 'delete')">
        删?</td>
      <td>名称</td>
      <td>图片地址</td>
      <td>勋章图片</td>
      <td>可用</td>
    </tr>
    <%
	Dim Rs,i
	i=0
	Set Rs=team.execute("Select ID,MedalName,Medalimg,MedalSet From ["&Isforum&"Medals] Order By ID asc")
	Do While Not Rs.Eof
		i = i+1
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td bgcolor="#F8F8F8" width="48"><input type="checkbox" name="deleteid" value="<%=RS(0)%>">
        <Input Name="newid" type="hidden" value="<%=RS(0)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="MedalName<%=i%>" value="<%=RS(1)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="30" name="Medalimg<%=i%>" value="<%=RS(2)%>"></td>
      <td bgcolor="#FFFFFF"><img src="../images/plus/<%=RS(2)%>" align="absmiddle"> </td>
      <td bgcolor="#F8F8F8"><input type="checkbox" name="MedalSet<%=i%>" value="1" <%If Rs(3)=1 then%>checked<%end if%>></td>
    </tr>
    <%
		Rs.MoveNext
	Loop
	Rs.close:Set Rs=Nothing
	%>
    <tr>
      <td colspan="5" class="a4" height="2"></td>
    </tr>
    <tr bgcolor="#F8F8F8" align="center">
      <td>新增:</td>
      <td><input type="text" size="30" name="newname"></td>
      <td><input type="text" size="30" name="newimage"></td>
      <td>&nbsp;</td>
      <td><input type="checkbox" name="availablenew" value="1"></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="medalsubmit" value="提 交">
  </center>
</form>
<%
End Sub
Sub BBsList(V)
	Dim SQL,ii,RS,i
	Set Rs=Team.Execute("Select ID,BBSname,Followid From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		If RS(2)=0 Then 
			Echo "<optgroup label="""&Rs(1)&""">"
		Else
			Echo "<option value="&RS(0)&" " 
			If Isarray(boards) Then
				for i=0 to ubound(boards)
					If RS(0) = int(boards(i)) Then Echo "selected=""selected"" " 
				next
			end if
			Echo " >"&String(ii,"　") & RS(1)&"</option>"
		End if
		ii=ii+1
		BBsList RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub

Sub Menus()
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Name,url,followid,SortNum,Newtype From "&IsForum&"Menu Where Followid=0 Order By SortNum")
	If Rs.Eof Then
		Echo "<BR><ul><center> 目前没有添加任何菜单 </center></ul> "
	End if
	Do While Not RS.Eof
		Echo "<tr class=""a4"" align=""center""><td width=""10""> <Input Name=UID value="&RS(0)&" type=hidden> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&"></td><td width=""100""> 排序：  <input type=text name=SortNum Value="&RS(4)&" Size=""1""> </td><td width=""50%"" align=""left""> ┝ <a target=_blank href=../"&RS(2)&"><b>"&RS(1)&"</b></a> </td><td> "
		If Rs(5)=1 Then	
			Echo "前台菜单"
		Else
			Echo "后台菜单"
		End if
		Echo " </td><td> <a href=""?action=menuadd&fid="&RS(0)&"&Mid="&Rs(5)&""" title=""添加本分类或下级菜单"">[添加]</a> <a href=""?action=menuadd&uid="&RS(0)&"&edit=1&Mid="&Rs(5)&""" title=""编辑本菜单设置"">[编辑]</a> </td></tr>"
		Call Menus_1(Rs(0))
		Echo " "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub

Sub Menus_1(a)
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Name,url,followid,SortNum,Newtype From "&IsForum&"Menu Where Followid="&a&" Order By SortNum")
	Do While Not RS.Eof
		Echo "<tr class=""a4"" align=""center""><td width=""10""> <Input Name=UID value="&RS(0)&" type=hidden> <input type=""checkbox"" name=""deleteid"" value="&RS(0)&"></td><td width=""100""> 排序：  <input type=text name=SortNum Value="&RS(4)&" Size=""1""> </td><td width=""50%"" align=""left"">　　 ┕<a target=_blank href=../"&RS(2)&"><b>"&RS(1)&"</b></a> </td><td> "
		If Rs(5)=1 Then	
			Echo "前台菜单"
		Else
			Echo "后台菜单"
		End if
		Echo " </td><td> <a href=""?action=menuadd&uid="&RS(0)&"&edit=1&Mid="&Rs(5)&""" title=""编辑本菜单设置"">[编辑]</a> </td></tr>"
		Call Menus_1(Rs(0))
		Echo " "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub

footer()
%>
