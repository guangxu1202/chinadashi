<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim ii,ID
Dim Admin_Class
Call Master_Us()
Header()
ii=0:ID=Request("id")
Admin_Class=",2,"
Call Master_Se()
team.SaveLog ("论坛设置 [包括：编辑版块 ] 修改")
Select Case Request("Action")
	Case "Manages"
		Call Manages	'管理板块
	Case "FindForum"
		ID= Request.Form("ForumID")
		Call Manages	'查找版块
	Case "Forumadd"
		Call Forumadd	'编辑板块
	Case "ForumSort"	'版块排序
			Dim Myid,MySortNum
			Myid=Split(Request.Form("UID"),",")
			MySortNum=Split(Request.Form("SortNum"),",")
			For U=0 To Ubound(Myid)
				team.Execute("Update "&IsForum&"Bbsconfig set SortNum="&MySortNum(U)&" where ID="&Myid(U))
			Next
			Cache.DelCache("BoardLists")
			SuccessMsg("排序完成!")	
	Case "ForumAddok"	
		Dim fup
		fup=ReQuest.Form("fup")
		If Request.Form("newforum")="" Then Error2 "板块名称不能为空!"
		Select Case Request("add")
			Case "Forum_0"
				team.Execute("insert into "&IsForum&"Bbsconfig(Followid,bbsname,Board_Last,Board_Setting,today,toltopic,tolrestore,hide,Board_Model,SortNum) values (0,'"&Replace(Request.Form("newforum"),"'","")&"','暂无帖子$@$ - $@$"&Now&"','0$$$0$$$0$$$0$$$1$$$1$$$1$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0',0,0,0,0,0,0) ")
				Cache.DelCache("BoardLists")
				SuccessMsg("一级分类 ["&Request.Form("newforum")&"] 添加成功<BR><a href=Admin_Forum.asp>请转入主菜单进行详细参数的设置</a>!")
			Case "Forum_1"
				team.Execute("insert into "&IsForum&"Bbsconfig(Followid,bbsname,Board_Last,Board_Setting,today,toltopic,tolrestore,hide,Board_Model,SortNum) values ("&Request.Form("fup")&",'"&Replace(Request.Form("newforum"),"'","")&"','暂无帖子$@$ - $@$"&Now&"','0$$$0$$$0$$$0$$$1$$$0$$$1$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0|0,0,0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0$$$0',0,0,0,0,0,0) ")
				Cache.DelCache("BoardLists")
				SuccessMsg("分类 ["&Request.Form("newforum")&"] 添加成功<BR><a href=Admin_Forum.asp>请转入主菜单进行详细参数的设置</a>!")
		End Select
	Case "Forumeditok"
		Dim My_Board_Setting,U,My_ExtCredit,ExtCredits
		If Request.Form("BbsName")="" Then Error2 "分类名称不能为空!"
		My_Board_Setting=""
		My_ExtCredit = ""
		ExtCredits= Split(team.Club_Class(21),"|")
		For U=0 to Ubound(ExtCredits)
			If U=0 Then
				My_ExtCredit=Replace(Request.Form("ExtCredits0_0"),",","")&","&Replace(Request.Form("ExtCredits0_1"),",","")&","&Replace(Request.Form("ExtCredits0_2"),",","")
			Else				
				My_ExtCredit=My_ExtCredit & "|"&Replace(Request.Form("ExtCredits"&U&"_0"),",","")&","&Replace(Request.Form("ExtCredits"&U&"_1"),",","")&","&Replace(Request.Form("ExtCredits"&U&"_2"),",","")
			End If
		Next
		For U=0 to 19
			If U=0 Then
				My_Board_Setting=Replace(Request.Form("Board_Setting(0)"),"$$$","")
			ElseIf U=14 Then				
				My_Board_Setting=My_Board_Setting & "$$$"&My_ExtCredit
			Else
				My_Board_Setting=My_Board_Setting & "$$$"&Replace(Request.Form("Board_Setting("&U&")"),"$$$","")
			End If
		Next
		team.Execute("Update "&IsForum&"Bbsconfig Set bbsname='"&HTMLEncode(Request.Form("BbsName"))&"',Readme='"&Replace(Trim(Request.Form("Readme")),"'","''")&"',Icon='"&Replace(Trim(Request.Form("Icon")),"'","''")&"',Board_Key='"&Replace(Trim(Request.Form("Board_Key")),"'","''")&"',Hide="&HTMLEncode(Trim(Request.Form("Hide")))&",Pass='"&HTMLEncode(Trim(Request.Form("Pass")))&"',Followid="&Cid(Request.Form("fupnew"))&",Board_URL='"&HtmlEncode(Trim(Request.Form("Board_URL")))&"',Board_Setting='"&My_Board_Setting&"',Lookperm='"&Replace(Request.Form("lookperm")," ","")&",',Postperm='"&Replace(Request.Form("postperm")," ","")&",',Downperm='"&Replace(Request.Form("downperm")," ","")&",',upperm='"&Replace(Request.Form("upperm")," ","")&",' Where ID="&ID)
		Cache.DelCache("ForumsBoards_"&ID)
		Cache.DelCache("ThreadBoards_"&ID)
		Cache.DelCache("SaveThreadBoards_"&ID)
		Cache.DelCache("Boards_"&ID)
		Cache.DelCache("BoardLists")
		SuccessMsg("分类 ["&Request.Form("BbsName")&"] 编辑成功，请等待系统自动返回到 <a href=Admin_Forum.asp>编辑版块</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Forum.asp>。")
	Case "SetModerators"
		Call SetModerators
	Case "ModelSet_0"
		If ID="" or (Not isNumeric(ID)) Then 
			SuccessMsg " ID参数错误! "
		Else
			team.execute("Update ["&Isforum&"BbsConfig] Set Board_Model = 1 Where Id="&ID&" or Followid="&ID)
			Cache.DelCache("BoardLists")
			SuccessMsg "已经将本版块的排列方式修改为简洁模式，请等待系统自动返回到 <a href=Admin_Forum.asp>编辑版块</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Forum.asp>。"
		End If
	Case "ModelSet_1"
		If ID="" or (Not isNumeric(ID)) Then 
			SuccessMsg " ID参数错误! "
		Else
			team.execute("Update ["&Isforum&"BbsConfig] Set Board_Model = 0 Where Id="&ID&" or Followid="&ID)
			Cache.DelCache("BoardLists")
			SuccessMsg "已经将本版块的排列方式修改为标准模式，请等待系统自动返回到 <a href=Admin_Forum.asp>编辑版块</a> 页面 。<meta http-equiv=refresh content=3;url=Admin_Forum.asp>。"
		End If
	Case "DelForum"
		Call DelForums
	Case "ServerDelForum"
		Dim Rs
		If ID="" or (Not isNumeric(ID)) Then 
			SuccessMsg " ID参数错误! "
		Else
			team.Execute("Delete From "&IsForum&"Bbsconfig Where ID="&ID)
			Set Rs = team.execute("Select ID,ReList From ["&IsForum&"Forum] Where forumid=" & ID)
			Do While Not Rs.Eof
				team.Execute("Delete From ["&IsForum & RS(1) &"] Where topicid="& RS(0) )
				Rs.MoveNext
			Loop
			Rs.close:Set Rs=Nothing
			team.Execute("Delete From ["&IsForum&"Forum] Where forumid="&ID)
			Cache.DelCache("BoardLists")
			Cache.DelCache("ForumsBoards_"&ID)
			Cache.DelCache("Boards_"&ID)
			SuccessMsg("删除论坛成功<BR><a href=Admin_Forum.asp>请转入主菜单进行其他设置</a>或等待3秒钟后，系统自动转入主菜单界面。<meta http-equiv=refresh content=3;url=Admin_Forum.asp>")
		End if
	Case Else
		Call Main()
End Select

Sub Main()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2" >
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr class="a3">
   <td><BR><ul>
        <li>您需要首先提交一级分类。</li>
      </ul>
      <ul>
        <li>每级分类后面的管理功能后面带有添加下级版面功能，但分类版块推荐不要超过三级。</li>
      </ul>
      <ul>
        <li>您可以对在“显示顺序”里面对论坛进行排序，每个级别的排序从 0 开始。</li>
      </ul></td>
  </tr>
</table><BR>
<form method="post" action="?Action=FindForum">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="5">查找论坛</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="15%">名称:</td>
      <td bgcolor="#FFFFFF" width="70%"><input type="text" name="ForumID" value="需查找的论坛ID" size="40" OnFocus="this.value = ''"></td>
      <td bgcolor="#F8F8F8" width="15%"><input type="submit" name="forumsubmit" value="提 交"></td>
    </tr>
  </table>
</form>
<form method="Post" action="?Action=ForumAddok&add=Forum_0">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="3">添加新一级分类</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="15%">名称:</td>
      <td bgcolor="#FFFFFF" width="70%"><input type="text" name="newforum" value="新分类名称" size="40"></td>
      <td bgcolor="#F8F8F8" width="15%"><input type="submit" name="catsubmit" value="提 交"></td>
    </tr>
  </table>
</form>
<form method="post" action="?Action=ForumSort">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td>编辑论坛</td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" width="100%" valign=top height=200>
	  <table cellspacing="1" cellpadding="1" width="98%" align="center" class="a2">
	  <tr class="a4"><td>
			<% ForumList(0) %>
			</td></tr>
        </table></td>
    </tr>
  </table><br><center><input type="submit" name="detailsubmit" value="提 交"></form><br>
<%
End Sub

Sub Manages
		Dim Board_Setting,RS
		Dim B_Lookperm,B_Postperm,B_DownPerm,B_Upperm
		If ID="" or (Not isNumeric(ID)) Then SuccessMsg " ID参数只能是数字! "
		Set Rs=team.Execute("Select bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Readme,Board_Key,Board_URL,Lookperm,Postperm,DownPerm,Upperm From "&IsForum&"Bbsconfig Where ID="&ID)
		If RS.Eof or Rs.Bof Then
			SuccessMsg("ID参数错误!")
		Else
			Board_Setting = Split(RS("Board_Setting"),"$$$")
		%>
	<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
	<form method="post" action="?Action=Forumeditok">
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="detailsubmit" value="submit">
	<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td>技巧提示</td>
    </tr>
    <tr bgcolor="#F8F8F8">
      <td><br>
        <ul>
          以下设置没有继承性，即仅对当前论坛有效，不会对下级子论坛产生影响。
        </ul></td>
    </tr>
  </table>
  <br>
  <br>
  <a name="论坛详细设置 - <%=RS(0)%>"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">论坛详细设置 - <%=RS(0)%></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>显示论坛:</b><br>
        <span class="a3">选择“否”将暂时将论坛隐藏不显示，但论坛内容仍将保留，且用户仍可通过直接提供带有 id 的 URL 访问到此论坛，如果隐藏的是一级版块，那么其所在的下级版块将跟随主版块一起隐藏。</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Hide" value="0" <%If RS("Hide")=0 Then%>checked<%End If%>>
        是
        <input type="radio" name="Hide" value="1" <%If RS("Hide")=1 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>上级论坛:</b><br>
        <span class="a3">本论坛的上级论坛或分类</span></td>
      <td bgcolor="#FFFFFF">
	  <select name="fupnew">
			<option value="0">&nbsp;>>一级论坛</option>
			<% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>风格方案:</b><br>
        <span class="a3">访问者进入本论坛所使用的风格方案</span></td>
      <td bgcolor="#FFFFFF">
	  <select name="Board_Setting(0)">
		<option value="<%=Int(team.Forum_setting(18))%>" SELECTED>采用论坛默认模版</option>
      <%
		Dim RS1,SytyleID
		Set Rs1=team.Execute( "Select StyleName,ID From ["&IsForum&"Style] Order By ID Asc" )
		Do While Not RS1.Eof
			SytyleID = SytyleID &  "<option value="&RS1(1)&"" 
			If Int(Rs1(1)) = Int(Board_Setting(0)) Then SytyleID = SytyleID & " SELECTED"
			SytyleID = SytyleID &">"&RS1(0)&"</option>"
			Rs1.Movenext
		Loop
		RS1.CLOSE:Set RS1=Nothing
		Response.Write SytyleID
	'名称     参数         隐藏 密码  图标  权限     介绍   规则     转向地址
	'bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Readme,Board_Key,Board_URL
%>
        </select></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>论坛转向 URL</b><br>
        <span class="a3">如果设置转向 URL(例如 http://www.team5.cn)，用户点击本分论坛将进入转向中设置的 URL。一旦设定将无法进入论坛页面，请确认是否需要使用此功能，留空为不设置转向 URL，本站以外的URL地址必须加HTTP://</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Board_URL" value="<%=RS("Board_URL")%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>论坛名称:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="BbsName" value="<%=RS("BbsName")%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>论坛图标:</b><br>
        <span class="a3">论坛名称和简介左侧的小图标，可填写相对或绝对地址</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Icon" value="<%=RS("Icon")%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>论坛简介:</b><br>
        <span class="a3">将显示于论坛名称的下面，提供对本论坛的简短描述，支持Ubb代码 </span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="Readme" cols="30" style="height:70;overflow-y:visible;"><%=ReplaceStr(RS("Readme"),"<BR>",VbCrlf)%></textarea></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>本论坛规则:</b><br>
        <span class="a3">显示于主题列表页的当前论坛规则，支持 Ubb 代码，留空为不显示</span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="Board_Key" cols="30" style="height:70;overflow-y:visible;"><%=ReplaceStr(RS("Board_Key"),"<BR>",VbCrlf)%></textarea></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许版主修改本论坛规则:</b><br>
        <span class="a3">设置是否允许超级版主和版主通过系统设置修改本版规则</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(1)" value="0" <%If Board_Setting(1)=0 Then%>checked<%End If%>>
        不允许版主修改 
		<input type="radio" name="Board_Setting(1)" value="1" <%If Board_Setting(1)=1 Then%>checked<%End If%>>允许版主修改 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top"><b>访问权限限制:</b><br>
        <span class="a3">您可以通过此选项区分 <B>浏览论坛许可</B>， 限制该用户组查看帖子内容及帖子标题列表的权限。这样可以让该用户组在查看到帖子标题的情况下，但无法查看更多的帖子详细内容。</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(8)" value="0" <%If Board_Setting(8)=0 Then%>checked<%End If%>>
        全部限制
		<input type="radio" name="Board_Setting(8)" value="1" <%If Board_Setting(8)=1 Then%>checked<%End If%>>只在查看帖子内容时开启限制
		</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" valign="top">
	  <B>开启游客发帖权限:</B> 
	  <BR>如果想让游客发帖，还需要设置 <BR>
	     1. 组权限允许 ，<a href="Admin_Group.asp?Action=Editmodel_1&ID=28#帖子相关">点击设置游客组权限</a>。 <BR>
		 2.  <B><a href="#论坛权限">发新话题许可</a></B> 开放游客发帖权限 <br>
        <span class="a3">  </span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(9)" value="1" <%If CID(Board_Setting(9))=1 Then%>checked<%End If%>>
        是
		<input type="radio" name="Board_Setting(9)" value="0" <%If CID(Board_Setting(9))=0 Then%>checked<%End If%>>否
		</td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="提 交"><br>
  <br>
  <a name="帖子选项"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">帖子选项</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>发帖审核:</b><br>
        <span class="a3">选择“是”将使用户在本版发表的帖子待版主或管理员审查通过后才显示出来，打开此功能后，您可以在用户组中设定哪些组发帖可不经审核，也可以在管理组中设定哪些组可以审核别人的帖子</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(2)" value="0" <%If Board_Setting(2)=0 Then%>checked<%End If%>>
        无<br>
        <input type="radio" name="Board_Setting(2)" value="1" <%If Board_Setting(2)=1 Then%>checked<%End If%>>
        审核新主题<br>
        <input type="radio" name="Board_Setting(2)" value="2" <%If Board_Setting(2)=2 Then%>checked<%End If%>>
        审核新主题和新回复 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>限制回复主题:</b><br>
        <span class="a3">选择“是”将使用户在本版无法发表回复主题。</span></td>
      <td bgcolor="#FFFFFF">
	    <input type="radio" name="Board_Setting(5)" value="1" <%If CID(Board_Setting(5))=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(5)" value="0" <%If CID(Board_Setting(5))=0 Then%>checked<%End If%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许加入文集:</b><br>
        <span class="a3">是否允许用户将在本版发表的主题加入其自己的 文集 中。注意: 一旦主题被加入 文集，其内容会被公开而无论当前论坛被添加什么样的权限设定</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(4)" value="1" <%If Board_Setting(4)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(4)" value="0" <%If Board_Setting(4)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许使用 UBB 代码:</b><br>
        <span class="a3">UBB 代码是一种简化和安全的页面格式代码，可 <a href="../Help.asp?page=mise#1" target="_blank">点击这里查看本论坛提供的UBB代码</a></span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(6)" value="1" <%If Board_Setting(6)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(6)" value="0" <%If Board_Setting(6)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>启用内容干扰码:</b><br>
        <span class="a3">选择“是”将在帖子内容中增加随机的干扰字串，使得访问者无法复制原始内容。注意: 本功能会轻微加重服务器负担</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(7)" value="1" <%If Board_Setting(7)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(7)" value="0" <%If Board_Setting(7)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="提 交"><br>
  <br>
  <a name="积分设置"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">积分设置</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>自定义发主题增加积分:</b><br>
        <span class="a3">设置本论坛是否使用独立的发新主题积分规则。选择“是”，用户在本论坛发主题时，积分将按照如下设置增减，请在下面表格中输入各项积分增减数值；选择“否”，积分将按全论坛默认设定的规则增减</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(12)" value="1" <%If Board_Setting(12)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(12)" value="0" <%If Board_Setting(12)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>自定义发回复增加积分:</b><br>
        <span class="a3">设置本论坛是否使用独立的发新回复积分规则。选择“是”，用户在本论坛发回复时，积分将按照如下设置增减，请在下面表格中输入各项积分增减数值；选择“否”，积分将按全论坛默认设定的规则增减</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(13)" value="1" <%If Board_Setting(13)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(13)" value="0" <%If Board_Setting(13)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>自定义精华贴增加积分:</b><br>
        <span class="a3">设置本论坛是否使用独立的发新回复积分规则。选择“是”，用户帖子被加入精华时，积分将按照如下设置增减，请在下面表格中输入各项积分增减数值；选择“否”，积分将按全论坛默认设定的规则增减</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Board_Setting(3)" value="1" <%If Board_Setting(3)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(3)" value="0" <%If Board_Setting(3)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td colspan="2" bgcolor="#F8F8F8"><table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
          <tr align="center" class="a1">
            <td>积分代号</td>
            <td>积分名称</td>
            <td>发主题(+)</td>
            <td>回复(+)</td>
			<td>加精华(+)</td>
          </tr>
          <% 
	Dim ExtCredits,U,ExtSort,Board_Up,My_ExtSort,MY_ExtCredits,Mydisabled
	ExtCredits= Split(team.Club_Class(21),"|")
	MY_ExtCredits=Split(Board_Setting(14),"|")
	For U=0 to Ubound(ExtCredits)
		ExtSort=Split(ExtCredits(U),",")
		My_ExtSort=Split(MY_ExtCredits(U),",")
		If ExtSort(3)=1 then
			Mydisabled =""
		Else
			Mydisabled = "disabled"
		End If
	%>
          <tr align="center" <%=Mydisabled%>>
            <td bgcolor="#F8F8F8">ExtCredits<%=U+1%></td>
            <td bgcolor="#FFFFFF"><%=ExtSort(0)%></td>
            <td bgcolor="#F8F8F8"><input type="text" size="2" name="ExtCredits<%=U%>_0" value="<%=My_ExtSort(0)%>"></td>
            <td bgcolor="#FFFFFF"><input type="text" size="2" name="ExtCredits<%=U%>_1" value="<%=My_ExtSort(1)%>"></td>
			<td bgcolor="#FFFFFF"><input type="text" size="2" name="ExtCredits<%=U%>_2" value="<%=My_ExtSort(2)%>"></td>
          </tr>
          <%
	Next
	%>
        </table></td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="提 交"><br>
  <br>
  <a name="主题分类"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">主题分类</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>启用主题分类:</b><br>
        <span class="a3">设置是否在本论坛启用主题分类功能，您需要同时设定相应的分类选项，才能启用本功能</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(15)" value="1" <%If Board_Setting(15)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(15)" value="0" <%If Board_Setting(15)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>发帖必须归类:</b><br>
        <span class="a3">如果选择“是”，作者发新主题时，必须选择主题对应的类别才能发表。本功能必须“启用主题分类”后才可使用 </span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(16)" value="1" <%If Board_Setting(16)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(16)" value="0" <%If Board_Setting(16)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许按类别浏览:</b><br>
        <span class="a3">如果选择“是”，用户将可以在本论坛中按照不同的类别浏览主题。注意: 本功能必须“启用主题分类”后才可使用并会加重服务器负担</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(17)" value="1" <%If Board_Setting(17)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(17)" value="0" <%If Board_Setting(17)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>类别前缀:</b><br>
        <span class="a3">设置是否在主题列表中，给已分类的主题前加上类别的显示。注意: 本功能必须“启用主题分类”后才可使用</span></td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Board_Setting(18)" value="1" <%If Board_Setting(18)=1 Then%>checked<%End If%>>
        是
        <input type="radio" name="Board_Setting(18)" value="0" <%If Board_Setting(18)=0 Then%>checked<%End If%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>类别选项:</b><br>
        <span class="a3">请填写期望在本论坛中使用的类别选项，用户发表主题或浏览时，将可按照选中的类别归类或浏览，每个类别一行 。注意: 本功能必须“启用主题分类”后才可使用。</span></td>
      <td bgcolor="#FFFFFF"><textarea rows="5" name="Board_Setting(19)" cols="30" style="height:70;overflow-y:visible;"><%=Board_Setting(19)%></textarea></td>
    </tr>
  </table>
  <br>
  <center><input type="submit" name="detailsubmit" value="提 交"><br>
  <br>
  <a name="论坛权限"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">论坛权限 - 全不选则按照默认设置</td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>访问密码:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Pass" value="<%=RS("Pass")%>"></td>
    </tr>
    <tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>浏览论坛许可:</b><br>
        <span>默认为全部具有浏览论坛帖子权限的用户组<br>
        <input type="checkbox" name="chkall1" onClick="checkall(this.form, 'lookperm', 'chkall1')">
        全选</span></td>
      <td bgcolor="#FFFFFF">
	  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
	  <tr>
		<% 
		Dim Gs,Value,i,m
		Set Gs = team.execute("Select ID,MemberRank,GroupName From "&IsForum&"UserGroup Where Not (ID=7 or ID=6 or ID=5) Order By GroupRank Desc")
		If Gs.Eof or Gs.Bof Then
			SuccessMsg " 用户权限表数据损坏,请手动导入新表! "
		Else
			Value = Gs.GetRows(-1)
		End If
		Gs.Close:Set Gs=Nothing

		'bbsname,Board_Setting,Hide,Pass,Icon,Ismaster,Readme,Board_Key,Board_URL,Lookperm,Postperm,DownPerm,Upperm
		If Instr(RS(9),",")>0 Then B_Lookperm = Split(RS(9),",")
		If Instr(RS(10),",")>0 Then B_Postperm = Split(RS(10),",")
		If Instr(RS(11),",")>0 Then B_DownPerm = Split(RS(11),",")
		If Instr(RS(12),",")>0 Then B_Upperm = Split(RS(12),",")
		If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				response.write "<td><input type=""checkbox"" name=""lookperm"" class=""radio"" value="&Replace(Value(0,i)," ","")&" "
				If Isarray(B_Lookperm) Then
					for m = 0 to Ubound(B_Lookperm)-1
						If Cid(Trim(B_Lookperm(m))) = int(Value(0,i)) Then response.write "checked"
					next
				end if
				response.write "  >"&Value(2,i)&"</td> "
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table>
		</td>
    </tr><tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>发新话题许可:</b><br>
        <span>默认为除游客组以外具有发帖权限的用户组<br>
        <input type="checkbox" name="chkall2" onClick="checkall(this.form, 'postperm', 'chkall2')">
        全选</span></td>
      <td bgcolor="#FFFFFF">
	  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr><%
        If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				Echo "<td><input type=""checkbox"" name=""postperm"" class=""radio"" value="&Value(0,i)&" "
				If Isarray(B_Postperm) Then
					for m = 0 to Ubound(B_Postperm)-1
						If Cid(Trim(B_Postperm(m))) = cid(Value(0,i)) Then Echo "checked"
					next
				end if
				Echo "  >"&Value(2,i)&"</td> "
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table></td>
    </tr>
	<tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>下载/查看附件许可:</b><br>
        <span>默认为全部具有下载/查看附件权限的用户组<br>
        <input type="checkbox" name="chkall4" onClick="checkall(this.form, 'downperm', 'chkall4')">
        全选</span></td>
      <td bgcolor="#FFFFFF">
		  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr><%
        If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				Echo "<td><input type=""checkbox"" name=""downperm"" class=""radio"" value="&Value(0,i)&" "
				If Isarray(B_DownPerm) Then
					for m = 0 to Ubound(B_DownPerm)-1
						If Cid(Trim(B_DownPerm(m))) = Cid(Value(0,i)) Then Echo "checked"
					next
				end if
				Echo "  >"&Value(2,i)&"</td> "				
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table></td>
    </tr>
	<tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
    <tr>
      <td width="15%" bgcolor="#F8F8F8" valign="top"><b>上传附件许可:</b><br>
        <span>默认为除游客以外具有上传附件权限的用户组<br>
        <input type="checkbox" name="chkall5" onClick="checkall(this.form, 'upperm', 'chkall5')">
        全选</span></td>
      <td bgcolor="#FFFFFF">
		  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr><%
        If Isarray(Value) Then
			U=0
			For i=0 To Ubound(Value,2)	
				U = U+1
				Echo "<td><input type=""checkbox"" name=""upperm"" class=""radio"" value="&Value(0,i)&" "
				If Isarray(B_Upperm) Then
					for m = 0 to Ubound(B_Upperm)
						If Cid(Trim(B_Upperm(m))) = CID(Value(0,i)) Then Echo "checked"
					next
				end if
				Echo "  >"&Value(2,i)&"</td> "	
				If U= 4 Then 
					Echo "</tr><tr>"
					U=0
				End If
			Next
		End If
		%>
        </table></td>
    </tr>
	<tr>
      <td colspan="2" class="a4" height="2"></td>
    </tr>
  </table>
  <br>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="提 交">
</form>
<br>
<%	
	End If
	Rs.Close:Set Rs=Nothing
End Sub


Sub Forumadd%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr bgcolor="#F8F8F8">
    <td><br>
      <ul>
        你只有添加了版块以后才可以对版块进行详细的设置。
      </ul></td>
  </tr>
</table>
<br>
<form method="post" action="?Action=ForumAddok&add=Forum_1">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan=5>添加子论坛</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="15%">名称:</td>
      <td bgcolor="#FFFFFF" width="28%"><input type="text" name="newforum" value="子论坛名称" size="20"></td>
      <td bgcolor="#F8F8F8" width="15%">上级论坛:</td>
      <td bgcolor="#FFFFFF" width="27%"><select name="fup">
		  <option value="0">&nbsp;>>一级论坛</option>
          <% ForumList_Sel(0) %>
        </select></td>
      <td bgcolor="#F8F8F8" width="15%"><input type="submit" name="forumsubmit" value="提 交"></td>
    </tr>
  </table>
</form>
<br>
<%
End Sub
Sub SetModerators
	Dim Rs3,ho
	Dim newmoderator,newdisplayorder,Rs1
	If Request("UpModers")=1 Then
		Newmoderator = HTMLEncode(Request.Form("newmoderator"))
		Newdisplayorder = HTMLEncode(Request.Form("newdisplayorder"))
		for each ho in request.form("isdelete")
			Team.execute("Delete from ["&isforum&"Moderators] Where id="&ho)
		next
		If Request.form("isdelete")="" Then
			If Newmoderator="" or Newdisplayorder="" Then Error2("参数不能为空!")
			If team.execute("Select * from ["&isforum&"User] where UserName='"&Newmoderator&"' ").Eof Then
				Error2("指定用户不存在，请返回。")
			Else
				Set Rs3 = team.execute("Select UserGroupID from ["&isforum&"User] where UserName='"&Newmoderator&"' ")
				If Not RS3.Eof Then
					If Not (CID(Rs3(0)) = 1 Or CID(Rs3(0)) = 2) Then
						If team.execute("Select ManageUser from "&isforum&"Moderators where ManageUser='"&Newmoderator&"' and BoardID="&ID).Eof Then
							team.execute("insert into "&isforum&"Moderators (BoardID,ManageUser,Issort) values ("&ID&",'"&Newmoderator&"',"&Newdisplayorder&") ")
							team.execute("Update ["&isforum&"User] Set UserGroupID=3,Members='版主',Levelname='版主||||||16||0' where UserName='"&Newmoderator&"' ")
						Else
							error2 " 此版主已经存在! "
						End If
					End If 
				End If
				RS3.Close:Set Rs3=nothing
			End If
		End If
		Cache.DelCache("ManageUsers")
		SuccessMsg("版主设置成功!")	
	Else%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?Action=SetModerators&UpModers=1&ID=<%=request("ID")%>">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan=4>TEAM's - 编辑版主 - <%=team.execute("Select bbsname from "&isforum&"Bbsconfig where id="&id )(0)%></td>
    </tr>
    <tr align="center" class=a3>
      <td> <input type="checkbox" name="chkall" class="a4" onClick="checkall(this.form)">删?</td>
      <td>用户名</td>
      <td>显示顺序</td>
    </tr>
    <%  
		Dim Rs
		Set Rs=team.Execute("Select id,ManageUser,Issort from "&isforum&"Moderators Where BoardID="& ID)
		Do While Not Rs.Eof
		%>
    <tr align="center" class="a4">
      <td><input type="checkbox" name="isdelete" value="<%=rs(0)%>"></td>
      <td><%=rs(1)%></td>
      <td><%=rs(2)%></td>
    </tr>
    <% Rs.MoveNext
		Loop
		Rs.Close:Set Rs=Nothing
		%>
    <tr align="center" class="a3">
      <td>新增:</td>
      <td><input type='text' name="newmoderator" size="20"></td>
      <td><input type="text" name="newdisplayorder" size="2" value="0"></td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="forumsubmit" value=" 提 交 ">
  &nbsp;
</form>
</center>
<br>
<%
	End If
End Sub
Sub DelForums%>
<br>
<br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?Action=ServerDelForum&ID=<%=request("ID")%>">
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan=5>TEAM's 提示</td>
    </tr>
    <tr align="center">
      <td bgcolor="#FFFFFF"><br>
        <br>
        <br>
        本操作不可恢复，您确定要删除该论坛，清除其中帖子和附件吗?<br>
        注意: 删除论坛并不会更新用户发帖数和积分<br>
        <br>
        <br>
        <br>
        <input type="submit" name="forumsubmit" value=" 确 定 ">
        &nbsp;
        <input type="button" value=" 取 消 " onClick="history.go(-1);"></td>
    </tr>
  </table>
</form>
<br>
<%
End Sub
Sub UniForum()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<br>
<br>
<br>
<br>
<br>
<form method="post" action="?Action=Forumsmerge">
  <table cellspacing="1" cellpadding="4" width="85%" align="center" class="a2">
    <tr class="a1">
      <td colspan="3">合并论坛 - 源论坛的帖子全部转入目标论坛，同时删除源论坛</td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">源论坛:</td>
      <td bgcolor="#FFFFFF" width="60%"><select name="source">
          <option value="">┝ 请选择</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8F8F8" width="40%">目标论坛:</td>
      <td bgcolor="#FFFFFF" width="60%"><select name="target">
          <option value="">┝ 请选择</option>
          <% ForumList_Sel(0) %>
        </select></td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="submit" value="提 交">
  </center>
</form>
<br>
<br>
<%
End Sub

Sub ForumList_Sel(V)
	Dim SQL,ii,RS,W
	Set Rs=Team.Execute("Select ID,BBSname,Followid From "&IsForum&"Bbsconfig Where Followid="&V&" Order By SortNum")
	Do While Not RS.Eof
		W="　┕ "
		If V = 0 Then W="┝ "
		Response.Write "<option value="&RS(0)&""
		If Request("Action") = "Forumadd" Then
			If RS(0) = int(Request("ID")) Then Response.Write " selected "
		End If
		If Request("Action") = "Manages" Then
			If RS(0) = int(Request("RootID")) Then Response.Write " selected "
		End If
		Response.Write ">"&String(ii,"&nbsp;")&""&W&""&RS(1)&"</option>"
		ii=ii+1
		ForumList_Sel RS(0)
		ii=ii-1
		RS.MoveNext
	loop
	Rs.close: Set Rs = Nothing
End Sub

Dim ManageUsers,Moderuser
Sub ForumList(V)
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Hide,BbsName,SortNum,Board_Model From "&IsForum&"Bbsconfig Where Followid=0 Order By SortNum")
	Do While Not RS.Eof
		ManageUsers = team.GroupManages()
		Moderuser = ""
		If isarray(ManageUsers) Then
			for	u=0 to Ubound(ManageUsers,2)
				If ManageUsers(2,u) = Rs(0) Then
					Moderuser = Moderuser & ManageUsers(1,u) & " "
				End If
			Next
		End If
		Select Case RS(1)
			Case 1
				T="隐藏"
			Case Else
				T="正常"
		End Select
		If RS(4)=1 then
			sty = "<a href=?Action=ModelSet_1&ID="&RS(0)&" title=""点击转换显示模式"">简洁模式</a>"
		Else
			sty = "<a href=?Action=ModelSet_0&id="&rs(0)&" title=""点击转换显示模式"">正常模式</a>"
		End If
		Echo "<ul><li><a target=_blank href=../Default.asp?rootid="&RS(0)&"><b>"&RS(2)&"</b></a> - <span class=a4></a> - 显示顺序: <input type=text name=SortNum Value="&RS(3)&" Size=""1""><Input Name=UID value="&RS(0)&" type=hidden> - <a href=""?Action=Forumadd&ID="&RS(0)&""" title=""添加本分类或论坛的下级论坛"">[添加]</a> <a href=""?Action=Manages&ID="&RS(0)&""" title=""编辑本论坛设置"">[编辑]</a> <a href=""?Action=DelForum&ID="&RS(0)&""" title=""删除本论坛及其中所有帖子"">[删除]</a> - [状态: <b>"&T&"</b>]</a> - [显示模式: "&sty&" ] - [<a href=""?Action=SetModerators&ID="&RS(0)&""" title=""编辑本论坛版主"">版主 "&Moderuser&"</a>]</span>"
		Call ForumList_1(Rs(0))
		Echo " </li></ul> "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing
End Sub

Sub ForumList_1(a)
	Dim SQL,RS,Style,S,T,sty
	Set Rs=team.Execute("Select ID,Hide,BbsName,SortNum,Board_Model,Followid From "&IsForum&"Bbsconfig Where Followid="&a&" Order By SortNum")
	Do While Not RS.Eof
		ManageUsers = team.GroupManages()
		Moderuser = ""
		If isarray(ManageUsers) Then
			for	u=0 to Ubound(ManageUsers,2)
				If ManageUsers(2,u) = Rs(0) Then
					Moderuser = Moderuser & ManageUsers(1,u) & " "
				End If
			Next
		End If
		Select Case RS(1)
			Case 1
				T="隐藏"
			Case Else
				T="正常"
		End Select
		Echo "<ul><li>"&String(ii*2,"　")& S &"<a target=_blank href=../Forums.asp?fid="&RS(0)&"><b>"&RS(2)&"</b></a> - <span class=a4></a> - 显示顺序: <input type=text name=SortNum Value="&RS(3)&" Size=""1""><Input Name=UID value="&RS(0)&" type=hidden> - <a href=""?Action=Forumadd&ID="&RS(0)&""" title=""添加本分类或论坛的下级论坛"">[添加]</a> <a href=""?Action=Manages&ID="&RS(0)&"&RootID="&RS(5)&""" title=""编辑本论坛设置"">[编辑]</a> <a href=""?Action=DelForum&ID="&RS(0)&""" title=""删除本论坛及其中所有帖子"">[删除]</a> - [状态: <b>"&T&"</b>]</a> - [<a href=""?Action=SetModerators&ID="&RS(0)&""" title=""编辑本论坛版主"">版主 "&Moderuser&"</a>]</span>"
		Call ForumList_1(Rs(0))
		Echo " </li></ul> "
		RS.MoveNext
	loop
	RS.Close:Set Rs = Nothing

End Sub

%>
