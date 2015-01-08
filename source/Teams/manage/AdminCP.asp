<!--#include file="../Conn.asp"-->
<!--#include file="Const.asp"-->
<SCRIPT src="../js/Command.js"></SCRIPT>
<script>
function admin_Size(num,objname)
{
    var obj=document.getElementById(objname)
    if (parseInt(obj.rows)+num>=3) {
        obj.rows = parseInt(obj.rows) + num;    
    }
    if (num>0)
    {
        obj.width="90%";
    }
}
</script>
<%
Dim ID
Call Master_Us()
Header()
Dim Admin_Class
Admin_Class=",1,"
Call Master_Se()
Select Case Request("action")
	Case "Settingok"
		Call Settingok
	Case "upreg"
		team.Execute("Update ["&isforum&"Clubconfig] Set AgreeMent='"&Replace(Trim(Request.Form("myinfos")),"'","")&"'")
		Cache.DelCache("Club_Class")
		SuccessMsg("注册协议编辑成功!")
	Case Else
		Call Main()
End Select


Sub Main()
	Dim InstalledObjects(10)
	'水印
	InstalledObjects(0) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
	InstalledObjects(1) = "Persits.Jpeg"				'AspJpeg
	InstalledObjects(2) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
	InstalledObjects(3) = "sjCatSoft.Thumbnail"			'sjCatSoft.Thumbnail V2.6
	'上传
	InstalledObjects(4) = "Adodb.Stream"				'Adodb.Stream
	InstalledObjects(5) = "Persits.Upload"				'Aspupload3.0
	InstalledObjects(6) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
	InstalledObjects(7) = "Scripting.FileSystemObject"	'FSO
	'邮件
	InstalledObjects(8) = "JMail.Message"
	InstalledObjects(9) = "CDONTS.NewMail"
%>
<form method="Post" action="?action=Settingok">
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2" >
<tr class="a1"><td>技巧提示</td></tr>
<tr class="a3"><td>
<br><ul><li>选项以加下划线的斜体字显示时，说明此选项和系统效率、负载能力与资源消耗有关(提高效率、或降低效率)，建议依据自身服务器情况进行调整。
</ul></td></tr></table>
<br>

<a name="基本设置"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
<td colspan="2">基本设置</td>
</tr>
<tr>
	<td width="60%"  bgcolor="#F8F8F8"><b>论坛名称:</b><br><span class="a3">论坛名称，将显示在导航条和标题中</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(1)" value="<%=team.Club_Class(1)%>"></td>
</tr>
<tr><% team.Club_Class(2) = "" %>
	<td width="60%"  bgcolor="#F8F8F8"><b>论坛URL:</b><br></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(2)" value="<%
	If team.Club_Class(2)="" Then 
		Response.Write "http://"&Request.ServerVariables("server_name")&""&replace(Request.ServerVariables("script_name"),ManagePath&"Admincp.asp","")&"" 
	Else 
		Response.Write team.Club_Class(2)
	End If 
	%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8"><b>网站名称:</b><br><span class="a3">网站名称，将显示在页面底部的联系方式处</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(3)" value="<%=team.Club_Class(3)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>建站日期:</b><br><span class="a3">网站开始运行的日期，可调用参数。</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(29)" value="<%=team.Club_Class(29)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>网站 URL:</b><br><span class="a3">网站 URL，将作为链接显示在页面底部</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Club_Class(4)" value="<%=team.Club_Class(4)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>网站备案:</b><br><span class="a3">网站备案的号码，显示在页尾右下脚</span></td>
	<td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(59)" value="<%=team.Forum_setting(59)%>"></td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>开启定制首页:</b><br><span class="a3">开启定制首页需要空间支持INDEX.ASP默认首页访问</span></td>
	<td bgcolor="#FFFFFF"><input type="radio" class="radio" name="Forum_setting(111)" value="1" <%If CID(team.Forum_setting(111))=1 Then%>checked<%End If%>> 是 &nbsp; &nbsp; <input type="radio" class="radio" name="Forum_setting(111)" value="0" <%If CID(team.Forum_setting(111))=0 Then%>checked<%End If%>> 否</td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>论坛关闭:</b><br><span class="a3">暂时将论坛关闭，其他人无法访问，但不影响管理员访问</span></td>
	<td bgcolor="#FFFFFF"><input type="radio" class="radio" name="Forum_setting(2)" value="1" <%If team.Forum_setting(2)=1 Then%>checked<%End If%>> 是 &nbsp; &nbsp; <input type="radio" class="radio" name="Forum_setting(2)" value="0" <%If team.Forum_setting(2)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>论坛关闭的原因:</b><br><span class="a3">论坛关闭时出现的提示信息</span></td>
	<td bgcolor="#FFFFFF"><textarea rows="5" name="Forum_setting(3)" cols="60"><%=team.Forum_setting(3)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center><br>

<br><a name="注册与访问控制"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">注册与访问控制</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>允许新用户注册:</b><br><span class="a3">选择“否”将禁止游客注册成为会员，但不影响过去已注册的会员的使用</span></td><td bgcolor="#FFFFFF"><input type="radio" class="radio" name="Forum_setting(4)" value="1" <%If team.Forum_setting(4)=1 Then%>checked<%End If%>> 是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(4)" value="0" <%If team.Forum_setting(4)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>禁止新用户注册说明:</b><br><span class="a3">当论坛禁止游客注册成为会员时,所给出的提示文字!</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Forum_setting(5)" cols="60"><%=team.Forum_setting(5)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>允许同一 Email 注册不同用户:</b><br><span class="a3">选择“否”将只允许一个 Email 地址只能注册一个用户名</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(6)" value="1" <%If team.Forum_setting(6)=1 Then%>checked<%End If%>> 是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(6)" value="0" <%If team.Forum_setting(6)=0 Then%>checked<%End If%>> 否
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Email组件支持:</b><br><span class="a3">请选择正确的邮件发送组件, 只有组件支持的情况下,才可以向用户发送Email验证. 否则请勿开启Email验证功能!</span><br>
	<%If IsObjInstalled(InstalledObjects(8)) Then%>□JMail.Message √<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(9)) Then%>□ CDONTS.NewMail √<br><%End If%>
	<%If Not IsObjInstalled(InstalledObjects(9)) and Not IsObjInstalled(InstalledObjects(8)) Then%><font color=red>注意: 您的服务器不支持发送邮件功能!</font><%End If%>
	</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(1)" value="0" <%If team.Forum_setting(1)=0 Then%>checked<%End If%>> 无 <br>
	<input type="radio" class="radio" name="Forum_setting(1)" value="1" <%If team.Forum_setting(1)=1 Then%>checked<%End If%>> JMail.Message	<br>
	<input type="radio" class="radio" name="Forum_setting(1)" value="2" <%If team.Forum_setting(1)=2 Then%>checked<%End If%>> CDONTS.NewMail	<br></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>SMTP Server地址:</b><br><span class="a3">邮件服务器的地址</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(58)" value="<%=team.Forum_setting(58)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>邮件服务器登录名:</b><br><span class="a3">登录邮件服务器的用户名</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(41)" value="<%=team.Forum_setting(41)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>邮件服务器登录密码:</b><br><span class="a3">登录邮件服务器的密码</span></td><td bgcolor="#FFFFFF"><input type="password" size="30" name="Forum_setting(55)" value="<%=team.Forum_setting(55)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>邮件发送人地址:</b><br><span class="a3">显示在邮件的发送人地址</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(57)" value="<%=team.Forum_setting(57)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>新用户注册验证:</b><br><span class="a3">选择“无”用户可直接注册成功；选择“Email 验证”将向用户注册 Email 发送一封验证邮件以确认邮箱的有效性；选择“人工审核”将由管理员人工逐个确定是否允许新用户注册</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(7)" value="0" <%If team.Forum_setting(7)=0 Then%>checked<%End If%>> 无 <br>
	<input type="radio" class="radio" name="Forum_setting(7)" value="1" <%If team.Forum_setting(7)=1 Then%>checked<%End If%>> 注册成功发送Email	<br>
	<input type="radio" class="radio" name="Forum_setting(7)" value="2" <%If team.Forum_setting(7)=2 Then%>checked<%End If%>> Email 验证	<br>
	<input type="radio" class="radio" name="Forum_setting(7)" value="3" <%If team.Forum_setting(7)=3 Then%>checked<%End If%>> 人工审核	<br>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>注册Email的内容:</b><br><span class="a3">用户注册邮件的内容</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(23)" cols="60"  style="height:70;overflow-y:visible;"><%=team.Club_Class(23)%></textarea>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>IP 注册间隔限制(秒):</b><br><span class="a3">同一 IP 在本时间间隔内将只能注册一个帐号，限制对自修改后的新注册用户生效，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(10)" value="<%=team.Forum_setting(10)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>登陆用户尝试次数限制:</b><br><span class="a3">当用户登陆忘记密码时候， 可以尝试密码的次数限制，建议设置为默认值 5 或在不超过 10 范围内取值，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(54)" value="<%=team.Forum_setting(54)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>用户信息保留关键字:</b><br><span class="a3">用户在其用户信息(如用户名、昵称、自定义头衔等)中无法使用这些关键字。每个关键字一行，可使用通配符 "*" 如 "*版主*"(不含引号)</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(25)" cols="60"><%=team.Club_Class(25)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>新手见习期限(分钟):</b><br><span class="a3">新注册用户在本期限内将无法发帖，不影响斑竹和管理员，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(14)" value="<%=team.Forum_setting(14)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>发送欢迎短消息:</b><br><span class="a3">选择“是”将自动向新注册用户发送一条欢迎短消息</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(15)" value="1" <%If team.Forum_setting(15)=1 Then%>checked<%End If%>> 是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(15)" value="0" <%If team.Forum_setting(15)=0 Then%>checked<%End If%>> 否
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>欢迎短消息内容:</b><br><span class="a3">系统发送的欢迎短消息的内容</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Forum_setting(16)" cols="60" style="height:70;overflow-y:visible;"><%=team.Forum_setting(16)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>游客查看帖子权限:</b><br><span class="a3">当版块设置没有现在查看权限的时候，可以通过此选择限制游客查看帖子内容</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(110)" value="0" <%If CID(team.Forum_setting(110))=0 Then%>checked<%End If%>> 不限制 <BR> 
	<input type="radio" class="radio" name="Forum_setting(110)" value="1" <%If CID(team.Forum_setting(110))=1 Then%>checked<%End If%>> 部分限制
	<BR>  
	<input type="radio" class="radio" name="Forum_setting(110)" value="2" <%If CID(team.Forum_setting(110))=2 Then%>checked<%End If%>> 完全限制
	</td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>注册许可协议:</b><br><span class="a3">新用户注册时显示许可协议</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(17)" value="1" <%If team.Forum_setting(17)=1 Then%>checked<%End If%>> 是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(17)" value="0" <%If team.Forum_setting(17)=0 Then%>checked<%End If%>> 否
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>许可协议内容:</b></td><td bgcolor="#FFFFFF"><a href="#注册协议"><span class="a3">点击查看注册许可协议的详细内容</span></a></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="界面与显示方式"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">界面与显示方式</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>默认论坛风格:</b><br><span class="a3">论坛默认的界面风格，游客和使用默认风格的会员将以此风格显示</span></td><td bgcolor="#FFFFFF"><select name="Forum_setting(18)">
<%	Dim SQL,RS,SytyleID,MyCheck
	Set Rs=team.Execute("Select ID,StyleName From ["&IsForum&"Style] order by StyleHid Asc")
	Do While Not RS.Eof
		MyCheck = ""
		If Rs(0) = Int(team.Forum_setting(18)) Then 
			MyCheck = " selected=""selected"""
		End if
		Echo  "<option value="""&RS(0)&""""&MyCheck&">"&rs(1)&"</option>"
		Rs.MoveNext
	Loop
	RS.CLOSE:Set RS=Nothing
%>
	</select></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>每页显示主题数:</b><br><span class="a3">主题列表中每页显示主题数目</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(19)" value="<%=team.Forum_setting(19)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>主题显示字数:</b><br><span class="a3">主题列表中每个主题显示的字数</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(88)" value="<%=team.Forum_setting(88)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>每页显示贴数:</b><br><span class="a3">帖子列表中每页显示帖子数目</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(20)" value="<%=team.Forum_setting(20)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>每页显示会员数:</b><br><span class="a3">会员列表中每页显示会员数目</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(21)" value="<%=team.Forum_setting(21)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>主题列表最大页数:</b><br><span class="a3">主题列表中用户可以翻阅到的最大页数，建议设置为默认值 1000，或在不超过 2500 范围内取值，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(25)" value="<%=team.Forum_setting(25)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>热门话题最低贴数:</b><br><span class="a3">超过一定帖子数的话题将显示为热门话题</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(22)" value="<%=team.Forum_setting(22)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>星星升级阀值:</b><br><span class="a3">星星数在达到此阀值(设为 N)时，N 个星星显示为 1 个月亮、N 个月亮显示为 1 个太阳。默认值为 3，如设为 0 则取消此项功能，始终以星星显示</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(23)" value="<%=team.Forum_setting(23)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示最近访问论坛数量:</b><br><span class="a3">设置在论坛列表和帖子浏览中，显示的最近访问过的论坛下拉列表数量，建议设置为 30 以内，0 为关闭此功能</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(24)" value="<%=team.Forum_setting(24)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>版主显示方式:</b><br><span class="a3">首页论坛列表中版主显示方式</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(26)" value="1" <%If team.Forum_setting(26)=1 Then%>checked<%End If%>> 平面显示  &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(26)" value="0" <%If team.Forum_setting(26)=0 Then%>checked<%End If%>> 下拉菜单</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>首页显示论坛的下级子论坛:</b><br><span class="a3">首页论坛列表中在论坛描述下方显示下级子论坛名字和链接(如果存在的话)。注意: 本功能不考虑子论坛特殊浏览权限的情况，只要存在即会被显示出来</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(27)" value="1" <%If team.Forum_setting(27)=1 Then%>checked<%End If%>> 是&nbsp; &nbsp;  
	<input type="radio" class="radio" name="Forum_setting(27)" value="0" <%If team.Forum_setting(27)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>简洁版本的排列数:</b><br><span class="a3">当论坛版块采用简介版本排序的时候,多于此排列数则自动另起一排.</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(32)" value="<%=team.Forum_setting(32)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示风格下拉菜单:</b><br><span class="a3">设置是否在论坛导航显示可用的论坛风格下拉菜单，用户可以通过此菜单切换不同的论坛风格</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(28)" value="1" <%If team.Forum_setting(28)=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(28)" value="0" <%If team.Forum_setting(28)=0 Then%>checked<%End If%>> 否</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示自定义下拉菜单:</b><br><span class="a3">设置是否在论坛导航显示可用的论坛自定义下拉菜单，用户可以通过此菜单切换不同的论坛插件版块</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(92)" value="1" <%If team.Forum_setting(92)=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(92)" value="0" <%If team.Forum_setting(92)=0 Then%>checked<%End If%>> 否</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示友情链接:</b><br><span class="a3">设置是否在论坛首页显示<B>友情链接</B>状态栏。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(36)" value="1" <%If team.Forum_setting(36)=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(36)" value="0" <%If team.Forum_setting(36)=0 Then%>checked<%End If%>> 否</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示在线状况:</b><br><span class="a3">设置是否在论坛首页显示<B>在线状况</B>状态栏。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(40)" value="1" <%If team.Forum_setting(40)=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(40)" value="0" <%If team.Forum_setting(40)=0 Then%>checked<%End If%>> 否</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示最新帖子:</b><br><span class="a3">设置是否在论坛首页显示<B>最新帖子(3格)</B>状态栏。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(113)" value="1" <%If CID(team.Forum_setting(113))=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(113)" value="0" <%If CID(team.Forum_setting(113))=0 Then%>checked<%End If%>> 否</td> 
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>快速发帖:</b><br><span class="a3">浏览论坛和帖子页面底部显示快速发帖表单</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(29)" value="1" <%If team.Forum_setting(29)=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(29)" value="0" <%If team.Forum_setting(29)=0 Then%>checked<%End If%>> 否</td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="搜索引擎优化"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">搜索引擎优化</td></tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>开启伪静态:</b><br><span class="a3">将页面虚拟为静态文件，需要空间支持。（开启前请向空间商询问，并加载TEAM专用配置文件）</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(65)" value="1" <%If CID(team.Forum_setting(65))=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(65)" value="0" <%If CID(team.Forum_setting(65))=0 Then%>checked<%End If%>> 否</td>
</tr




<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>标题附加字:</b><br><span class="a3">网页标题通常是搜索引擎关注的重点，本附加字设置将出现在标题中论坛名称的后面，如果有多个关键字，建议用 "|"、","(不含引号) 等符号分隔。 </span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(66)" value="<%=team.Forum_setting(66)%>">
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Meta Keywords:</b><br><span class="a3">Keywords 项出现在页面头部的 Meta 标签中，用于记录本页面的关键字，多个关键字间请用半角逗号 "," 隔开</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(30)" value="<%=team.Forum_setting(30)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>Meta Description:</b><br><span class="a3">Description 出现在页面头部的 Meta 标签中，用于记录本页面的概要与描述</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(31)" value="<%=team.Forum_setting(31)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>其他头部信息:</b><br><span class="a3">如需在 &lt;head&gt;&lt;/head&gt; 中添加其他的 html 代码，可以使用本设置，否则请留空</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(26)" cols="60"  style="height:70;overflow-y:visible;"><%=team.Club_Class(26)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="论坛功能"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">论坛功能</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>启用 RSS</i></u>:</b><br><span class="a3">选择“是”，论坛将允许用户使用 RSS 客户端软件接收最新的论坛帖子更新。注意: 在分论坛很多的情况下，本功能可能会加重服务器负担</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(33)" value="1" <%If team.Forum_setting(33)=1 Then%>checked<%End If%>>
	是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(33)" value="0" <%If team.Forum_setting(33)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>RSS TTL(分钟)</i></u>:</b><br><span class="a3">TTL(Time to Live) 是 RSS 2.0 的一项属性，用于控制订阅内容的自动刷新时间，时间越短则资料实时性就越高，但会加重服务器负担，通常可设置为 30～180 范围内的数值</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(34)" value="<%=team.Forum_setting(34)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>网站脚本过期时间:</b><br><span class="a3">使用Server.scripttimeout来减少ASP意外错误而使服务器瘫痪，默认设置为20秒，优先度低于服务器本身设置。</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(91)" value="<%=team.Forum_setting(91)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示程序运行信息:</b><br><span class="a3">选择“是”将在页脚处显示程序运行时间和数据库查询次数</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(37)" value="1" <%If team.Forum_setting(37)=1 Then%>checked<%End If%>> 是 &nbsp; &nbsp; 
	<input type="radio" class="radio" name="Forum_setting(37)" value="0" <%If team.Forum_setting(37)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>生日显示与邮件祝福:</b><br><span class="a3">设置是否在首页显示今日过生日的会员，并向他们发送邮件祝福。如果您的论坛用户数量很大，过生日会员的列表可能会影响首页页面美观，向其逐个发送邮件祝福也会耗费一定的系统资源</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(38)" value="0" <%If team.Forum_setting(38)=0 Then%>checked<%End If%>> 无<br>
	<input type="radio" class="radio" name="Forum_setting(38)" value="1" <%If team.Forum_setting(38)=1 Then%>checked<%End If%>> 仅在首页显示过生日会员<br>
	<input type="radio" class="radio" name="Forum_setting(38)" value="2" <%If team.Forum_setting(38)=2 Then%>checked<%End If%>> 仅向过生日会员发送邮件祝福<br>
	<input type="radio" class="radio" name="Forum_setting(38)" value="3" <%If team.Forum_setting(38)=3 Then%>checked<%End If%>> 显示并发送邮件祝福</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>显示在线用户:</b><br><span class="a3">在首页和论坛列表页显示在线会员列表</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(39)" value="0" <%If team.Forum_setting(39)=0 Then%>checked<%End If%>> 不显示<br>
	<input type="radio" class="radio" name="Forum_setting(39)" value="1" <%If team.Forum_setting(39)=1 Then%>checked<%End If%>> 仅在首页显示<br>
	<input type="radio" class="radio" name="Forum_setting(39)" value="2" <%If team.Forum_setting(39)=2 Then%>checked<%End If%>> 仅在分论坛显示<br>
	<input type="radio" class="radio" name="Forum_setting(39)" value="3" <%If team.Forum_setting(39)=3 Then%>checked<%End If%>> 在首页和分论坛显示</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>显示论坛跳转菜单</i></u>:</b><br><span class="a3">选择“是”将在列表页面下部显示快捷跳转菜单。注意: 当分论坛很多时，本功能会严重加重服务器负担</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(42)" value="1" <%If team.Forum_setting(42)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(42)" value="0" <%If team.Forum_setting(42)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>帖子在新窗口打开:</b><br><span class="a3">选择是否让帖子列表的打开方式是否为新窗口打开还是在本窗口打开</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(43)" value="1" <%If team.Forum_setting(43)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(43)" value="0" <%If team.Forum_setting(43)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>统计系统缓存时间(分钟)</i></u>:</b><br><span class="a3">统计数据缓存更新的时间，数值越大，数据更新频率越低，越节约资源，但数据实时程度越低，建议设置为 60 以上，以免占用过多的服务器资源。</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(44)" value="<%=team.Forum_setting(44)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>用户在线时间更新时长(分钟)</i></u>:</b><br><span class="a3">TEAM! 可统计每个用户总共和当月的在线时间，本设置用以设定更新用户在线时间的时间频率。例如设置为 10，则用户每在线 10 分钟更新一次记录。本设置值越小，则统计越精确，但消耗资源越大。建议设置为 5～30 范围内，0 为不记录用户在线时间</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(45)" value="<%=team.Forum_setting(45)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>管理记录保留时间(月)</i></u>:</b><br><span class="a3">系统中保留管理记录的时间，默认为 3 个月，建议在 3~6 个月的范围内取值</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(47)" value="<%=team.Forum_setting(47)%>"></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>
<br>
<a name="安全控制"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
	<td colspan="2">安全控制</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>启用CC防护:</b><br><span class="a3">启用CC防护将导致使用代理服务器的用户无法访问论坛，建议仅在必需时打开。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(106)" value="1" <%If team.Forum_setting(106)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(106)" value="0" <%If team.Forum_setting(106)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>启用验证码:</b><br><span class="a3">图片验证码可以避免用灌水或刷新程序恶意批量发布或提交信息，请选择需要打开验证码的操作。设置当用户登录或发帖的几次才出现验证码验证页面，设置为0时，则关闭此功能</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(48)" value="<%=team.Forum_setting(48)%>">
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>启用来源检测:</b><br><span class="a3">为了防止用户从外部提交数据!</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(49)" value="1" <%If team.Forum_setting(49)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(49)" value="0" <%If team.Forum_setting(49)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>发帖灌水预防(秒):</b><br><span class="a3">两次发帖间隔小于此时间，或两次发送短消息间隔小于此时间的二倍将被禁止，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(50)" value="<%=team.Forum_setting(50)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>搜索时间限制(秒):</b><br><span class="a3">两次搜索间隔小于此时间将被禁止，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(51)" value="<%=team.Forum_setting(51)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b><u><i>60 秒最大搜索次数:</u></i></b><br><span class="a3">论坛系统每 60 秒系统响应的最大搜索次数，0 为不限制。注意: 如果服务器负担较重，建议设置为 5，或在 5~20 范围内取值，以避免过于频繁的搜索造成数据表被锁</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(52)" value="<%=team.Forum_setting(52)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>最大搜索结果:</b><br><span class="a3">每次搜索获取的最大结果数，建议设置为默认值 500，或在不超过 1500 范围内取值</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(53)" value="<%=team.Forum_setting(53)%>"></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="时间段及过滤设置"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">时间段及过滤设置</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>论坛时间段设置:</b><br><span class="a3">设置每天各个时间段内用户的访问权限及动作权限，此功能对应下面的时间段设置功能。打开此功能，下面的时间段设置功能才会开启。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(56)" value="0" <%If team.Forum_setting(56)=0 Then%>checked<%End If%>> 关闭<br> 
	<input type="radio" class="radio" name="Forum_setting(56)" value="1" <%If team.Forum_setting(56)=1 Then%>checked<%End If%>> 定时关闭<br> 
	<input type="radio" class="radio" name="Forum_setting(56)" value="2" <%If team.Forum_setting(56)=2 Then%>checked<%End If%>> 定时只读<br> 
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>禁止发帖时间段:</b><br><span class="a3">每天该时间段内用户不能发帖，需要与上面的功能配合使用!</span></td><td bgcolor="#FFFFFF"><table cellspacing="1" cellpadding="1" width="90%"><tr><%
		Dim Openclock
		openclock=Split(team.Forum_setting(0),"*")
		For i= 0 to UBound(openclock)
			%>
			<td><input type="checkbox" name="openclock<%=i%>" value="1" <%If openclock(i)="1" Then %>checked<%End If%>><%=i%>点开</td>
			<%
			If (i+1) mod 3 = 0 Then Response.Write "</tr><tr>"
		Next
 %></table></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>用户帖子过滤设置:</b><br><span class="a3">自动屏蔽该列表显示的用户所发表的帖子内容，每行一个用户名。
	</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(7)" cols="60"><%=team.Club_Class(7)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>发帖敏感设置:</b><br><span class="a3">自动替换用户发贴中出现的各种关键字，匹配格式如下：需要过虑的字=替换后的字，过虑完成后，将只显示过滤后的文字，每行一个匹配段。
	</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(5)" cols="60"><%=team.Club_Class(5)%></textarea></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>用户IP屏蔽设置:</b><br><span class="a3">限制该IP的用户登陆到论坛，每行一个IP，可以采用通配符 * 进行IP段的限制，如192.168.1.*，即显示了IP 192.168.1.1 - 192.168.1.255的所有IP访问权限。</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(6)" cols="60"><%=team.Club_Class(6)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="用户权限"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">用户权限</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>版主评分限制:</b><br><span class="a3">设置版主只能在自身所管辖的论坛范围内对帖子进行评分。本限制只对版主有效，允许评分的普通用户及超级版主、管理员不受此限制，因此如果赋予这些用户评分权限，他们仍将可以在全论坛范围内进行评分</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(61)" value="1" <%If team.Forum_setting(61)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(61)" value="0" <%If team.Forum_setting(61)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>允许重复评分:</b><br><span class="a3">选择“是”将允许用户对一个帖子进行多次评分，默认为“否”</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(62)" value="1" <%If team.Forum_setting(62)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(62)" value="0" <%If team.Forum_setting(62)=0 Then%>checked<%End If%>> 否</td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>评分时间限制(小时):</b><br><span class="a3">帖子发表后超过此时间限制其他用户将不能对此帖评分，版主和管理员不受此限制，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(60)" value="<%=team.Forum_setting(60)%>"></td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>允许向管理级别报告帖子:</b><br><span class="a3">允许会员通过短消息向版主或管理员报告反映帖子。注意: 如果当前论坛或分类没有设置版主，同时本设定设置为“只允许报告给版主”，系统会自动将报告内容发送给超级版主，以此类推</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(63)" value="0" <%If team.Forum_setting(63)=0 Then%>checked<%End If%>> 禁止用户报告<br> 
	<input type="radio" class="radio" name="Forum_setting(63)" value="1" <%If team.Forum_setting(63)=1 Then%>checked<%End If%>> 仅允许向版主报告<br>
	<input type="radio" class="radio" name="Forum_setting(63)" value="2" <%If team.Forum_setting(63)=2 Then%>checked<%End If%>> 仅允许向版主和超级版主报告<br>
	<input type="radio" class="radio" name="Forum_setting(63)" value="3" <%If team.Forum_setting(63)=3 Then%>checked<%End If%>> 许向所有管理人员报告
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>帖子和标题最小字数(字节):</b><br><span class="a3">管理组成员可通过“发帖不受限制”设置而不受影响，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(64)" value="<%=team.Forum_setting(64)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>标题最大字数(字节):</b><br><span class="a3">管理组成员可通过“发帖不受限制”设置而不受影响，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(89)" value="<%=team.Forum_setting(89)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>帖子最大字数(字节):</b><br><span class="a3">管理组成员可通过“发帖不受限制”设置而不受影响</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(67)" value="<%=team.Forum_setting(67)%>"></td>
</tr>

<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>投票最大选项数:</b><br><span class="a3">设定发布投票包含的最大选项数</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(68)" value="<%=team.Forum_setting(68)%>"></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="附件设置"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">附件设置</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>帖子中显示图片附件:</b><br><span class="a3">在帖子中直接将图片或动画附件显示出来，而不需要点击附件链接</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(69)" value="1" <%If team.Forum_setting(69)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(69)" value="0" <%If team.Forum_setting(69)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>隐藏附件的网站目录路径:</b><br><span class="a3">开启此功能，用户将不能查看附件的所在路径，对防止盗链有一定帮助</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(93)" value="1" <%If team.Forum_setting(93)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(93)" value="0" <%If team.Forum_setting(93)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>选择上传组件:</b><br><span class="a3">
	请选择合适的上传方式：<br>
	<%If IsObjInstalled(InstalledObjects(4)) Then%>□ 无组件  √<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(5)) Then%>□ Aspupload  √<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(6)) Then%>□ SA-FileUp  √<%End If%></span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(70)" value="999" <%If team.Forum_setting(70)="999" Then%>checked<%End If%>> 关闭 <BR>
	<input type="radio" class="radio" name="Forum_setting(70)" value="0" <%If team.Forum_setting(70)=0 Then%>checked<%End If%>> 无组件上传类 <BR>
	<input type="radio" class="radio" name="Forum_setting(70)" value="1" <%If team.Forum_setting(70)=1 Then%>checked<%End If%>> Aspupload3.0组件 <BR>
	<input type="radio" class="radio" name="Forum_setting(70)" value="2" <%If team.Forum_setting(70)=2 Then%>checked<%End If%>> SA-FileUp 4.0组件
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>允许附件文件的大小:</b><br><span class="a3">限制论坛所有等级用户每次上传的附件大小，另外在用户组可以设置每个分组的上传大小限制，大小为不超过此处的设置为准。</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(71)" value="<%=team.Forum_setting(71)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>允许用户头像附件的大小:</b><br><span class="a3">限制用户上传的头像附件大小。</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(72)" value="<%=team.Forum_setting(72)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>头像尺寸设置(宽度|高度):</b><br><span class="a3">默认120*120</span></td><td bgcolor="#FFFFFF">
	宽: <input type="text" size="10" name="Forum_setting(108)" value="<%=team.Forum_setting(108)%>">&nbsp;
	高: <input type="text" size="10" name="Forum_setting(109)" value="<%=team.Forum_setting(109)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>允许用户上传附件的类型:</b><br><span class="a3">设置本论坛中允许上传的附件扩展名，多个扩展名之间用竖号 "|" 分割。如: rar|jpg|txt，此处设置为总的允许上传的文件类型，用户组可以独立设置每个组的详细分类。系统自动屏蔽如EXE和ASP为后缀的文件类型。</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(73)" value="<%=team.Forum_setting(73)%>"></td>
</tr>
<tr class="a1"><td colspan="2">图片水印设置</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>选取水印组件:</b><br><span class="a3">打开此功能，系统将自动为用户上传的图片添加水印效果。此功能需要下例组件支持，暂不支持动画 GIF 格式。请选择合适的水印组件。<br>
	<%If IsObjInstalled(InstalledObjects(0)) Then%>□ CreatePreviewImage  √<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(1)) Then%>□ AspJpeg组件  √<br><%End If%>
	<%If IsObjInstalled(InstalledObjects(2)) Then%>□ SA-ImgWriter  √<%End If%>
	<%If Not (IsObjInstalled(InstalledObjects(0)) or IsObjInstalled(InstalledObjects(1)) or IsObjInstalled(InstalledObjects(2)) ) Then%><font color=red> 系统不支持任何水印组件 </font><%End If%>
</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(74)" value="999" <%If team.Forum_setting(74)=999 Then%>checked<%End If%>>
	关闭 <br>
	<input type="radio" class="radio" name="Forum_setting(74)" value="0" <%If team.Forum_setting(74)=0 Then%>checked<%End If%>> CreatePreviewImage组件 <br>
	<input type="radio" class="radio" name="Forum_setting(74)" value="1" <%If team.Forum_setting(74)=1 Then%>checked<%End If%>> AspJpeg组件 <br>
	<input type="radio" class="radio" name="Forum_setting(74)" value="2" <%If team.Forum_setting(74)=2 Then%>checked<%End If%>> SA-ImgWriter组件</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>水印效果设置开关:</b><br><span class="a3">此功能需要上面的组件功能支持，只有组件支持，并选择了合适的组件支持，才可以使用本功能。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(75)" value="0" <%If team.Forum_setting(75)=0 Then%>checked<%End If%>> 关闭水印效果 <br>
	<input type="radio" class="radio" name="Forum_setting(75)" value="1" <%If team.Forum_setting(75)=1 Then%>checked<%End If%>> 图片水印效果<br>
	<input type="radio" class="radio" name="Forum_setting(75)" value="2" <%If team.Forum_setting(75)=2 Then%>checked<%End If%>> 文字水印效果</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>图片地址:</b><br><span class="a3">默认水印图片位于 images/uplogo.gif，您可替换此文件或修改以下的图片地址以实现不同的水印效果。</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(76)" value="<%=team.Forum_setting(76)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>水印图片或文字区域设置(宽度|高度):</b><br><span class="a3">默认88*31</span></td><td bgcolor="#FFFFFF">
	宽: <input type="text" size="10" name="Forum_setting(77)" value="<%=team.Forum_setting(77)%>">&nbsp;
	高: <input type="text" size="10" name="Forum_setting(35)" value="<%=team.Forum_setting(35)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>生成预览图片大小:</b><br><span class="a3"></span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(78)" value="1" <%If team.Forum_setting(78)=1 Then%>checked<%End If%>> 固定 
	<input type="radio" class="radio" name="Forum_setting(78)" value="0" <%If team.Forum_setting(78)=0 Then%>checked<%End If%>> 等比例缩小</td></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>文字水印:</b><br><span class="a3">仅支持纯文字</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(79)" value="<%=team.Forum_setting(79)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>水印字体大小:</b><br><span class="a3">添加水印文字的字体大小</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(80)" value="<%=team.Forum_setting(80)%>">PX</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>水印字体颜色:</b><br><span class="a3">添加水印文字的字体颜色</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(81)" value="<%=team.Forum_setting(81)%>" id=d_bgcolor> 
	<img border="0" src="images/rect.gif" style="cursor:pointer;background-Color:<%=team.Forum_setting(81)%>;" width="18" id="s_bgcolor" onclick="SelectColor('bgcolor')">
	<Script>
	function SelectColor(what){
		var dEL = document.getElementById("d_"+what);
		var sEL = document.getElementById("s_"+what);
		var arr = showModalDialog("images/selcolor.htm", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
		if (arr) {
			dEL.value=arr;
			sEL.style.backgroundColor=arr;
		}
	}
	</script></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>透明度颜色设置:</b><br><span class="a3">添加水印透明度颜色</span></td><td bgcolor="#FFFFFF">
	<input type="text" size="30" name="Forum_setting(107)" value="<%=team.Forum_setting(9)%>" id="d_bgcolor1"> 
	<img border="0" src="images/rect.gif" style="cursor:pointer;background-Color:<%=team.Forum_setting(9)%>;" width="18" id="s_bgcolor" onclick="SelectColor('bgcolor1')">
	<Script>
	function SelectColor(what){
		var dEL = document.getElementById("d_"+what);
		var sEL = document.getElementById("s_"+what);
		var arr = showModalDialog("images/selcolor.htm", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
		if (arr) {
			dEL.value=arr;
			sEL.style.backgroundColor=arr;
		}
	}
	</script></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>水印文字字体:</b><br><span class="a3">添加水印文字的字体</span></td><td bgcolor="#FFFFFF">
	<SELECT name="Forum_setting(82)">
		<option value="宋体" <%SetColors("宋体")%>>宋体</option>
		<option value="楷体_GB2312" <%SetColors("楷体_GB2312")%>>楷体</option>
		<option value="新宋体" <%SetColors("新宋体")%>>新宋体</option>
		<option value="黑体" <%SetColors("黑体")%>>黑体</option>
		<option value="隶书" <%SetColors("隶书")%>>隶书</option>
		<OPTION value="Andale Mono" <%SetColors("Andale Mono")%>>Andale Mono</OPTION> 
		<OPTION value="Arial" <%SetColors("Arial")%>>Arial</OPTION> 
		<OPTION value="Arial Black" <%SetColors("Arial Black")%>>Arial Black</OPTION> 
		<OPTION value="Book Antiqua" <%SetColors("Book Antiqua")%>>Book Antiqua</OPTION>
		<OPTION value="Century Gothic" <%SetColors("Century Gothic")%>>Century Gothic</OPTION> 
		<OPTION value="Comic Sans MS" <%SetColors("Comic Sans MS")%>>Comic Sans MS</OPTION>
		<OPTION value="Courier New" <%SetColors("Courier New")%>>Courier New</OPTION>
		<OPTION value="Georgia" <%SetColors("Georgia")%>>Georgia</OPTION>
		<OPTION value="Impact" <%SetColors("Impact")%>>Impact</OPTION>
		<OPTION value="ahoma" <%SetColors("ahoma")%>>Tahoma</OPTION>
		<OPTION value="Times New Roman" <%SetColors("Times New Roman")%>>Times New Roman</OPTION>
		<OPTION value="Trebuchet MS" <%SetColors("Trebuchet MS")%>>Trebuchet MS</OPTION>
		<OPTION value="Script MT Bold" <%SetColors("Script MT Bold")%>>Script MT Bold</OPTION>
		<OPTION value="Stencil" <%SetColors("Stencil")%>>Stencil</OPTION>
		<OPTION value="Verdana" <%SetColors("Verdana")%>>Verdana</OPTION>
		<OPTION value="Lucida Console" <%SetColors("Lucida Console")%>>Lucida Console</OPTION>
	</SELECT></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>图片附件添加水印的坐标:</b><br><span class="a3">请在此选择水印添加的位置(共 5 个位置可选)。</span></td><td bgcolor="#FFFFFF">
	<table border="0" cellspacing="1" cellpadding="4" width="100%" class="a2">
	<tr class="tab4">
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="0" <%If CID(team.Forum_setting(83))=0 Then%>checked<%End If%>> 左上 
	 </td>
	 <td>&nbsp;</td>
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="3" <%If CID(team.Forum_setting(83))=3 Then%>checked<%End If%>> 右上
	 </td>
	</tr>
	<tr class="tab4">
	 <td>&nbsp;</td>
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="2" <%If CID(team.Forum_setting(83))=2 Then%>checked<%End If%>> 居中 </td> 
	 <td>&nbsp;</td>
	</tr>
	<tr class="tab4">
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="1" <%If CID(team.Forum_setting(83))=1 Then%>checked<%End If%>> 左下
	 </td>
	 <td>&nbsp;</td>
	 <td>
	 <input type="radio" class="radio" class="radio" name="Forum_setting(83)" value="4" <%If CID(team.Forum_setting(83))=4 Then%>checked<%End If%>> 右下
	 </td>
	</tr>
	</table>
</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>水印融合度:</b><br><span class="a3">设置水印图片与原始图片的融合度，数值越大水印图片透明度越低。本功能需要开启水印功能后才有效</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(84)" value="<%=team.Forum_setting(84)%>">如60%请填写0.6</td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="JS 调用"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">JS 调用</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>启用 JS 调用:</b><br><span class="a3">JS(JavaScript)调用将使得您可以将论坛新帖、排行等资料嵌入到您的普通网页中，访问者无需访问论坛即可获知论坛最近更新的情况</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(85)" value="1" <%If team.Forum_setting(85)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(85)" value="0" <%If team.Forum_setting(85)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>JS 数据缓存时间(秒):</b><br><span class="a3">由于一些排序检索操作比较耗费资源，JS 调用程序采用缓存技术来实现数据的定期更新，默认值 1440 分钟 ，建议设置不低于 600 的数值，0 为不缓存(极耗费系统资源)</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(86)" value="<%=team.Forum_setting(86)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>JS 来路限制:</b><br><span class="a3">为了避免其他网站非法调用论坛数据，加重您的服务器负担，您可以设置允许调用论坛 JS 的来路域名列表，只有在列表中的域名和网站，才能通过 JS 调用您论坛的信息。每个域名一行，不支持通配符，请勿包含 http:// 或其他非域名内容，留空为不限制来路，即任何网站均可调用</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(28)" cols="60"><%=team.Club_Class(28)%></textarea></td>
</tr></table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center>

<br><a name="其他设置"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
	<td colspan="2">其他设置</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>默认时差:</b><br><span class="a3">当地时间与 GMT 的时差</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(90)" value="<%=team.Forum_setting(90)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>编辑帖子时间限制(分钟):</b><br><span class="a3">帖子作者发帖后超过此时间限制将不能再编辑帖，版主和管理员不受此限制，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(94)" value="<%=team.Forum_setting(94)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>编辑帖子附加编辑记录:</b><br><span class="a3">在 60 秒后编辑帖子添加“本帖由 xxx 于 xxxx-xx-xx 编辑”字样。管理员编辑不受此限制</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(95)" value="1" <%If team.Forum_setting(95)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(95)" value="0" <%If team.Forum_setting(95)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>打开工资自动发放功能:</b><br><span class="a3">选择“是”将打开自动发放工资的功能，系统自动在每月的第一天发放工资给用户。具体设置请进入 <B>用户管理</B> 的 <B>工资管理</B> 选项。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(96)" value="1" <%If team.Forum_setting(96)=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(96)" value="0" <%If team.Forum_setting(96)=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>论坛头像的总数:</b><br><span class="a3">请将头像图片上传到论坛的 images/Upface/下面，命名按照原来的格式排列，如默认30个图片为从1-30，另外添加的则从31开始。并在此处填写正确的图片总数。</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(100)" value="<%=team.Forum_setting(100)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>默认发贴模式:</b><br><span class="a3">设置默认的发贴模式。</span></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(98)" value="1" <%If team.Forum_setting(98)=1 Then%>checked<%End If%>> 所见即所得模式 
	<input type="radio" class="radio" name="Forum_setting(98)" value="0" <%If team.Forum_setting(98)=0 Then%>checked<%End If%>> UBB模式</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" valign="top"><b>管理操作及评分理由选项</b><br><span class="a3">本设定将在用户执行部分管理操作或评分时显示，每个理由一行，如果空行则显示一行分隔符“--------”，用户可选择本设定中预置的理由选项或自行输入</span></td><td bgcolor="#FFFFFF"><textarea rows="5" name="Club_Class(8)" cols="60"><%=team.Club_Class(8)%></textarea>
	</td>
</tr>
<tr>
	<td colspan="2">CC视频联盟设置</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>您在CC视频联盟的数字ID</b><br><span class="a3">CC视频联盟为网站、论坛、博客提供专业的视频发布服务（包括上传、录制、分享等），同时使用智能匹配技术为用户带来源源不断的视频点播量,您只要在 <A HREF="http://union.bokecc.com/">http://union.bokecc.com/</A>注册,然后将您的ID号码填写到此处即可.  </span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(114)" value="<%=team.Forum_setting(114)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>您在CC视频联盟的展区ID</b><br><span class="a3"> </span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(115)" value="<%=team.Forum_setting(115)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>是否开放展区:</b></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(116)" value="1" <%If CID(team.Forum_setting(116))=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(116)" value="0" <%If CID(team.Forum_setting(116))=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td colspan="2">www.fs2you.com的大附件支持</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>开启fs2you网盘功能:</b></td><td bgcolor="#FFFFFF">
	<input type="radio" class="radio" name="Forum_setting(119)" value="1" <%If CID(team.Forum_setting(119))=1 Then%>checked<%End If%>> 是 
	<input type="radio" class="radio" name="Forum_setting(119)" value="0" <%If CID(team.Forum_setting(119))=0 Then%>checked<%End If%>> 否</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>fs2you帐号:</b><br><span class="a3"> fs2you网站提供了&lt;=1G的附件上传功能,如果您开启了此功能,就可以利用fs2you.com的空间来保存您的论坛所需要的大尺寸附件. 开启前您必须先申请fs2you的帐号. 立即到<A HREF="http://www.fs2you.com/">http://www.fs2you.com/</A>注册,然后将您的帐号填写到此处即可.  </span>
	</td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(117)" value="<%=team.Forum_setting(117)%>">
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center><br>
<a name="积分设置"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
<td colspan="2">积分设置</td>
</tr>
<tr class="a4"><td colspan="2">
	<table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
	<tr class="a4"><td><BR><Ul>
		<li>备注: ExtCredits0 为系统内置属性,不能删除。</li>
		<li>ExtCredits0 为普通用户的权限等级参照对象!</li> 
		<li>只有开启<B>交易积分设置</B>选项，论坛相应的积分功能，如工资发放，悬赏功能才可以开启!</li>
		</ul></td></tr>
	</table>
</td></tr>
<tr><td colspan="2" bgcolor="#F8F8F8">
<table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
<tr class="a1"><td colspan="7">扩展积分设置</td></tr>
<tr align="center" class="a3">
	<td>积分代号</td>
	<td>积分名称</td>
	<td>积分单位</td>
	<td>注册初始积分</td>
	<td>启用此积分</td>
	<td>在帖子中显示</td>
</tr>
<%
Dim ExtCredits,ExtSort,MustOpen,MustSort,U,M
ExtCredits= Split(team.Club_Class(21),"|")
For U=0 to Ubound(ExtCredits)
	ExtSort=Split(ExtCredits(U),",")
%>
<tr align="center">
	<td bgcolor="#F8F8F8">ExtCredits<%=U%></td>
	<td bgcolor="#FFFFFF"><input type="text" size="8" name="ExtCredits<%=U%>_0" value="<%=ExtSort(0)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="5" name="ExtCredits<%=U%>_1" value="<%=ExtSort(1)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="3" name="ExtCredits<%=U%>_2" value="<%=ExtSort(2)%>"></td>
	<td bgcolor="#FFFFFF"><input type="checkbox" name="ExtCredits<%=U%>_3" value="1" <%If ExtSort(3)=1 or U<2 Then%>checked<%End If%> onclick="findobj('policy<%=U%>').disabled=!this.checked"></td>
	<td bgcolor="#F8F8F8"><input type="checkbox" name="ExtCredits<%=U%>_4" value="1" <%If ExtSort(4)=1 or U<2 Then%>checked<%End If%>></td>
</tr>
<% Next %>
</table></td></tr>
<tr><td colspan="2" bgcolor="#F8F8F8">
<table cellspacing="1" cellpadding="4" width="100%" align="center" class="a2">
<tr class="a1"><td colspan="11">扩展积分增减策略</td></tr>
<tr align="center" class="a4">
	<td>积分代号</td>
	<td>发主题(+)</td>
	<td>回复(+)</td>
	<td>加精华(+)</td>
	<td>上传附件(+)</td>
	<td>下载附件(-)</td>
	<td>发短消息(-)</td>
	<td>搜索(-)</td>
	<td>访问推广(+)</td>
	<td>积分策略下限</td>
</tr>
<%
MustOpen = Split(team.Club_Class(22),"|")
For M=0 to Ubound(MustOpen)
	MustSort=Split(MustOpen(M),",")
%>
<tr align="center" id="policy<%=M%>" <%If Split(ExtCredits(M),",")(3)=0 Then%>disabled<%End If%>>
	<td bgcolor="#F8F8F8">Extcredits<%=M%></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_0" value="<%=MustSort(0)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_1" value="<%=MustSort(1)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_2" value="<%=MustSort(2)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_3" value="<%=MustSort(3)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_4" value="<%=MustSort(4)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_5" value="<%=MustSort(5)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_6" value="<%=MustSort(6)%>"></td>
	<td bgcolor="#F8F8F8"><input type="text" size="2" name="MustSort<%=M%>_7" value="<%=MustSort(7)%>"></td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="MustSort<%=M%>_8" value="<%=MustSort(8)%>"></td>
</tr>
<%Next%>
<tr><td colspan="11" class="a4">&nbsp;</td></tr>
<tr>
	<td class="a3" align="center">发主题(+)</td><td class="a4" colspan="10">作者发新主题增加的积分数，如果该主题被删除，作者积分也会按此标准相应减少</td>
</tr>
<tr>
	<td class="a3" align="center">回复(+)</td><td class="a4" colspan="10">作者发新回复增加的积分数，如果该回复被删除，作者积分也会按此标准相应减少</td>
</tr>
<tr>
	<td class="a3" align="center">加精华(+)</td><td class="a4" colspan="10">主题被加入精华时作者增加的积分数，如果该主题被移除精华，作者积分也会按此标准相应减少</td>
</tr>
<tr>
	<td class="a3" align="center">上传附件(+)</td><td class="a4" colspan="10">用户每上传一个附件增加的积分数，如果该附件被删除，发布者积分也会按此标准相应减少</td>
</tr>
<tr>
	<td class="a3" align="center">下载附件(-)</td><td class="a4" colspan="10">用户每下载一个附件扣除的积分数。注意: 如果允许游客组下载附件，本策略将可能被绕过</td>
</tr>
<tr>
	<td class="a3" align="center">发短消息(-)</td><td class="a4" colspan="10">用户每发送一条短消息扣除的积分数</td>
</tr>
<tr>
	<td class="a3" align="center">搜索(-)</td><td class="a4" colspan="10">用户每进行一次帖子搜索扣除的积分数</td>
</tr>
<tr>
	<td class="a3" align="center">访问推广(+)</td><td class="a4" colspan="10">访问者通过用户提供的推广链接(如 ForumAdv.asp?Uid=1)访问论坛，推广人所得的积分数</td>
</tr>
<tr>
	<td class="a3" align="center">积分策略下限</td><td class="a4" colspan="10">当用户该项积分低于此下限时，将禁止用户执行积分策略中涉及扣减此项积分的操作。例如设置为 -100，而“搜索”扣减该积分 10 个单位，则当用户该项积分小于 -100 时，将不能再执行“搜索”操作</td>
</tr>
<tr>
	<td class="a3" colspan="11">以上标明(+)的为增加的积分数，标明(-)的为减少的积分数，您也可以通过设置负值的方式变更积分的增减，各项积分增减允许的范围为 -99～+99。如果为更多的操作设置积分策略，系统就需要更频繁的更新用户积分，同时意味着消耗更多的系统资源，因此请根据实际情况酌情设置</td>
</tr></table></td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>交易积分设置:</b><br><span class="smalltxt">交易积分是一种可以由用户间自行转让、买卖交易的积分类型，您可以指定一种积分作为交易积分。如果不指定交易积分，则用户间积分交易功能将不能使用。注意: 交易积分必须是已启用的积分，一旦确定请尽量不要更改，否则以往记录及交易可能会产生问题。</span></td><td bgcolor="#FFFFFF">
		<select name="Forum_setting(99)">
			<option value="0">无</option>
			<%
			for i=1 to 7
				Response.Write "<option value="""&i&""" " 
				If Cid(team.Forum_setting(99)) = i Then Response.Write "selected"
				Response.Write ">extcredits"&i&"</option>"
			Next
			%>
		</select>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>积分评分设置:</b><br><span class="smalltxt">指定一直评分设置中需要的积分类型，用于用户的评分管理。注意: 评分使用的积分属性必须是已启用的积分，使用本积分选项目的是将论坛的交易积分与评分积分区分开，让交易积分的管理可以更合理。</span></td><td bgcolor="#FFFFFF">
		<select name="Forum_setting(46)">
			<option value="0">无</option>
			<%
			for i=1 to 7
				Response.Write "<option value="""&i&""" " 
				If Cid(team.Forum_setting(46)) = i Then Response.Write "selected"
				Response.Write ">extcredits"&i&"</option>"
			Next
			%>
		</select>
	</td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>积分交易税:</b><br><span class="smalltxt">积分交易税(损失率)为用户在利用积分进行转让、兑换、买卖时扣除的税率，范围为 0～1 之间的浮点数，例如设置为 0.2，则用户在转换 100 个单位积分时，损失掉的积分为 20 个单位，0 为不损失</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(11)" value="<%=team.Forum_setting(11)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>转账最低余额:</b><br><span class="smalltxt">积分转账后要求用户所拥有的余额最小数值。利用此功能，您可以设置较大的余额限制，使积分小于这个数值的用户无法转账；也可以将余额限制设置为负数，使得转账在限额内可以透支</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(12)" value="<%=team.Forum_setting(12)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>单主题最高出售时限(小时):</b><br><span class="smalltxt">设置当主题被作者出售时，系统允许自主题发布时间起，其可出售的最长时间。超过此时间限制后将变为普通主题，阅读者无需支付积分购买，作者也将不再获得相应收益，以小时为单位，0 为不限制</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(13)" value="<%=team.Forum_setting(13)%>"></td>
</tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center><br>
<a name="电子商务"></a>
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1"><td colspan="2">电子商务</td></tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>收款支付宝账号:</b><br><span class="smalltxt">如果开启兑换或交易功能，请填写真实有效的支付宝账号，用于收取用户以现金兑换交易积分的相关款项。如账号无效或安全码有误，将导致用户支付后无法正确对其积分账户自动充值，或进行正常的交易。</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(101)" value="<%=team.Forum_setting(101)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>支付宝安全校验码:</b><br><span class="smalltxt">如果开启兑换或交易功能，请填写与前面支付宝账户相匹配的安全校验码，如您已经设置了安全码，基于安全考虑将只显示*号。您可以在请登录支付宝后，在商家工具中的“获取安全校验码”中设置或查看安全码的内容，如账号无效或安全码有误，将导致用户支付后无法正确对其积分账户自动充值，或进行正常的交易。</span></td><td bgcolor="#FFFFFF"><input type="password" size="30" name="Forum_setting(102)" value="<%=team.Forum_setting(102)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>合作者身份ID:</b><br><span class="smalltxt">填写此ID后，支付宝交易功能才能使用。如果要获取此ID，您需要在支付宝网站申请商品交易服务。</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(103)" value="<%=team.Forum_setting(103)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>现金/积分兑换比率:</b><br><span class="smalltxt">用户购买的积分与真实货币之间的比例状况，基数为货币1元，如兑换比例为1，则1点积分＝1块货币，如兑换比例为3，则3点积分兑换1块货币。</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(104)" value="<%=team.Forum_setting(104)%>"></td>
</tr>
<tr>
	<td width="60%" bgcolor="#F8F8F8" ><b>积分兑换最低余额:</b><br><span class="smalltxt">积分兑换后要求用户所拥有的余额最小数值。利用此功能，您可以设置较大的余额限制，使积分小于这个数值的用户无法兑换。</span></td><td bgcolor="#FFFFFF"><input type="text" size="30" name="Forum_setting(105)" value="<%=team.Forum_setting(105)%>"></td>
</tr>

<tr class="a2"><td colspan="2"><img src="../images/dian.gif" align="absmiddle">  <a href="Admin_plus.asp?action=buyalipays" target="_blank"> 查看系统交易订单 </a></td></tr>

<tr class="a2"><td colspan="2"><img src="../images/dian.gif" align="absmiddle"> 快捷链接: <a href="https://www.alipay.com/user/user_register.htm" target="_blank">注册支付宝帐号</a> | <a href="https://www.alipay.com/" target="_blank">登陆支付宝</a> | <a href="http://help.alipay.com/support/index.htm/" target="_blank">支付宝客服</a></td></tr>
</table>
<br><center><input type="submit" name="settingsubmit" value="提 交"></center><br>
</form>
<a name="注册协议"></a>
<form name="myform" method="post"  action="?action=upreg">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
<tr class="a1">
	<td colspan="2">注册协议</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#F8F8F8">
		<textarea rows="10" name="myinfos" cols="60" style="overflow-y:visible;width:100%;"><%=server.htmlencode(team.Club_Class(13))%></textarea>
		<li> <B>[ 文本框支持 UBB ] </B>
		<li> <B>{$clubname}变量为论坛名称</B>
	</td>
</tr>
</table><BR><center><input type="Submit" value="发 表" name="Submit" />&nbsp;</center>
</form><br><br>
<%
End Sub

Sub SetColors(str)
	if Trim(team.Forum_setting(82))=Trim(""&str&"") Then
		Response.Write "selected"
	end if
End Sub


Sub Settingok	
	Dim ClubSystem,openclock,Co,u
	Dim ExtSort,MustSort,ExtCredits,MustOpen
	openclock=""
	For Co=0 to 23
		If openclock="" Then
			If Request.form("openclock"&Co)="1" Then
				openclock="1"
			Else
				openclock="0"
			End If
		Else
			If Request.form("openclock"&Co)="1" Then
				openclock=openclock&"*1"
			Else
				openclock=openclock&"*0"
			End If
		End If
	Next
	ClubSystem = ClubSystem & openclock &"$$$"
	If CID(Request.Form("Forum_setting(7)"))=1 Then
		If CID(Request.Form("Forum_setting(1)"))=0 Then SuccessMsg "您必须选择Email组件支持，才可以使用此功能。"
	End If
	For i=1 to 120
		If i=8 Then ClubSystem= ClubSystem & team.Forum_setting(8)
		'If i = 87 Then ClubSystem= ClubSystem & Request.Cookies("userlogininfo")
		ClubSystem= ClubSystem & Replace(Request.Form("Forum_setting("&i&")"),"$$$","")&"$$$"
	Next
	For i=0 to 7
		ExtSort=""
		ExtSort= Request.Form("ExtCredits"&i&"_0")&","&Request.Form("ExtCredits"&i&"_1")&","&Request.Form("ExtCredits"&i&"_2")
		If Request.Form("ExtCredits"&i&"_3")=1 Then
			ExtSort= ExtSort& ",1"
		Else
			ExtSort= ExtSort& ",0"
		End If
		If Request.Form("ExtCredits"&i&"_4")=1 Then
			ExtSort= ExtSort& ",1"
		Else
			ExtSort= ExtSort& ",0"
		End If
		If ExtCredits="" Then
			ExtCredits = ExtCredits & ExtSort
		Else
			ExtCredits = ExtCredits & "|" &  ExtSort
		End If
	Next
	For i=0 to 7
		MustSort=""
		For u=0 to 9
			If MustSort="" Then
				If Request.Form("MustSort"&i&"_"&u&"")="" Then
					MustSort = "0"
				Else
					MustSort = Request.Form("MustSort"&i&"_"&u&"")
				End If
			Else
				If Request.Form("MustSort"&i&"_"&u&"")="" Then
					MustSort = MustSort &",0"
				Else
					MustSort = MustSort &","&Request.Form("MustSort"&i&"_"&u&"")
				End If
			End If
		Next
		If MustOpen="" Then
			MustOpen = MustOpen & MustSort
		Else
			MustOpen = MustOpen &"|"& MustSort
		End If
	Next
	If Len(Trim(Request.Form("Club_Class(8)")))>255 Then
		SuccessMsg "管理操作及评分理由选项不能多于255个字符"
	End If
	If Len(Trim(Request.Form("Club_Class(28)")))>255 Then
		SuccessMsg "JS 来路限制不能多于255个字符"
	End If	
	If Len(Trim(Request.Form("Club_Class(7)")))>255 Then
		SuccessMsg "用户帖子过滤设置不能多于255个字符"
	End If	
	If Len(Trim(Request.Form("Club_Class(5)")))>255 Then
		SuccessMsg "发帖敏感设置不能多于255个字符"
	End If	
	If Len(Trim(Request.Form("Club_Class(6)")))>255 Then
		SuccessMsg "用户IP屏蔽设置不能多于255个字符"
	End If
	team.Execute("Update "&IsForum&"Clubconfig set Allclass='"&Replace(ClubSystem,"'","")&"',ClubName='"&Replace(Trim(Request.Form("Club_Class(1)")),"'","")&"',Cluburl='"&Replace(Trim(Request.Form("Club_Class(2)")),"'","")&"',Homename='"&Replace(Trim(Request.Form("Club_Class(3)")),"'","")&"',Homeurl='"&Replace(Trim(Request.Form("Club_Class(4)")),"'","")&"',ExtCredits='"&ExtCredits&"',MustOpen='"&Replace(MustOpen,"'","")&"',ClearMail='"&Replace(Trim(Request.Form("Club_Class(23)")),"'","")&"',ClearIP='"&Replace(Trim(Request.Form("Club_Class(24)")),"'","")&"',UserKey='"&Replace(Trim(Request.Form("Club_Class(25)")),"'","")&"',BodyMeta='"&Replace(Trim(Request.Form("Club_Class(26)")),"'","")&"',ClearPost='"&Replace(Trim(Request.Form("Club_Class(27)")),"'","")&"',Badlist='"&Replace(Trim(Request.Form("Club_Class(7)")),"'","")&"',BadWords='"&Replace(Trim(Request.Form("Club_Class(5)")),"'","")&"',Badip='"&Replace(Trim(Request.Form("Club_Class(6)")),"'","")&"',JSUrl='"&Replace(Trim(Request.Form("Club_Class(28)")),"'","")&"',Starday='"&Replace(Trim(Request.Form("Club_Class(29)")),"'","")&"',ManageText='"&Replace(Trim(Request.Form("Club_Class(8)")),"'","")&"'")
	team.Execute("update ["&IsForum&"Style] Set StyleHid=0")
	team.Execute("update ["&IsForum&"Style] Set StyleHid=1 Where ID="& CID(request.Form("Forum_setting(18)")) )
	Cache.DelCache("Club_Class")
	team.SaveLog ("基本设置更新")
	SuccessMsg "论坛基本设置更新成功!"
End Sub
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
'自动建立文件夹,需要ＦＳＯ组件支持。
Private Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	'On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & "HTML"))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & "HTML")
		End If
		If Err.Number = 0 Then
			CreatePath = PathValue & "HTML" & "/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function
Call Footer()
%>