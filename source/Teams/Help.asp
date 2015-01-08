<!-- #include File="Conn.Asp" -->
<!-- #include File="Inc/Const.Asp" -->
<%
Dim X1,X2,Fid,Acc
team.Headers(Team.Club_Class(1) &" - 论坛帮助")
X2=" <A Href=Help.Asp><B>论坛帮助</B></A> "
Select Case Request("page")
	Case "custom"
		X1=" TEAM's Board 的特别使用帮助 "
		Echo Team.Menutitle
		Call custom
	Case "usermaint"
		X1=" 用户须知 "
		Echo Team.Menutitle
		Call usermaint
	Case "using"
		X1=" 论坛使用 "
		Echo Team.Menutitle
		Call using	
	Case "messages"
		X1=" 读写帖子和收发短消息 "
		Echo Team.Menutitle
		Call messages	
	Case "mise"
		X1=" 其他问题 "
		Echo Team.Menutitle
		Call mise
	Case Else
		X1="  "
		Echo Team.Menutitle
		Call Main
End Select
Team.footer

Sub Main
	Call Menu01 : Call Menu02 : Call Menu03 : Call Menu04
	If team.UserLoginED Then Call Menu05
End Sub

Sub Menu01 %>
	<div class="a2" id="center">
		<div class="a1"  style="padding: 5px;">TEAM's Board 的特别使用帮助</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=custom#0">互联网电子公告服务管理规定</a></li>
			<li><a href="Help.asp?page=custom#1">互联网信息服务管理办法</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu02 %>
	<div class="a2" id="center">
		<div class="a1"  style="padding: 5px;">用户须知</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=usermaint#1">我必须要注册吗？</a></li>
			<li><a href="Help.asp?page=usermaint#2">TEAM's 论坛使用 Cookies 吗？</a></li>
			<li><a href="Help.asp?page=usermaint#3">如何使用签名？</a></li>
			<li><a href="Help.asp?page=usermaint#4">如何使用个性化的头像？</a></li>
			<li><a href="Help.asp?page=usermaint#5">如果我遗忘了密码，我该怎么办？</a></li>
			<li><a href="Help.asp?page=usermaint#6">什么是“短消息”？</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu03 %>
	<div class="a2" id="center">
		<div class="a1" style="padding: 5px;">论坛使用</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=using#1">在哪里可以登录？</a></li>
			<li><a href="Help.asp?page=using#2">在哪里可以退出？</a></li>
			<li><a href="Help.asp?page=using#3">我要搜索论坛，应该怎么做？</a></li>
			<li><a href="Help.asp?page=using#4">怎样给其他人发送“短消息”？</a></li>
			<li><a href="Help.asp?page=using#5">怎样看到全部的会员？</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu04 %>
	<div class="a2" id="center">
		<div class="a1" style="padding: 5px;">读写帖子和收发短消息</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=messages#1">如何发布新帖子？</a></li>
			<li><a href="Help.asp?page=messages#2">如何回复帖子？</a></li>
			<li><a href="Help.asp?page=messages#3">我能够删除主题吗？</a></li>
			<li><a href="Help.asp?page=messages#4">怎样编辑自己发表的帖子？</a></li>
			<li><a href="Help.asp?page=messages#5">我可不可以上传附件？</a></li>
			<li><a href="Help.asp?page=messages#6">该怎样发起一个投票？</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu05 
	%>
	<div class="a2" id="center">
		<div class="a1" style="padding: 5px;">其他问题</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=mise#1">UBB Code的使用方法？</a></li>
			<li><a href="Help.asp?page=mise#2">普通用户如何成为版主？ </a></li>
			<li><a href="Help.asp?page=mise#3">我如何在TEAM Board 具有更多的权限？</a></li>
			<li><a href="Help.asp?page=mise#4">查看我的权限</a></li>
		</div>
	</div>
	<br><%
End Sub


Sub mise
	If Not team.UserLoginED Then 
		Call Main()
	Else
		Call Menu05
		%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">UBB Code的使用方法？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 您可以使用 TEAM's 代码--一个 HTML 代码的简化版本，来简化对帖子显示格式的控制。<br><br>
<ol type="1">
<li>[b]粗体文字 Abc[/b] &nbsp; 效果:<b>粗体文字 Abc</b> （粗体字）<br><br></li>
<li>[i]斜体文字 Abc[/i] &nbsp; 效果:<i>斜体文字 Abc</i> （斜体字）<br><br></li>
<li>[u]下划线文字 Abc[/u] &nbsp; 效果:<u>下划线文字 Abc</u> （下划线）<br><br></li>
<li>[color=red]红颜色[/color] &nbsp; 效果:<font color="red">红颜色</font> （改变文字颜色）<br><br></li>
<li>[size=3]文字大小为 3[/size] &nbsp; 效果:<font size="3">文字大小为 3</font> （改变文字大小）<br><br></li>
<li>[font=仿宋]字体为仿宋[/font] &nbsp; 效果:<font face"仿宋">字体为仿宋</font> （改变字体）<br><br></li>
<li>[align=Center]内容居中[/align] &nbsp; （格式内容位置） 效果:<br><center>内容居中</center><br></li>
<li>[url]http://www.team5.cn[/url] &nbsp; 效果:<a href="http://www.team5.cn" target="_blank">http://www.team5.cn</a> （超级连接）<br><br></li>
<li>[url=http://www.team5.cn]TEAM's 论坛[/url] &nbsp; 效果:<a href="http://www.TEAM5.cn" target="_blank">TEAM's 论坛</a> （超级连接）<br><br></li>
<li>[email]myname@mydomain.com[/email] &nbsp; 效果:<a href="mailto:myname@mydomain.com">myname@mydomain.com</a> （E-Mail 链接）<br><br></li>
<li>[email=teamserver@163.com]TEAM's 技术支持[/email] &nbsp; 效果:<a href="mailto:teamserver@163.com">TEAM's 技术支持</a> （E-Mail 链接）<br><br></li>
<li>[quote]TEAM Board 是由TEAM Studio 开发的论坛软件[/quote] &nbsp; （引用内容，类似的代码还有 [code][/code]）<br><br></li>
<li>[REPLAYVIEW]免费帐号为: username/password[/REPLAYVIEW] &nbsp; （按回复隐藏内容）<br>效果:只有当浏览者回复本帖时，才显示其中的内容，否则显示为<fieldset class=textquote><legend><strong>回复可见贴</strong></legend>本帖内容已被隐藏,请登陆后查看!</fieldset><br><br></li>
<li>[money=20]免费帐号为: username/password[/money] &nbsp; （按金币隐藏内容）<br>效果:只有当浏览者金币高于 20 点时，才显示其中的内容，否则显示为<fieldset class=textquote><legend><strong>限金钱数贴</strong></legend>本帖限金钱数大于20才可以浏览!</fieldset><br><br></li>
<li>[marquee]This is sample text[/marquee] &nbsp; (产生水平移动的效果。类似于HTML&lt;marquee&gt;标签。注意：仅在IE浏览器下可用。)<br><br></li>
<li>[qq]688888[/qq] &nbsp; (显示QQ在线状态，可以通过点击此图标和此人聊天。)<br><br></li>
<br>以下 TEAM's 代码需论坛可用 [img] 代码才能使用<hr noshade size="0" width="50%" color="#698CC3" align="left"><br>
<li>[img]http://www.team5.cn/images/default/logo.gif[/img] &nbsp; （链接图像）<br>效果:<br><img src="images/logo.gif"> <br><br></li>
<li>[flash=480,360]http://www.team5.cn/images/banner.swf[/flash]&nbsp; （链接 flash 动画，用法与 [img] 类似）<br><br></li>
</ol>
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">普通用户如何成为版主？  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 论坛的版主是自愿申请的，管理员可能会要求版主需要达到一定积分，或在论坛注册超过一定时间等。版主应该是诚实守信、乐于助人、大公无私的表率，同时还要熟悉专业，经验丰富，有良好的口碑。如果你确认已经达到上面几点，并希望担任本站的版主，可以与管理员联系。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">我如何在TEAM Board 具有更多的权限？  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 本站所使用的 TEAM 论坛是按照系统头衔和用户积分区分的，积分可以参考您的发帖量，以及管理员的评分，或两者综合来决定。当积分达到一定等级要求时，系统会自动为您开通新的权限，并给予相应等级标志。因此，拥有较高的积分数，不仅代表您在本论坛的资历与活跃程度，同时也意味着能够拥有比其他用户更多的高级权限。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">查看我的权限  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>用户组名称：<%=team.Levelname(0)%> </li>
			<li>用户名显示样式：<span Style='<%=team.Levelname(1)%>'> 会员名 </span> </li>
		</div>
		<div class="a1" style="padding: 5px;">基本权限  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>允许访问论坛：<%if Team.Group_Browse(0)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>阅读权限：<%=Team.Group_Browse(1)%></li>
			<li>允许查看用户资料：<%if Team.Group_Browse(2)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许积分转账：<%if Team.Group_Browse(3)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许使用搜索：<%if Team.Group_Browse(4)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许使用头像：<%if Team.Group_Browse(5)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许对用户评分：<%if Team.Group_Browse(10)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许使用文集：<%if Team.Group_Browse(7)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许发起投票：<%if Team.Group_Browse(8)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许发起活动：<%if Team.Group_Browse(9)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许悬赏问题：<%if Team.Group_Browse(20)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许自定义头衔：<%if Team.Group_Browse(11)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>短消息收件箱容量：<%=Team.Group_Browse(12)%></li>
		</div>
		<div class="a1" style="padding: 5px;">帖子相关选项 </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>允许发新话题：<%if Team.Group_Browse(13)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许发表回复：<%if Team.Group_Browse(14)=0 then%>禁止<%Else%>允许<%End If%></li>
			<li>允许参与投票：<%if Team.Group_Browse(15)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许发匿名贴：<%if Team.Group_Browse(17)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许设置帖子权限：<%if Team.Group_Browse(18)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许使用主题标色：<%if Team.Group_Browse(19)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许签名中使用 UBB 代码：<%if Team.Group_Browse(21)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许签名中使用 [img] 代码：<%if Team.Group_Browse(22)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>最大签名长度：<%=Team.Group_Browse(23)%> </li>
		</div>
		<div class="a1" style="padding: 5px;">附件相关选项 </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>允许下载/查看附件：<%if Team.Group_Browse(24)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>允许发布附件：<%if Team.Group_Browse(25)=0 then%>禁止<%Else%>允许<%End If%> </li>
			<li>每次上传附件个数：<%=Team.Group_Browse(26)%> </li>
			<li>最大附件尺寸(KB)：<%=Team.Group_Browse(27)%> </li>
			<li>每天上传附件的最大个数：<%=Team.Group_Browse(28)%></li>
			<li>允许附件类型：<%
				If Team.Group_Browse(29)&""="" Then
					Echo team.Forum_setting(73)
				Else
					Echo Team.Group_Browse(29)
				End if
				%> </li>
		</div>
	</div><BR>
	<%
	End If
End Sub

Sub messages
	Call Menu04()
	%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">如何发布新帖子？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 在论坛版块中，点“发布主贴”即可进入功能齐全的发帖界面。当然您也可以使用版块下面的“快速发帖”发表新帖(如果此选项打开)。注意，一般论坛都设置为需要登录后才能发帖。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">如何回复帖子？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 在浏览帖子时，点“回复帖子”即可进入功能齐全的回复界面。当然您也可以使用版块下面的“快速回复”发表回复(如果此选项打开)。注意，一般论坛都设置为需要登录后才能回复。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">我能够删除主题吗？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 论坛设置了只有拥有管理等级的用户才可以删除帖子。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">怎样编辑自己发表的帖子？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 在帖子显示的下栏用“编辑”就可以编辑自己发表的帖子。如果管理员通过论坛设置将这个功能屏蔽掉则不再可以进行此操作
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="5"></a>
		<div class="a1" style="padding: 5px;">我可不可以上传附件？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 可以。您可以在任何支持上传附件的版块中，通过发新帖、或者回复的方式上传附件（只要您的权限足够）。附件不能超过系统限定尺寸，且在可用类型的范围内才能上传。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="6"></a>
		<div class="a1" style="padding: 5px;">该怎样发起一个投票？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 您可以像发帖一样在版块中发起投票。每行输入一个可能的选项（最多10个），您可以通过阅读这个投票帖选出自己的答案，每人只能投票一次，之后将不能再对您的选择做出更改。<br><br>&nbsp; &nbsp; 管理员拥有随时关闭和修改投票选项的权力。
		</div>
	</div><BR>
	<%
End Sub


Sub using
	Call Menu03()
	%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">在哪里可以登录？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 如果您尚未登录，点击左上角的“登录”，输入用户名和密码，确定即可。如果需要保持登录，请选择相应的 Cookie 时间，在此时间范围内您可以不必输入密码而保持上次的登录状态。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">在哪里可以退出？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 如果您已经登录，点击左上角的“退出”，系统会清除 Cookie，退出登录状态。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">我要搜索论坛，应该怎么做？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 点击上面的 <a href="search.asp">搜索</a>，输入搜索的关键字并选择一个范围，就可以检索到您有权限访问论坛中的相关的帖子。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">怎样给其他人发送“短消息”？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 如果您已登录，菜单上会显示出 <a href="Message.asp" target="_blank">短信服务</a>  项，以及在帖子显示栏上面显示"短消息"的图片框, 点击后弹出短消息窗口，通过类似发送邮件一样的填写，点“发送”，消息就被发到对方收件箱中了。当他/她访问论坛的主要页面时，系统都会提示他/她收信息。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="5"></a>
		<div class="a1" style="padding: 5px;">怎样看到全部的会员？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 您可以通过点击 <a href="ShowBBS.asp">排行榜 </a> 查看所有的会员及其资料，并可实现会员资料的排序输出。
		</div>
	</div><BR>
	<%
End Sub



Sub usermaint
	Call Menu02()
	%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">我必须要注册吗？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 这取决于管理员如何设置 TEAM 论坛的用户组权限选项，您甚至有可能必须在注册成正式用户后后才能浏览帖子。当然，在通常情况下，您至少应该是正式用户才能发新帖和回复已有帖子。请 <a href="Reg.asp">点击这里</a> 免费注册成为我们的新用户！<br><br>&nbsp; &nbsp; 强烈建议您注册，这样会得到很多以游客身份无法实现的功能。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">TEAM's 论坛使用 Cookies 吗？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; TEAM 采用 Session+Cookie 的双重方式保存用户信息，以确保在各种环境，包括 Cookie 完全无法使用的情况下您都能正常使用论坛各项功能。但 Cookies 的使用仍然可以为您带来一系列的方便和好处，因此我们强烈建议您在正常情况下不要禁止 Cookie 的应用，TEAM's 的安全设计将全力保证您的资料安全。<br><br>&nbsp; &nbsp; 在登录页面中，您可以选择 Cookie 记录时间，在该时间范围内您打开浏览器访问论坛将始终保持您上一次访问时的登录状态，而不必每次都输入密码。但出于安全考虑，如果您在公共计算机访问论坛，建议选择“浏览器进程”，或在离开公共计算机前选择“退出”(<a href="Login.asp?menu=out">点击这里</a> 退出论坛)以杜绝资料被非法使用的可能。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">如何使用签名？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 签名是加在您发表的帖子下面的小段文字，注册之后，您就可以设置自己的个性签名了。<br><br>&nbsp; &nbsp; <a href="EditProfile.asp?menu=index">点击这里</a> 进入控制面板 - 资料修改，在签名框中输入签名文字，并确定不要超过管理员设置的相关限制(如字数、贴图等)，这样系统会自动选中您登录后发帖页面的显示签名选项，您的的签名将在帖子中自动被显示。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">如何使用个性化的头像？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; 同样在 <a href="EditProfile.asp?menu=index">控制面板</a>  - 资料修改 中，有一处“头像”选项。头像是显示在您用户名下面的小图像，使用头像可能需要一定的权限，否则将不会显示出来。详情请查询<a href="?page=usergroup">本论坛的级别设定</a>。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="5"></a>
		<div class="a1" style="padding: 5px;">如果我遗忘了密码，我该怎么办？ </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; TEAM 提供发送取回密码链接到 Email 的服务，点击登录页面中的 <a href="Modification.asp">取回密码</a> 功能，可以为您把取回密码的方法发送到注册时填写的 Email 信箱中。如果您的 Email 已失效或无法收到信件，请与论坛管理员联系。
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="6"></a>
		<div class="a1" style="padding: 5px;">什么是“短消息”？  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; “短消息”是论坛注册用户间交流的工具，信息只有发件和收件人可以看到，收到信息后系统会出现铃声和相应提示，您可以通过短消息功能与同一论坛上的其他用户保持私人联系。<a href="Message.asp" target="_blank">收件箱</a> 或 <a href="EditProfile.asp">控制面板</a> 中提供了短消息的收发服务。
		</div>
	</div><BR>
	<%
End Sub


Sub custom	
	Call Menu01()
	%>
	<div class="a2" id="center">
		<a name="0"></a>
		<div class="a1" style="padding: 5px;">联网电子公告服务管理规定</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			中华人民共和国信息产业部第三号令 <br />
<br />
《互联网电子公告服务管理规定》已经2000年10月8日第四次部务会议通过，现予发布，自发布之日起施行。<br />
<br />
信息产业部部长吴基传<br />
<br />
第一条 为了加强对互联网电子公告服务(以下简称电子公告服务)的管理，规范电子公告信息发布行为，维护国家安全和社会稳定，保障公民、法人和其他组织的合法权益，根据《互联网信息服务管理办法》的规定，制定本规定。<br />
第二条 在中华人民共和国境内开展电子公告服务和利用电子公告发布信息，适用本规定。<br />
　　本规定所称电子公告服务，是指在互联网上以电子布告牌、电子白板、电子论坛、网络聊天室、留言板等交互形式为上网用户提供信息发布条件的行为。<br />
第三条 电子公告服务提供者开展服务活动，应当遵守法律、法规，加强行业自律，接受信息产业部及省、自治区、直辖市电信管理机构和其他有关主管部门依法实施的监督检查。 <br />
第四条 上网用户使用电子公告服务系统，应当遵守法律、法规，并对所发布的信息负责。<br />
第五条 从事互联网信息服务，拟开展电子公告服务的，应当在向省、自治区、直辖市电信管理机构或者信息产业部申请经营性互联网信息服务许可或者办理非经营性互联网信息服务备案时，提出专项申请或者专项备案。<br />
　　省、自治区、直辖市电信管理机构或者信息产业部经审查符合条件的，应当在规定时间内连同互联网信息服务一并予以批准或者备案，并在经营许可证或备案文件中专项注明；不符合条件的，不予批准或者不予备案，书面通知申请人并说明理由。 <br />
第六条 开展电子公告服务，除应当符合《互联网信息服务管理办法》规定的条件外，还应当具备下列条件：<br />
　　(一)有确定的电子公告服务类别和栏目；<br />
　　(二)有完善的电子公告服务规则；<br />
　　(三)有电子公告服务安全保障措施，包括上网用户登记程序、上网用户信息安全管理制度、技术保障设施；<br />
　　(四)有相应的专业管理人员和技术人员，能够对电子公告服务实施有效管理。<br />
第七条 已取得经营许可或者已履行备案手续的互联网信息服务提供者，拟开展电子公告服务的，应当向原许可或者备案机关提出专项申请或者专项备案。<br />
　　省、自治区、直辖市电信管理机构或者信息产业部，应当自收到专项申请或者专项备案材料之日起60日内进行审查完毕。经审查符合条件的，予以批准或者备案，并在经营许可证或备案文件中专项注明；不符合条件的，不予批准或者不予备案，书面通知申请人并说明理由。<br />
第八条 未经专项批准或者专项备案手续，任何单位或者个人不得擅自开展电子公告服务。<br />
第九条 任何人不得在电子公告服务系统中发布含有下列内容之一的信息：<br />
　　(一)反对宪法所确定的基本原则的；<br />
　　(二)危害国家安全，泄露国家秘密，颠覆国家政权，破坏国家统一的；<br />
　　(三)损害国家荣誉和利益的；<br />
　　(四)煽动民族仇恨、民族歧视，破坏民族团结的；<br />
　　(五)破坏国家宗教政策，宣扬邪教和封建迷信的；<br />
　　(六)散布谣言，扰乱社会秩序，破坏社会稳定的；<br />
　　(七)散布淫秽、色情、赌博、暴力、凶杀、恐怖或者教唆犯罪的；<br />
　　(八)侮辱或者诽谤他人，侵害他人合法权益的；<br />
　　(九)含有法律、行政法规禁止的其他内容的； <br />
第十条 电子公告服务提供者应当在电子公告服务系统的显著位置刊载经营许可证编号或者备案编号、电子公告服务规则，并提示上网用户发布信息需要承担的法律责任。<br />
第十一条 电子公告服务提供者应当按照经批准或者备案的类别和栏目提供服务，不得超出类别或者另设栏目提供服务。<br />
第十二条 电子公告服务提供者应当对上网用户的个人信息保密，未经上网用户同意不得向他人泄露，但法律另有规定的除外。<br />
第十三条 电子公告服务提供者发现其电子公告服务系统中出现明显属于本办法第九条所列的信息内容之一的，应当立即删除，保存有关记录，并向国家有关机关报告。 <br />
第十四条 电子公告服务提供者应当记录在电子公告服务系统中发布的信息内容及其发布时间、互联网地址或者域名。记录备份应当保存60日，并在国家有关机关依法查询时，予以提供。 <br />
第十五条 互联网接入服务提供者应当记录上网用户的上网时间、用户帐号、互联网地址或者域名、主叫电话号码等信息，记录备份应保存60日，并在国家有关机关依法查询时，予以提供。<br />
第十六条 违反本规定第八条、第十一条的规定，擅自开展电子公告服务或者超出经批准或者备案的类别、栏目提供电子公告服务的，依据《互联网信息服务管理办法》第十九条的规定处罚。<br />
第十七条 在电子公告服务系统中发布本规定第九条规定的信息内容之一的，依据《互联网信息服务管理办法》第二十条的规定处罚。<br />
第十八条 违反本规定第十条的规定，未刊载经营许可证编号或者备案编号、未刊载电子公告服务规则或者未向上网用户作发布信息需要承担法律责任提示的，依据《互联网信息服务管理办法》第二十二条的规定处罚。 <br />
第十九条 违反本规定第十二条的规定，未经上网用户同意，向他人非法泄露上网用户个人信息的，由省、自治区、直辖市电信管理机构责令改正；给上网用户造成损害或者损失的，依法承担法律责任。 <br />
第二十条 未履行本规定第十三条、第十四条、第十五条规定的义务的，依据《互联网信息服务管理办法》第二十一条、第二十三条的规定处罚。<br />
第二十一条 在本规定施行以前已开展电子公告服务的，应当自本规定施行之日起60日内，按照本规定办理专项申请或者专项备案手续。<br />
第二十二条 本规定自发布之日起施行。
</div>
<a name="1"></a>
<div class="a1" style="padding: 5px;">互联网信息服务管理办法</div>
<div class="a4" style="padding: 5px;">
<ul style="margin-top: 2px">
中华人民共和国国务院令（第292号）<br />
&nbsp; &nbsp; 第一条 为了规范互联网信息服务活动，促进互联网信息服务健康有序发展，制定本办法。 <br />
&nbsp; &nbsp; 第二条 在中华人民共和国境内从事互联网信息服务活动，必须遵守本办法。 <br />
&nbsp; &nbsp; 本办法所称互联网信息服务，是指通过互联网向上网用户提供信息的服务活动。 <br />
&nbsp; &nbsp; 第三条 互联网信息服务分为经营性和非经营性两类。 <br />
&nbsp; &nbsp; 经营性互联网信息服务，是指通过互联网向上网用户有偿提供信息或者网页制作等服务活动。 <br />
&nbsp; &nbsp; 非经营性互联网信息服务，是指通过互联网向上网用户无偿提供具有公开性、共享性信息的服务活动。 <br />
&nbsp; &nbsp; 第四条 国家对经营性互联网信息服务实行许可制度；对非经营性互联网信息服务实行备案制度。 <br />
&nbsp; &nbsp; 未取得许可或者未履行备案手续的，不得从事互联网信息服务。 <br />
&nbsp; &nbsp; 第五条 从事新闻、出版、教育、医疗保健、药品和医疗器械等互联网信息服务，依照法律、行政法规以及国家有关规定须经有关主管部门审核同意的，在申请经营许可或者履行备案手续前，应当依法经有关主管部门审核同意。 <br />
&nbsp; &nbsp; 第六条 从事经营性互联网信息服务，除应当符合《中华人民共和国电信条例》规定的要求外，还应当具备下列条件： <br />
&nbsp; &nbsp; （一）有业务发展计划及相关技术方案； <br />
&nbsp; &nbsp; （二）有健全的网络与信息安全保障措施，包括网站安全保障措施、信息安全保密管理制度、用户信息安全管理制度； <br />
&nbsp; &nbsp; （三）服务项目属于本办法第五条规定范围的，已取得有关主管部门同意的文件。 <br />
&nbsp; &nbsp; 第七条 从事经营性互联网信息服务，应当向省、自治区、直辖市电信管理机构或者国务院信息产业主管部门申请办理互联网信息服务增值电信业务经营许可证（以下简称经营许可证）。 <br />
&nbsp; &nbsp; 省、自治区、直辖市电信管理机构或者国务院信息产业主管部门应当自收到申请之日起60日内审查完毕，作出批准或者不予批准的决定。予以批准的，颁发经营许可证；不予批准的，应当书面通知申请人并说明理由。 <br />
&nbsp; &nbsp; 申请人取得经营许可证后，应当持经营许可证向企业登记机关办理登记手续。 <br />
&nbsp; &nbsp; 第八条 从事非经营性互联网信息服务，应当向省、自治区、直辖市电信管理机构或者国务院信息产业主管部门办理备案手续。办理备案时，应当提交下列材料： <br />
&nbsp; &nbsp; （一）主办单位和网站负责人的基本情况； <br />
&nbsp; &nbsp; （二）网站网址和服务项目； <br />
&nbsp; &nbsp; （三）服务项目属于本办法第五条规定范围的，已取得有关主管部门的同意文件。 <br />
&nbsp; &nbsp; 省、自治区、直辖市电信管理机构对备案材料齐全的，应当予以备案并编号。 <br />
&nbsp; &nbsp; 第九条 从事互联网信息服务，拟开办电子公告服务的，应当在申请经营性互联网信息服务许可或者办理非经营性互联网信息服务备案时，按照国家有关规定提出专项申请或者专项备案。 <br />
&nbsp; &nbsp; 第十条 省、自治区、直辖市电信管理机构和国务院信息产业主管部门应当公布取得经营许可证或者已履行备案手续的互联网信息服务提供者名单。 <br />
&nbsp; &nbsp; 第十一条 互联网信息服务提供者应当按照经许可或者备案的项目提供服务，不得超出经许可或者备案的项目提供服务。 <br />
&nbsp; &nbsp; 非经营性互联网信息服务提供者不得从事有偿服务。 <br />
&nbsp; &nbsp; 互联网信息服务提供者变更服务项目、网站网址等事项的，应当提前30日向原审核、发证或者备案机关办理变更手续。 <br />
&nbsp; &nbsp; 第十二条 互联网信息服务提供者应当在其网站主页的显著位置标明其经营许可证编号或者备案编号。 <br />
&nbsp; &nbsp; 第十三条 互联网信息服务提供者应当向上网用户提供良好的服务，并保证所提供的信息内容合法。 <br />
&nbsp; &nbsp; 第十四条 从事新闻、出版以及电子公告等服务项目的互联网信息服务提供者，应当记录提供的信息内容及其发布时间、互联网地址或者域名；互联网接入服务提供者应当记录上网用户的上网时间、用户帐号、互联网地址或者域名、主叫电话号码等信息。 <br />
&nbsp; &nbsp; 互联网信息服务提供者和互联网接入服务提供者的记录备份应当保存60日，并在国家有关机关依法查询时，予以提供。 <br />
&nbsp; &nbsp; 第十五条 互联网信息服务提供者不得制作、复制、发布、传播含有下列内容的信息： <br />
&nbsp; &nbsp; （一）反对宪法所确定的基本原则的； <br />
&nbsp; &nbsp; （二）危害国家安全，泄露国家秘密，颠覆国家政权，破坏国家统一的； <br />
&nbsp; &nbsp; （三）损害国家荣誉和利益的； <br />
&nbsp; &nbsp; （四）煽动民族仇恨、民族歧视，破坏民族团结的； <br />
&nbsp; &nbsp; （五）破坏国家宗教政策，宣扬邪教和封建迷信的； <br />
&nbsp; &nbsp; （六）散布谣言，扰乱社会秩序，破坏社会稳定的； <br />
&nbsp; &nbsp; （七）散布淫秽、色情、赌博、暴力、凶杀、恐怖或者教唆犯罪的； <br />
&nbsp; &nbsp; （八）侮辱或者诽谤他人，侵害他人合法权益的； <br />
&nbsp; &nbsp; （九）含有法律、行政法规禁止的其他内容的。 <br />
&nbsp; &nbsp; 第十六条 互联网信息服务提供者发现其网站传输的信息明显属于本办法第十五条所列内容之一的，应当立即停止传输，保存有关记录，并向国家有关机关报告。 <br />
&nbsp; &nbsp; 第十七条 经营性互联网信息服务提供者申请在境内境外上市或者同外商合资、合作，应当事先经国务院信息产业主管部门审查同意；其中，外商投资的比例应当符合有关法律、行政法规的规定。 <br />
&nbsp; &nbsp; 第十八条 国务院信息产业主管部门和省、自治区、直辖市电信管理机构，依法对互联网信息服务实施监督管理。 <br />
&nbsp; &nbsp; 新闻、出版、教育、卫生、药品监督管理、工商行政管理和公安、国家安全等有关主管部门，在各自职责范围内依法对互联网信息内容实施监督管理。 <br />
&nbsp; &nbsp; 第十九条 违反本办法的规定，未取得经营许可证，擅自从事经营性互联网信息服务，或者超出许可的项目提供服务的，由省、自治区、直辖市电信管理机构责令限期改正，有违法所得的，没收违法所得，处违法所得3倍以上5倍以下的罚款；没有违法所得或者违法所得不足5万元的，处10万元以上100万元以下的罚款；情节严重的，责令关闭网站。 <br />
&nbsp; &nbsp; 违反本办法的规定，未履行备案手续，擅自从事非经营性互联网信息服务，或者超出备案的项目提供服务的，由省、自治区、直辖市电信管理机构责令限期改正；拒不改正的，责令关闭网站。 <br />
&nbsp; &nbsp; 第二十条 制作、复制、发布、传播本办法第十五条所列内容之一的信息，构成犯罪的，依法追究刑事责任；尚不构成犯罪的，由公安机关、国家安全机关依照《中华人民共和国治安管理处罚条例》、《计算机信息网络国际联网安全保护管理办法》等有关法律、行政法规的规定予以处罚；对经营性互联网信息服务提供者，并由发证机关责令停业整顿直至吊销经营许可证，通知企业登记机关；对非经营性互联网信息服务提供者，并由备案机关责令暂时关闭网站直至关闭网站。 <br />
&nbsp; &nbsp; 第二十一条 未履行本办法第十四条规定的义务的，由省、自治区、直辖市电信管理机构责令改正；情节严重的，责令停业整顿或者暂时关闭网站。 <br />
&nbsp; &nbsp; 第二十二条 违反本办法的规定，未在其网站主页上标明其经营许可证编号或者备案编号的，由省、自治区、直辖市电信管理机构责令改正，处5000元以上5万元以下的罚款。 <br />
&nbsp; &nbsp; 第二十三条 违反本办法第十六条规定的义务的，由省、自治区、直辖市电信管理机构责令改正；情节严重的，对经营性互联网信息服务提供者，并由发证机关吊销经营许可证，对非经营性互联网信息服务提供者，并由备案机关责令关闭网站。 <br />
&nbsp; &nbsp; 第二十四条 互联网信息服务提供者在其业务活动中，违反其他法律、法规的，由新闻、出版、教育、卫生、药品监督管理和工商行政管理等有关主管部门依照有关法律、法规的规定处罚。 <br />
&nbsp; &nbsp; 第二十五条 电信管理机构和其他有关主管部门及其工作人员，玩忽职守、滥用职权、徇私舞弊，疏于对互联网信息服务的监督管理，造成严重后果，构成犯罪的，依法追究刑事责任；尚不构成犯罪的，对直接负责的主管人员和其他直接责任人员依法给予降级、撤职直至开除的行政处分。 <br />
&nbsp; &nbsp; 第二十六条 在本办法公布前从事互联网信息服务的，应当自本办法公布之日起60日内依照本办法的有关规定补办有关手续。 <br />
&nbsp; &nbsp; 第二十七条 本办法自公布之日起施行。
		</div>
	</div>
	<br>
<%
End Sub

%>
