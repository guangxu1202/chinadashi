<!--#Include File="../inc/ClsMain.asp"-->
<!--#Include File="../inc/Function.asp"-->
<%
	Dim iCacheName,iCache,mCacheName,CacheName,Forum_sn
	Dim TK_UserName,TK_Userpass,RemoteAddr
	MyDbPath = "../"
	iCacheName = Server.MapPath("admin.asp")
	iCacheName = Split(iCacheName,"\")
	For iCache = 0 To Ubound(iCacheName)-2
		mCacheName = mCacheName & iCacheName(iCache)
	Next
	CacheName = "team_" & Replace(mCacheName,":","")
	Forum_sn = Replace(CacheName,"_","")
	Set Cache = New Cls_Cache
	Set team = New Cls_Forum
	TK_UserName = DecodeCookie(team.Checkstr(Trim(Request.Cookies(Forum_sn)("UserName"))))
	TK_Userpass = team.Checkstr(Trim(Request.Cookies(Forum_sn)("Userpass")))
	RemoteAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If RemoteAddr = "" Then RemoteAddr = Request.ServerVariables("REMOTE_ADDR")
	RemoteAddr = team.Checkstr(RemoteAddr)
	team.GetForum_Setting
	team.CheckUserLogin
	'**************************************************
	'过程名：Error
	'作  用：显示错误提示信息
	'**************************************************
	'无条件转向
	Public Sub Error(Message)
		Response.Redirect "../Error.asp?Message="&SerVer.URLencode(Message)&""
	End Sub
	Public Sub Error1(Message)
		Response.Redirect "../Error.asp?Message="&SerVer.URLencode(Message)&""
	End Sub
	'弹出提示
	Public Sub Error2(Message)
		Response.Redirect "../Error.asp?Message2="&SerVer.URLencode(Message)&""
	End Sub

	Function iUbb_Code(Str)
		If Str="" Or IsNull(Str) Then Exit Function
		Dim s,re,r
		set re = New RegExp
		re.Global = True
		re.IgnoreCase = True
		s = str
		re.Pattern="\[font=([^<>\]]*?)\](.*?)\[\/font]"
		s=re.Replace(s,"<font face=""$1"">$2</font>")
		re.Pattern="\[color=([^<>\]]*?)\](.*?)\[\/color]"
		s=re.Replace(s,"<font color=""$1"">$2</font>")
		re.Pattern="\[align=([^<>\]]*?)\](.*?)\[\/align]"
		s=re.Replace(s,"<div align=""$1"">$2</div>")
		re.Pattern="\[size=(\d*?)\](.*?)\[\/size]"
		s=re.Replace(s,"<font size=""$1"">$2</font>")
		re.Pattern="\[qq\](\d*?)\[\/qq]"
		s=re.Replace(s,"<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=$1&Site=team5.cn&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:$1:5 alt=""点击这里给我发消息""></a>")
		re.Pattern="\[b\](.*?)\[\/b]"
		s=re.Replace(s,"<strong>$1</strong>")
		re.Pattern="\[p\](.*?)\[\/p]"
		s=re.Replace(s,"<p>$1</p>")
		re.Pattern="\[strike\](.*?)\[\/strike]"
		s=re.Replace(s,"<strike>$1</strike>")
		re.Pattern="\[li\](.*?)\[\/li]"
		s=re.Replace(s,"<li>$1</li>")
		re.Pattern="\[s\](.*?)\[\/s]"
		s=re.Replace(s,"<s>$1</s>")	
		re.Pattern="\[i\](.*?)\[\/i]"
		s=re.Replace(s,"<em>$1</em>")	
		re.Pattern="\[u\](.*?)\[\/u]"
		s=re.Replace(s,"<u>$1</u>")
		re.Pattern="\[p\](.*?)\[\/p]"
		s=re.Replace(s,"<p>$1</p>")
		re.Pattern="\[sub\](.*?)\[\/sub]"
		s=re.Replace(s,"<sub>$1</sub>")	
		re.Pattern="\[sup\](.*?)\[\/sup]"
		s=re.Replace(s,"<sup>$1</sup>")
		re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
		s= re.Replace(s,"<A HREF=""mailto:$2"">$2</A>")
		re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.*)(\[\/EMAIL\])"
		s= re.Replace(s,"<A HREF=""mailto:$2"">$3</A>")
		re.Pattern="\[glow\](.*?)\[\/glow]"
		s=re.Replace(s,"<span style='behavior:url(inc/font.htc)'>$1</span>")
		re.Pattern="\[URL\](.*?)\[\/URL]"
		s=re.Replace(s,"<A HREF=""$1"" TARGET=_blank>$1</A>")
		re.Pattern="(\[URL=(.[^\[]*)\])(.*?)(\[\/URL\])"
		s= re.Replace(s,"<A HREF=""$2"" TARGET=_blank>$3</A>")
		re.Pattern="\[QUOTE\](.*?)\[\/QUOTE]"
		s=re.Replace(s,"<b>QUOTE:</b><div class=""quote"">$1</div>")
		re.Pattern="\[code\](.*?)\[\/code\]"
		s=re.Replace(s,"<b>CODE:</b><div class=""code"">"&Server.HtmlEncode("$1")&"</div>")
		re.Pattern="\[sound\](.*?)\[\/sound\]"
		s=re.Replace(s,"<Img src=../images/ismp.gif alt='背景音乐播放' border=0><bgsound src=""$1"" loop=-1>")
		re.Pattern="\[em(\d*?)\]"
		s=re.Replace(s,"<Img Src=""../images/Emotions/$1.Gif"" Border=""0"" Align=""AbsMiddle"" Alt=""表情图标EM$1"">")
		re.Pattern="\[RM=*([0-9]*),*([0-9]*),*([true|false]*)\](.[^\[]*)\[\/RM]"
		s=re.Replace(s,"<object classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2""><param name=""SRC""value=""$4""><param name=""CONSOLE"" value=""$4""><param name=""CONtrOLS"" value=""imagewindow""><param name=""AUTOSTART"" value=""$3"" ></object><br/><object classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video"" width=""$1""><param name=""SRC""value=""$4""><param name=""AUTOSTART"" value=""$3""><param name=""CONtrOLS"" value=""controlpanel""><param name=""CONSOLE"" value=""$4""></object>")
		'MP-UBB
		re.Pattern="\[MP=*([0-9]*),*([0-9]*),*([true|false]*)\](.[^\[]*)\[\/MP]"
		s=re.Replace(s,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2><PARAM NAME=AUTOSTART VALUE=$3><param name=ShowStatusBar value=-1><param name=Filename value=$4><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=""$4"" width=$1 height=$2></embed></object>")	
		'FLASH
		re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(http://|ftp://|../)(.[^\[]*)(.swf)(\[\/FLASH\])"
		If team.Forum_setting(69) = 1 Then
			s= re.Replace(s,"<a href=""$4$5$6"" TARGET=_blank><IMG SRC=../images/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4$5$6""><PARAM NAME=quality VALUE=high><embed src=""$4$5$6"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4$5$6</embed></OBJECT>")
		Else
			s= re.Replace(s,"<a href=""$4$5$6"" TARGET=_blank><IMG SRC=../images/swf.gif border=0 align=absmiddle height=16 width=16>[全屏欣赏,注意Flash可能含有不安全内容]</a>")
		End If
		re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\](.*?)\[\/UPLOAD]"
		If team.Forum_setting(69) = 1 Then
			s=re.Replace(s,"<BR><A HREF=""../$2"" TARGET=_blank><IMG SRC=""../$2"" border=0 alt=""按此在新窗口浏览图片""  onmouseover=""javascript:if(this.width>520)this.width=520;"" style=""CURSOR: hand"" onload=""javascript:if(this.width>520)this.width=520;""'></A>")
		Else
			s=re.Replace(s,"<BR><A HREF=""../$2"" TARGET=_blank><IMG SRC=""../images/type/$1.gif"" border=0 alt=""按此在新窗口浏览图片""></A>")
		End if
		If team.Forum_setting(69) = 1 Then
			re.Pattern="\[img\]\s*([^\[\<\r\n]+?)\s*\[\/img\]"
			s=re.Replace(s,"<img src=""$1"" border=""0"" onload=""if(this.width>screen.width*0.7) {this.resized=true; this.width=screen.width*0.7; this.alt='Click here to open new window\\nCTRL+Mouse wheel to zoom in/out';}"" onmouseover=""if(this.width>screen.width*0.7) {this.resized=true; this.width=screen.width*0.7; this.style.cursor='hand'; this.alt='Click here to open new window\\nCTRL+Mouse wheel to zoom in/out';}"" onclick=""if(!this.resized) {return true;} else {window.open('$1');}"" onmousewheel=""return imgzoom(this);"" alt="""" />")
			re.Pattern="\[img=(\d{1,3})[x|\,](\d{1,3})\]\s*([^\[\<\r\n]+?)\s*\[\/img\]"
			s=re.Replace(s,"<img width=""$1"" height=""$2"" src=""$3"" border=""0"" alt="""" />")
		Else
			re.Pattern="\[img\]\s*([^\[\<\r\n]+?)\s*\[\/img\]"
			s=re.Replace(s,"<a href=""$1"" target=""_blank"">$1</a>")
			re.Pattern="\[img=(\d{1,3})[x|\,](\d{1,3})\]\s*([^\[\<\r\n]+?)\s*\[\/img\]"
			s=re.Replace(s,"<a href=""$1"" target=""_blank"">$1</a>")
		End If
		re.Pattern="\[UPLOAD=(txt|rar|zip)\]([0-9]*)\[\/UPLOAD]"
		If team.Group_Browse(24) = 0 Then 
			s=re.Replace(s,"<img src=""../images/type/$1.gif"" border=""0"" align=""absmiddle""> 您所在的组没有查看附件的权限。")
		Else
			s=re.Replace(s,"<img src=""../images/type/$1.gif"" border=""0"" align=""absmiddle""><A HREF=""../ShowFile.asp?ID=$2"" TARGET=""_blank"">点击浏览该文件</A>")
		End If
		re.Pattern="\[UPLOAD=(swf|swi)\](.*?)\[\/UPLOAD]"
		If team.Forum_setting(14) = 1 Then
			s=re.Replace(s,"<img src=""../images/type/$1.gif"" border=""0"" align=""absmiddle""><br><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=300></embed>")
		Else
			s=re.Replace(s,"<img src=""../images/type/$1.gif"" border=""0"" align=""absmiddle""><A HREF=""../$2"" TARGET=""_blank"">[全屏欣赏,注意Flash可能含有不安全内容]</A>")
		End if
		re.Pattern="\[UPLOAD=(.[^\[]*)\]([0-9]*)\[\/UPLOAD]"
		s=re.Replace(s,"<img src=""../images/type/$1.gif"" border=""0"" align=""absmiddle""><A HREF=""../ShowFile.asp?ID=$2"" TARGET=""_blank"">点击浏览该文件</A>")
		iUbb_Code = s
	End Function

	Function ReadPowers(s)
		Dim re
		set re = New RegExp
		re.Global = True
		re.IgnoreCase = True

		'特殊贴+购买贴
		Dim Uid,CodeRs,U,match,ts
		UID=int(Request.QueryString("tid"))
		Set CodeRs = team.Execute("Select UserName,Relist,Replies From Forum Where ID="& UID )
		If Instr(s,"[REPLAYVIEW]")>0 or Instr(s,"[replayview]")>0 Then
			IF Not CodeRs.Eof And Request.QueryString("retopicid")="" Then
				re.Pattern="\[REPLAYVIEW\][\s\n]*\[\/REPLAYVIEW\]"
				s=re.Replace(s,"")
				re.Pattern="\[\/REPLAYVIEW\]"
				s=re.replace(s, chr(1)&"/REPLAYVIEW]")
				re.Pattern="\[REPLAYVIEW\]([^\x01]*)\x01\/REPLAYVIEW\]"
				If Not team.UserLoginED Then 
					s=re.Replace(s,"<fieldset class=textquote><legend><strong>回复可见贴</strong></legend>本帖内容已被隐藏,请登陆后查看!</fieldset>")
				Else
					If tk_UserName = CodeRs(0) or team.ManageUser Or Team.Execute("Select Count(ID) From "&CodeRs(1)&" Where Topicid="&UID&" And UserName='"&TK_UserName&"'")(0)>0  Then
						s=re.Replace(s,"<fieldset class=textquote><legend><strong>回复可见贴</strong></legend>$1</fieldset>")
					Else
						s=re.Replace(s,"<fieldset class=textquote><legend><strong>回复可见贴</strong></legend>本帖内容已被隐藏,回复本帖后才可查看!</fieldset>")
					End If
				End if
			End If
		End If
		If Instr(s,"[buy=")>0 or Instr(s,"[BUY=")>0 Then
			U=0
			IF Not CodeRs.Eof And Request.QueryString("retopicid")="" Then
				re.Pattern="\[BUY=*([0-9]+)\]((.|\n)*)\[\/BUY\]"
				Set match = re.Execute(s)
				U=int(re.Replace(match.item(0),"$1"))
				If Not team.UserLoginED Then
					s=re.Replace(s,"<fieldset class=textquote><legend><strong>出售贴</strong></legend>请登陆查看此帖购买金额!</fieldset>")
				Else
					If team.ManageUser Or Trim(tk_UserName) = Trim(CodeRs(0)) Then
						s=re.Replace(s,"<fieldset class=textquote><legend><strong>出售贴</strong></legend>$2<hr class=""a3"" width=""90%""><li>因为你的等级或您是发起人,此帖售价$1元对您无效,你可以<a href=""../Command.asp?action=seebuy&buyid="& tID &""" target=""_blank""><B>查看购买者列表</B>! </a><li> 因为国家鼓励消费的缘故,也可以<a href=""../Command.asp?action=buypost&buyid="& tID &"&postname="& CodeRs(0) &"&money=$1""> <B>购买此帖</B>! </a></fieldset>")
					Else
						Set ts=Team.Execute("Select Name From ["&IsForum&"ListRec] Where  PostID="&tID) 
						If (ts.Eof and ts.Bof) Then
							s=re.Replace(s,"<fieldset class=textquote><legend><strong>出售贴</strong></legend>此帖售价$1元,你需要购买才可以浏览<hr class=""a3"" width=""90%""><li> <a href=""../Command.asp?action=buypost&buyid="& tID &"&postname="& CodeRs(0) &"&money=$1""> 购买此帖! </a></fieldset>")
						Else
							If Instr(Ts(0),TK_UserName&",") >0 Then
								s=re.Replace(s,"<fieldset class=textquote><legend><strong>出售贴</strong></legend>$2 <hr class=""a3"" width=""90%""><li>此帖售价$1元,你已经购买过<li><a href=""../Command.asp?action=seebuy&buyid="& tID &""" target=""_blank""> 查看购买者列表! </a></fieldset>")
							Else
								s=re.Replace(s,"<fieldset class=textquote><legend><strong>出售贴</strong></legend>此帖售价$1元,你需要购买才可以浏览<hr class=""a3"" width=""90%""><li> <a href=""../Command.asp?action=buypost&buyid="& tID &"&postname="& CodeRs(0) &"&money=$1""> 购买此帖! </a></fieldset>")
							End If
						End IF
					End IF
				End If
			End If
		End If
		ReadPowers = s
		Set re = Nothing
	End Function

%>