<%
	'**************************************************
	'��������UBB_Code
	'��  �ã�UBB����ת��
	'��  ����str   ----��ת�����ַ�
	'����ֵ��ת������ַ�
	'**************************************************
	Const MaxLoopcount=100
	Function UBB_Code(Str)
		If Str="" Or IsNull(Str) Then Exit Function
		Dim s,re,r,smode
		set re = New RegExp
		smode = 0
		re.Global = True
		re.IgnoreCase = True
		s = str
		If Not Request("newpage") = "edit" Then 
			smode = 1
		End If
		If smode = 1 Then
			re.Pattern="\javascript"
			s=re.Replace(s,"<span>java</span>script")
			re.Pattern="alert"
			s=re.Replace(s,"<span>alert</span>")
			re.Pattern="script"
			s=re.Replace(s,"<span>script</span>")
			re.Pattern="document"
			s=re.Replace(s,"<span>document</span>")
			re.Pattern="meta"
			s=re.Replace(s,"<span>meta</span>")
			re.Pattern="expression"
			s=re.Replace(s,"<span>expression</span>")
			re.Pattern="xss"
			s=re.Replace(s,"<span>xss</span>")
			re.Pattern="\[b\]((.|\n)*?)\[\/b\]"
			s=re.Replace(s,"<strong>$1</strong>")
			re.Pattern="\[code\]((.|\n)*?)\[\/code]"
			s=re.Replace(s,"<div class=""quote""><b>������ʾ:</b><BR>$1</div>")
			'��������
			If Instr(s,"[ipass=")>0 or Instr(s,"[IPASS=")>0 Then
				Dim u,match,i
				re.Pattern="\[ipass=(.*?)\]((.|\n)*)\[\/ipass\]"
				Set match = re.Execute(s)
				For i= 0 to  match.count-1
					u = re.Replace(match.item(i),"$1")
				Next
				Response.cookies("ipass" & tid) = u
				Response.cookies("ipass" & tid).expires = Date+30
				If Request.Cookies("ulookpass"& tid) = "ok" Then
					s=re.Replace(s,"<fieldset class=textquote><legend><strong>����鿴��</strong></legend>$2</fieldset>")
				Else
					s=re.Replace(s,"<fieldset class=textquote><legend><strong>����鿴��</strong></legend>����������: <input type=text id=""ipass"" onchange=""getpass(this.value,"& tid &")"">  -- >> ȷ��?  </fieldset>")
				End if
			End If
			'��������
			re.Pattern="\[li\]((.|\n)*?)\[\/li]"
			s=re.Replace(s,"<li>$1</li>")
			re.Pattern="\[marquee\]((.|\n)*?)\[\/marquee]"
			s=re.Replace(s,"<marquee width=""90%"" scrollamount=""3"">$1</marquee>")
			re.Pattern="\[qq\](\d+?)\[\/qq]"
			s=re.Replace(s,"<a target=""blank"" href=""http://wpa.qq.com/msgrd?V=1&Uin=$1&Site=team5.cn&Menu=yes""><img border=""0"" SRC=""http://wpa.qq.com/pa?p=1:$1:5"" alt=""���������ҷ���Ϣ""></a>")
			re.Pattern="\[QUOTE\]((.|\n)*?)\[\/QUOTE]"
			s=re.Replace(s,"<div class=""quote""><b>����:</b><br>$1</div>")
			re.Pattern="\[sound\]((.|\n)*?)\[\/sound\]"
			s=re.Replace(s,"<Img src='images/ismp.gif' alt='�������ֲ���' border='0'><bgsound src='$1' loop='-1'>")
			re.Pattern="\[mp3\]((.|\n)*?)\[\/mp3\]"
			s=re.Replace(s,"<Img src='images/ismp.gif' alt='�������ֲ���' border='0'><bgsound src='$1' loop='-1'>")
			re.Pattern="\[em=*([0-9]*)\]"
			s=re.Replace(s,"<Img Src=""images/Emotions/$1.Gif"" Border=""0"" Align=""AbsMiddle"" Alt=""����ͼ��EM$1"">")
			re.Pattern="\[url=www.([^\[\'']+?)\](.+?)\[\/url\]"
			s=re.Replace(s,"<a href=""http://www.$1"" target=""_blank"">$2</a>")
			re.Pattern="\[url=(https?|ftp|gopher|news|telnet|rtsp|mms|callto|bctp|ed2k){1}:\/\/([^\[\'']+?)\]([\s\S]+?)\[\/url\]"
			s=re.Replace(s,"<a href=""$1://$2"" target=""_blank"">$3</a>")
			re.Pattern="\[url=(.*?)\](.*?)\[\/url\]"
			s=re.Replace(s,"<a href=""$1"" target=""_blank"">$2</a>")
			re.Pattern="\[email\](.*?)\[\/email\]"
			s=re.Replace(s,"<a href=""mailto:$1"" target=""_blank"">$2</a>")
			re.Pattern="\[email=(.[^\[]*)\](.*?)\[\/email\]"
			s=re.Replace(s,"<a href=""mailto:$1"" target=""_blank"">$2</a>>")
			re.Pattern="\[color=([^\[\<]+?)\]((.|\n)*?)\[\/color]"
			s=re.Replace(s,"<font color=""$1"">$2</font>")
			re.Pattern="\[font=([^\[\<]+?)\]((.|\n)*?)\[\/font]"
			s=re.Replace(s,"<font face=""$1"">$2</font>")
			re.Pattern="\[size=(\d+(\.\d+)?(px|pt|in|cm|mm|pc|em|ex|%)+?)\]((.|\n)*?)\[\/size]"
			s=re.Replace(s,"<font style=""font-size:$1"">$4</font>")
			re.Pattern="\[size=(\d+?)\]((.|\n)*?)\[\/size]"
			s=re.Replace(s,"<font size=""$1"">$2</font>")
			re.Pattern="\[align=(left|center|right)\]((.|\n)*?)\[\/align]"
			s=re.Replace(s,"'<p align=""$1"">$2</p>")
			re.Pattern="\[align=([^\[\<]+?)\]((.|\n)*?)\[\/align]"
			s=re.Replace(s,"'<br style=""clear: both""><span style=""float: $1;"">$2</span")


			'==============================================================
			re.Pattern="\[fieldset\]((.|\n)*?)\[\/fieldset]"
			s=re.Replace(s,"<br style=""clear: both""><p><fieldset class=""fieldset""><legend style=""text-align: center;"">����������ּ�¼</legend>$1</fieldset></p>")
			re.Pattern="\[legend\]((.|\n)*?)\[\/legend]"
			s=re.Replace(s,"")
			If InStr(s,"payto:") = 0 Then
				s = Replace(s,"https://www.alipay.com/payt","https://www.alipay.com/payto:")
			End If
			s=TM_Alipay_PayTo(s)

			'======================  CC  ======================
			re.Pattern="\[cc\]((.|\n)*?)\[\/cc]"
			s=re.Replace(s,"<!-- cc��Ƶ�������/by team board --><object title=""teams"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""  width=""500"" height=""400"">"&_
			"<param name=""movie"" value=""http://union.bokecc.com/$1"" /><PARAM NAME=""AllowScriptAccess"" VALUE=""never""><param name=""quality"" value=""high"" />"&_
			"<embed title=""teams"" src=""http://union.bokecc.com/$1"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""500"" height=""400"">$1</embed></object><!-- cc��Ƶ�������/by team board -->")
			'======================P2P-UBB======================
			re.Pattern="\[P2P=*([0-9]*),*([0-9]*),*([true|false]*)\](.[^\[]*)\[\/P2P]"
			s=re.Replace(s,"<EMBED name=RealObj src=test.rpm width=$1 height=$2 MAINTAINASPECT=""true"" CONTROLS=""ImageWindow"" CONSOLE=""one""></embed><BR><OBJECT id=WMPObj style=""DISPLAY: none"" codeBase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 type=application/x-oleobject height=300 standby=""Loading Microsoft Windows Media Player components..."" width=$1 classid=CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6 viewastext=""""><PARAM NAME=""URL"" VALUE=""""><PARAM NAME=""animationatStart"" VALUE=""true""><PARAM NAME=""autoStart"" VALUE=""$3""><PARAM NAME=""showControls"" VALUE=""true""></OBJECT><OBJECT id=TEAM codeBase=INC/BoBo_ActiveX_V3.ocx height=22 width=$1 classid=clsid:EC0978ED-24E3-403C-AB7A-060E388553E6><PARAM NAME=""BGColor1"" VALUE=""#000000""><PARAM NAME=""MaxLinkCount"" VALUE=""1000""><PARAM NAME=""MinHTTPPort"" VALUE=""26888""><PARAM NAME=""MaxCacheTimeS"" VALUE=""120""><PARAM NAME=""MinCacheTimeS"" VALUE=""60""><PARAM NAME=""MaxCacheSizeMB"" VALUE=""300""><PARAM NAME=""MaxDownloadKbps"" VALUE=""350""><PARAM NAME=""MaxUploadKbps"" VALUE=""0""><PARAM NAME=""MediaPlayerDelay"" VALUE=""2000""><PARAM NAME=""ActName"" VALUE=""$4""></OBJECT><BR><EMBED style=""DISPLAY: inline"" src=real.rpm width=$1 height=30 MAINTAINASPECT=""true"" CONTROLS=""StatusBar"" CONSOLE=""one""></embed>")
			'======================RM-UBB======================
			re.Pattern="\[RM=*([0-9]*),*([0-9]*),*([true|false]*)\](.[^\[]*)\[\/RM]"
			s=re.Replace(s,"<object classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2""><param name=""SRC""value=""$4""><param name=""CONSOLE"" value=""$4""><param name=""CONtrOLS"" value=""imagewindow""><param name=""AUTOSTART"" value=""$3"" ></object><br/><object classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video"" width=""$1""><param name=""SRC""value=""$4""><param name=""AUTOSTART"" value=""$3""><param name=""CONtrOLS"" value=""controlpanel""><param name=""CONSOLE"" value=""$4""></object>")
			'======================MP-UBB======================
			re.Pattern="\[MP=*([0-9]*),*([0-9]*),*([true|false]*)\](.[^\[]*)\[\/MP]"
			s=re.Replace(s,"<object align='middle' classid='CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95' class='OBJECT' id='MediaPlayer' width='$1' height='$2'><PARAM NAME='AUTOSTART' VALUE='$3'><param name='ShowStatusBar' value=-1><param name=Filename value=$4><embed type=application/x-oleobject codebase='http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701' flename='mp' src='$4' width='$1' height='$2'></embed></object>")	
			'======================FLASH======================
			re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(http://|ftp://|../)(.[^\[]*)(\[\/FLASH\])"
			If team.Forum_setting(69) = 1 Then
				s= re.Replace(s,"<a href=""$4$5"" TARGET=""_blank""><IMG SRC=""images/type/swf.gif"" border=""0"" alt=""������´������͸�FLASH����!"" height=""16"" width=""16"">[ȫ������]</a><br><OBJECT codeBase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" width=""$2"" height=""$3""><PARAM NAME=""movie"" VALUE=""$4$5""><PARAM NAME=""quality"" VALUE=""high""><embed src=""$4$5"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""$2"" height=""$3"">$4$5</embed></OBJECT>")
			Else
				s= re.Replace(s,"<a href=""$4$5"" TARGET=""_blank""><IMG SRC=""images/type/swf.gif"" border=""0"" align=""bsmiddle"" height=""16"" width=""16"">[ȫ������,ע��Flash���ܺ��в���ȫ����]</a>")
			End If
			'=======================QVOD============================================
			re.Pattern="\[QVOD\]((.|\n)*?)\[\/QVOD]"
			s=re.Replace(s,"<object classid='clsid:F3D0D36F-23F8-4682-A195-74C92B03D4AF' width='600' height='480' id='QvodPlayer' name='QvodPlayer' onError=if(window.confirm('�����Ȱ�װQvodPlayer���,Ȼ��ˢ�±�ҳ�ſ�����������.')){window.open('http://www.qvod.com/download.htm')}else{self.location='http://www.qvod.com/play.htm?setup=0&qvod=$1'}><PARAM NAME='URL' VALUE='$1'><PARAM NAME='AutoPlay' VALUE='1'></object>")
			'==============================Upload===================================
			re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\]((.|\n)*?)\[\/UPLOAD]"
			If team.Forum_setting(69) = 1 Then
				s=re.Replace(s,"<BR><A HREF=""$2"" TARGET=""_blank"" rel=""lytebox[vacation]""><IMG SRC=""$2"" border=0 alt=""�������´������ͼƬ""  onmouseover=""javascript:if(this.width>520)this.width=520;"" style=""CURSOR: hand"" onload=""javascript:if(this.width>520)this.width=520;""'></A>")
			Else
				s=re.Replace(s,"<BR><A HREF=""$2"" TARGET=_blank><IMG SRC=""images/type/$1.gif"" border=0 alt=""�������´������ͼƬ""></A>")
			End If
			If team.Forum_setting(69) = 1 Then
				re.Pattern="\[img\]((.|\n)*?)\[\/img\]"
				s=re.Replace(s,"<img src=""$1"" border=""0"" alt="""" />")
				re.Pattern="\[img=(\d{1,3})[x|\,](\d{1,3})\]\s*([^\[\<\r\n]+?)\s*\[\/img\]"
				s=re.Replace(s,"<img width=""$1"" height=""$2"" src=""$3"" border=""0"" alt="""" />")
			Else
				re.Pattern="\[img\]((.|\n)*?)\[\/img\]"
				s=re.Replace(s,"<a href=""$2"" target=""_blank"" rel=""lytebox[vacation]"">$2</a>")
				re.Pattern="\[img=(\d{1,3})[x|\,](\d{1,3})\]\s*([^\[\<\r\n]+?)\s*\[\/img\]"
				s=re.Replace(s,"<a href=""$1"" target=""_blank"" rel=""lytebox[vacation]"">$1</a>")
			End If
			re.Pattern="\[UPLOAD=(swf|swi)\]((.|\n)*?)\[\/UPLOAD]"
			If team.Forum_setting(14) = 1 Then
				s=re.Replace(s,"<img src=""images/type/$1.gif"" border=""0"" align=""absmiddle""><br><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=300></embed>")
			Else
				s=re.Replace(s,"<img src=""images/type/$1.gif"" border=""0"" align=""absmiddle""><A HREF=""$2"" TARGET=""_blank"">[ȫ������,ע��Flash���ܺ��в���ȫ����]</A>")
			End if
			'�����ϰ汾���ϴ�
			re.Pattern="\[UPLOAD=(txt|rar|zip)\]([0-9]*)\[\/UPLOAD]"
			If team.Group_Browse(24) = 0 Then 
				s=re.Replace(s,"<img src=""images/type/$1.gif"" border=""0"" align=""absmiddle""> �����ڵ���û�в鿴������Ȩ�ޡ�")
			Else
				s=re.Replace(s,"<img src=""images/type/$1.gif"" border=""0"" align=""absmiddle""><A HREF=""ShowFile.asp?ID=$2"" TARGET=""_blank"">���������ļ�</A>")
			End If
			'�µĸ�����������ͳ��
			re.Pattern="\[UPLOAD=(.[^\[]*)\]ShowFile\.asp\?ID=*([0-9]*)\[\/UPLOAD]"
			If team.Group_Browse(24) = 0 Then
				s=re.Replace(s,"<img src=""images/type/$1.gif"" border=""0"" align=""absmiddle"">�����ڵ���û�в鿴�ø�����Ȩ��")
			Else
				s=re.Replace(s,"<img src=""images/type/$1.gif"" border=""0"" align=""absmiddle""><A HREF=""ShowFile.asp?ID=$2"" TARGET=""_blank"">���������ļ�</A>  (�������ѱ�����&nbsp;<FONT color=red><script language=""javascript"" src=""ShowFile.asp?action=iunm&num=$2""></script></FONT>&nbsp;��)")
			End If
			s=Replace(s,"www.fs2you.com","dyn.www.rayfile.com")'�����ϵ�ַ�Ĵ���
			s=Replace(s,"twffan=""done""","rel=""lytebox[vacation]""")
		End if
		UBB_Code=ChkBadWords(s)
		Set re = Nothing
	End Function

	Function ReadPowers(s)
		Dim re
		set re = New RegExp
		re.Global = True
		re.IgnoreCase = True
		'������+������
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
					s=re.Replace(s,"<fieldset class=textquote><legend><strong>�ظ��ɼ���</strong></legend>���������ѱ�����,���½��鿴!</fieldset>")
				Else
					If tk_UserName = CodeRs(0) or team.ManageUser Or Team.Execute("Select Count(ID) From "&CodeRs(1)&" Where Topicid="&UID&" And UserName='"&TK_UserName&"'")(0)>0  Then
						s=re.Replace(s,"<fieldset class=textquote><legend><strong>�ظ��ɼ���</strong></legend>$1</fieldset>")
					Else
						s=re.Replace(s,"<fieldset class=textquote><legend><strong>�ظ��ɼ���</strong></legend>���������ѱ�����,�ظ�������ſɲ鿴!</fieldset>")
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
					s=re.Replace(s,"<fieldset class=textquote><legend><strong>������</strong></legend>���½�鿴����������!</fieldset>")
				Else
					If team.ManageUser Or Trim(tk_UserName) = Trim(CodeRs(0)) Then
						s=re.Replace(s,"<fieldset class=textquote><legend><strong>������</strong></legend>$2<hr class=""a3"" width=""90%""><li>��Ϊ��ĵȼ������Ƿ�����,�����ۼ�$1Ԫ������Ч,�����<a href=""Command.asp?action=seebuy&buyid="& tID &""" target=""_blank""><B>�鿴�������б�</B>! </a><li> ��Ϊ���ҹ������ѵ�Ե��,Ҳ����<a href=""Command.asp?action=buypost&buyid="& tID &"&postname="& CodeRs(0) &"&money=$1""> <B>�������</B>! </a></fieldset>")
					Else
						Set ts=Team.Execute("Select Name From ["&IsForum&"ListRec] Where  PostID="&tID) 
						If (ts.Eof and ts.Bof) Then
							s=re.Replace(s,"<fieldset class=textquote><legend><strong>������</strong></legend>�����ۼ�$1Ԫ,����Ҫ����ſ������<hr class=""a3"" width=""90%""><li> <a href=""Command.asp?action=buypost&buyid="& tID &"&postname="& CodeRs(0) &"&money=$1""> �������! </a></fieldset>")
						Else
							If Instr(Ts(0),TK_UserName&",") >0 Then
								s=re.Replace(s,"<fieldset class=textquote><legend><strong>������</strong></legend>$2 <hr class=""a3"" width=""90%""><li>�����ۼ�$1Ԫ,���Ѿ������<li><a href=""Command.asp?action=seebuy&buyid="& tID &""" target=""_blank""> �鿴�������б�! </a></fieldset>")
							Else
								s=re.Replace(s,"<fieldset class=textquote><legend><strong>������</strong></legend>�����ۼ�$1Ԫ,����Ҫ����ſ������<hr class=""a3"" width=""90%""><li> <a href=""Command.asp?action=buypost&buyid="& tID &"&postname="& CodeRs(0) &"&money=$1""> �������! </a></fieldset>")
							End If
						End IF
					End IF
				End If
			End If
		End If
		ReadPowers = s
		Set re = Nothing
	End Function

	'ǩ����UBB
	Function Sign_Code(Str,a)
		If Str="" Or IsNull(Str) Then Exit Function
		Dim s,re
		s = Str
		Set re=new RegExp		
		re.IgnoreCase =true
		re.Global=True
		s=Replace(s,"<BR>","<br>")
		s=Replace(s,"</P><P>","</p><p>")
		s=Replace(s,"&lt;","&lt")
		s=Replace(s,"&nbsp;","&nbsp")
		If Int(a) = 0 Then
			Sign_Code = s
			Exit Function
		End if
		re.Pattern="\[marquee\](.*?)\[\/marquee]"
		s=re.Replace(s,"<marquee width=90% behavior=alternate scrollamount=""3"">$1</marquee>")
		re.Pattern="\[font=([^<>\]]*?)\](.*?)\[\/font]"
		s=re.Replace(s,"<font face=""$1"">$2</font>")
		re.Pattern="\[color=([^<>\]]*?)\](.*?)\[\/color]"
		s=re.Replace(s,"<font color=""$1"">$2</font>")
		re.Pattern="\[align=([^<>\]]*?)\](.*?)\[\/align]"
		s=re.Replace(s,"<div align=""$1"">$2</div>")
		re.Pattern="\[size=(\d*?)\](.*?)\[\/size]"
		s=re.Replace(s,"<font size=""$1"">$2</font>")
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
		re.Pattern="\[glow\](.*?)\[\/glow]"
		s=re.Replace(s,"<span style='behavior:url(inc/font.htc)'>$1</span>")
		re.Pattern="\[qq\](\d*?)\[\/qq]"
		s=re.Replace(s,"<a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=$1&Site=team5.cn&Menu=yes><img border=""0"" SRC=http://wpa.qq.com/pa?p=1:$1:5 alt=""���������ҷ���Ϣ""></a>")
		re.Pattern="\[URL\](.*?)\[\/URL]"
		s=re.Replace(s,"<A HREF=""$1"" TARGET=_blank>$1</A>")
		re.Pattern="(\[URL=(.[^\[]*)\])(.*?)(\[\/URL\])"
		s= re.Replace(s,"<A HREF=""$2"" TARGET=_blank>$3</A>")
		re.Pattern="\[IMG\](.*?)\[\/IMG]"
		s=re.Replace(s,"<IMG SRC=""$1"" border=0>")
		re.Pattern="\[QUOTE\](.*?)\[\/QUOTE]"
		s=re.Replace(s,"<div class=""quote"">$1</div>")
		Sign_Code=ChkBadWords(s)
		Set re = Nothing
	End Function

	Private Function TM_Alipay_PayTo(strText)
		If Not Isnull(strText) Then
			Dim s,ss,re
			Dim match,match2,urlStr,re2
			Dim t(2),temp,check,fee,i,encode8_tmp
			s=strText
			Set re=new RegExp
			re.IgnoreCase =true
			re.Global=true
			Set re2=new RegExp
			re2.IgnoreCase =true
			re2.Global=False
			t(0)="���ҳе��˷�"
			t(1)="��ҳе��˷�"
			t(2)="������Ʒ�����ʵ�"
			s=strText
			re.Pattern="\[\/payto\]"
			s=re.replace(s, chr(1)&"/payto]")
			re.Pattern="\[payto\]([^\x01]+)\x01\/payto\]"
			Set match = re.Execute(s)
			re.Global=False
			For i=0 To match.count-1
				re2.Pattern="\(seller\)([^\n]+?)\(\/seller\)"
				If re2.Test(match.item(i)) Then
					Set match2 = re2.Execute(match.item(i))
					temp=re2.replace(match2.item(0),"$1")
					temp= replace(temp,"#","@")
					ss=""
					urlStr="API/payto.asp?seller="&temp
					re2.Pattern="\(subject\)([^\n]+?)\(\/subject\)"
					If re2.Test(match.item(i)) Then
						Set match2 = re2.Execute(match.item(i))
						temp=re2.replace(match2.item(0),"$1")
						ss=ss&"<div class=code><br/><b>��Ʒ����</b>��"&temp&"<br/><br/>"
						urlStr = urlStr & "&subject=" & Server.UrlEncode(temp)
						re2.Pattern="\(body\)((.|\n)*?)\(\/body\)"
						If re2.Test(match.item(i)) Then
							Set match2 = re2.Execute(match.item(i))
							temp=re2.replace(match2.item(0),"$1")
							ss=ss&"<b>��Ʒ˵��</b>��"&temp&"<br/><br/>"
							urlStr = urlStr & "&body=" & Server.UrlEncode(Cutstr(temp,200))
							re2.Pattern="\(price\)([\d\.]+?)\(\/price\)"
							If re2.Test(match.item(i)) Then
								Set match2 = re2.Execute(match.item(i))
								temp=re2.replace(match2.item(0),"$1")
								ss=ss&"<b>��Ʒ�۸�</b>��"&temp&" Ԫ<br/><br/>"
								urlStr=urlStr&"&price="&temp
								re2.Pattern="\(transport\)([1-3])\(\/transport\)"
								If re2.Test(match.item(i)) Then
									Set match2 = re2.Execute(match.item(i))
									temp=re2.replace(match2.item(0),"$1")
									check=true
									If int(temp)=2 Then
										re2.Pattern="\(express_fee\)([\d\.]+?)\(\/express_fee\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											fee=re2.replace(match2.item(0),"$1")
											ss=ss&"<b>�ʵ���Ϣ</b>��"&t(temp-1)&"����� "&fee&" Ԫ<br/><br/>"
											urlStr=urlStr&"&transport="&temp&"&express_fee="&fee
										Else
											re2.Pattern="\(ordinary_fee\)([\d\.]+?)\(\/ordinary_fee\)"
											If re2.Test(match.item(i)) Then
												Set match2 = re2.Execute(match.item(i))
												fee=re2.replace(match2.item(0),"$1")
												ss=ss&"<b>�ʵ���Ϣ</b>��"&t(temp-1)&"��ƽ�� "&fee&" Ԫ<br/><br/>"
												urlStr=urlStr&"&transport="&temp&"&ordinary_fee="&fee
											Else
												check=False
											End If
										End If
									Else
										ss=ss&"<b>�ʵ���Ϣ</b>��"&t(temp-1)&"<br/><br/>"
										urlStr=urlStr&"&transport="&temp
									End If
									If check=true Then
										check=False
										re2.Pattern="\(ww\)([^\n]+?)\(\/ww\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											temp=re2.replace(match2.item(0),"$1")
											encode8_tmp=EncodeUtf8(temp)
											ss=ss&"<b>��ϵ����</b>��<a target=""_blank"" href=""http://amos1.taobao.com/msg.ww?v=2&amp;uid="&encode8_tmp&"&amp;s=1""><img border=""0"" src=""http://amos1.taobao.com/online.ww?v=2&amp;uid="&encode8_tmp&"&amp;s=1""/></a>"
											check=true
										End If
										re2.Pattern="\(qq\)(\d+?)\(\/qq\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											temp=re2.replace(match2.item(0),"$1")
											If check=true Then
												ss=ss&"&nbsp;&nbsp;<a target=""_blank"" href=""http://wpa.qq.com/msgrd?V=1&Uin="&temp&"&Site=team5.cn&Menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=1:"&temp&":10"" alt=""��ϵ��"" /></a><br/><br/>"
											Else
												ss=ss&"<b>��ϵ����</b>��<a target=""_blank"" href=""http://wpa.qq.com/msgrd?V=1&Uin="&temp&"&Site=team5.cn&Menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=1:"&temp&":10"" alt=""��ϵ��"" /></a><br/><br/>"
											End If
										ElseIf check=true Then
											ss=ss&"<br/><br/>"
										End If
										re2.Pattern="\(demo\)([^\n]+?)\(\/demo\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											temp=re2.replace(match2.item(0),"$1")
											ss=ss&"<b>��ʾ��ַ</b>��"&temp&"<br/><br/>"
											'urlStr=urlStr&"&url="&temp
										End If
										ss=ss&"<a href="""&Server.HtmlEncode(urlStr&"&partner=2088002048522272&type=1&readonly=true")&""" target=""_blank""><img src=""images/alipay.gif"" border=""0"" alt=""ͨ��֧�������ף����������ģ��������ѡ���ȫ����ݣ�"" /></a>&nbsp;&nbsp;<a href=""https://www.alipay.com/static/help/help.htm"" target=""_blank""><font color=""blue"">�鿴���װ�������������</font></a><br/></div>"
										s=re.replace(s,ss)
									End If
								End If
							End If
						End If
					End If
				End If
			Next
			Set match=Nothing
			Set re2=Nothing
			Set match2=Nothing
			re.Global=true
			re.Pattern="\x01\/payto\]"
			s=re.replace(s,"[/payto]")
			TM_Alipay_PayTo=s
		End If
	End Function
%>
<script type="text/javascript" runat="server" language=javascript>
 function EncodeUtf8(s1)
  {
      var s = escape(s1);
      var sa = s.split("%");
      var retV ="";
      if(sa[0] != "")
      {
         retV = sa[0];
      }
      for(var i = 1; i < sa.length; i ++)
      {
           if(sa[i].substring(0,1) == "u")
           {
               retV += Hex2Utf8(Str2Hex(sa[i].substring(1,5))) + sa[i].substring(5,sa[i].length);
               
           }
           else retV += "%" + sa[i];
      }
      
      return retV;
  }
  function Str2Hex(s)
  {
      var c = "";
      var n;
      var ss = "0123456789ABCDEF";
      var digS = "";
      for(var i = 0; i < s.length; i ++)
      {
         c = s.charAt(i);
         n = ss.indexOf(c);
         digS += Dec2Dig(eval(n));
           
      }
      //return value;
      return digS;
  }
  function Dec2Dig(n1)
  {
      var s = "";
      var n2 = 0;
      for(var i = 0; i < 4; i++)
      {
         n2 = Math.pow(2,3 - i);
         if(n1 >= n2)
         {
            s += '1';
            n1 = n1 - n2;
          }
         else
          s += '0';
          
      }
      return s;
      
  }
  function Dig2Dec(s)
  {
      var retV = 0;
      if(s.length == 4)
      {
          for(var i = 0; i < 4; i ++)
          {
              retV += eval(s.charAt(i)) * Math.pow(2, 3 - i);
          }
          return retV;
      }
      return -1;
  } 
  function Hex2Utf8(s)
  {
     var retS = "";
     var tempS = "";
     var ss = "";
     if(s.length == 16)
     {
         tempS = "1110" + s.substring(0, 4);
         tempS += "10" +  s.substring(4, 10); 
         tempS += "10" + s.substring(10,16); 
         var sss = "0123456789ABCDEF";
         for(var i = 0; i < 3; i ++)
         {
            retS += "%";
            ss = tempS.substring(i * 8, (eval(i)+1)*8);
            
            
            
            retS += sss.charAt(Dig2Dec(ss.substring(0,4)));
            retS += sss.charAt(Dig2Dec(ss.substring(4,8)));
         }
         return retS;
     }
     return "";
  } 
</script>