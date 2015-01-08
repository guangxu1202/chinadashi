<%
	'��̳���ú���
	Function SetColors(a)	
		Dim tmp
		Select Case a
			case "1"
				tmp = "font-weight:bold;color:#808080;"
			case "2"
				tmp = "font-weight:bold;color:#808000;"
			case "3"
				tmp = "font-weight:bold;color:#008000;"
			case "4"
				tmp = "font-weight:bold;color:#0000ff;"
			case "5"
				tmp = "font-weight:bold;color:#800000;"
			case "6"
				tmp = "font-weight:bold;color:#ff0000;"
			case "7"
				tmp = "font-weight:bold;color:#cc0066;"
			Case Else
				tmp=""
		End Select
		SetColors = tmp
	End Function

    Public Function CodeCookie(Str)
		If set_cookies = 1 Then
			Dim i
			Dim StrRtn
			For i = Len(Str) To 1 Step -1
				StrRtn = StrRtn & AscW(Mid(Str, i, 1))
				If (i <> 1) Then StrRtn = StrRtn & "a"
			Next
			CodeCookie = StrRtn
		Else
			CodeCookie = Str
		End If
    End Function
    Public Function DecodeCookie(Str)
		If set_cookies = 1 Then
			Dim i
		    Dim StrArr, StrRtn
			StrArr = Split(Str, "a")
			For i = 0 To UBound(StrArr)
				If IsNumeric(StrArr(i)) = True Then
				    StrRtn = ChrW(StrArr(i)) & StrRtn
			    Else
				    StrRtn = Str
				    Exit Function
				End If
			Next
			DecodeCookie = StrRtn
		Else
			DecodeCookie = Str
		End If
    End Function

	'����ַ����ֵĴ��� GetRepeatTimes(�������ַ�,��Ҫ�����ı�)
	Function GetRepeatTimes(TheChar,TheString)
		GetRepeatTimes = (Len(TheString)-Len(Replace(TheString,TheChar,"")))/Len(TheChar)
	End Function

	'����ַ���������
    Function Echo(Str)
		Response.Write Str & VbCrlf
	End Function

	'��ҳ�ж�
	Function CheckNum(ByVal strStr,ByVal blnMin,ByVal blnMax,ByVal intMin,ByVal intMax)
		Dim i,s,iMi,iMa
		s=Left(Trim(""&strStr),32):iMi=intMin:iMa=intMax
		If IsNumeric(s) Then
			i=CDbl(s)
			i=IIf(blnMin=1 And i<iMi,iMi,i)
			i=IIf(blnMax=1 And i>iMa,iMa,i)
		Else
			i=iMi
		End If
		CheckNum=i
	End Function
	'**************************************************
	'�� �� ����CID
	'��    �ã�ת��Ϊ��Ч�� ID
	'����ֵ���ͣ�Integer (>=0)
	'**************************************************
	Function CID(strS)
		Dim intI
		intI = 0
		If IsNull(strS) Or strS = "" Then
			intI = 0
		Else
			If Not IsNumeric(strS) Then
				intI = 0
			Else
				Dim intk
				On Error Resume Next
				intk = Abs(Clng(strS))
				If Err.Number = 6 Then intk = 0  ''�������
				Err.Clear
				intI = intk
				'intI = Int(intk)
			End If
		End If
		CID = intI
	End Function

	'**************************************************
	'�� �� ����HRF
	'��    �ã�ת��Ϊ��Ч�� Request����������
	'����ֵ���ͣ����Ǻ���ַ�
	'**************************************************
	Function HRF(a,b,c)
		Dim Str
		Select Case a
			Case 1
				Str = Request.Form(c)
			Case 2
				Str = Request.QueryString(c)
			Case 3
				Str = Request.Cookies(c)
			Case Else
				Str = Request(c)
		End Select
		Select Case b
			Case 1
				Str = HtmlEncode(str)
			Case 2
				Str = CID(str)
		End Select
		HRF = Str
	End Function
	'�ж��û���
	Function IstrueName(uName)
		Dim Hname,i
		IstrueName = False
		Hname = Array("=","%",chr(32),"?","&",";",",","'",",",chr(34),chr(9),"��","$","|")
		For i = 0 To Ubound(Hname)
			If InStr(uName,Hname(i)) > 0 Then
				Exit Function
			End If
		Next
		IstrueName=True 
	End Function

	Function Fixjs(Strings)
		Dim s,R,Re
		s = Strings
		Set re=New RegExp
		re.IgnoreCase =True
		re.Global=True
		If Not IsNull(s) Then
			R="(expression|xss:|alert|function|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
			re.Pattern= "<((.[^>]*" & r &"[^>]*))>"
			s=re.Replace(s,"")
			re.Pattern=R
			s=re.Replace(s,"")
			s=Replace(s,"..","")
			s=Replace(s,"\","/")
			s=Replace(s,"^","")
			s=Replace(s,"#","")
			s=Replace(s,"%","")
			s=Replace(s,"|","")
		End if
		Fixjs = s
		Set re = Nothing
	End Function

	'**************************************************
	'��������HTMLEncode
	'��  �ã������ַ�
	'��  ����str-----Ҫ���ǵ��ַ�
	'����ֵ�����Ǻ���ַ�
	'**************************************************
	Public Function HTMLEncode(fString)
		If fString="" or IsNull(fString) Then 
			Exit Function
		Else
			If Instr(fString,"'")>0 Then 
				fString = replace(fString, "'","&#39;")
			End If
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), "&nbsp;")
			fString = Replace(fString, CHR(9), "&nbsp;")
			fString = Replace(fString, CHR(34), "&quot;")
			fString = Replace(fString, CHR(13),"")
			fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
			fString = Replace(fString, CHR(10), "<BR>")
			fString = Replace(fString, CHR(39), "&#39;")
			fString = Replace(fString, CHR(0), "")
			fString = ChkBadWords(fString)
			HTMLEncode = fString
		End If
	End Function
	'��ԭ�ַ�����
	Public Function iHTMLEncode(fString)
		If fString="" or IsNull(fString) Then 
			Exit Function
		Else
			If Instr(fString,"'")>0 Then 
				fString = replace(fString, "'","&#39;")
			End If
			fString = replace(fString, "&gt;"	, ">")
			fString = replace(fString, "&lt;"	, "<")
			fString = Replace(fString, "&nbsp;"	, CHR(32))
			fString = Replace(fString, "&nbsp;"	, CHR(9))
			fString = Replace(fString, "&quot;"	, CHR(34))
			fString = Replace(fString, ""		, CHR(13))
			fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
			fString = Replace(fString, "<BR>"	, CHR(10))
			fString = Replace(fString, ""		, CHR(0))
			fString = Replace(fString, "&#39;"	, CHR(39))
			fString = ChkBadWords(fString)
			iHTMLEncode = fString
		End If
	End Function

	Function TempCode(strContent,a)
		If a="" or IsNull(a) Then 
			Exit Function
		Else
			Dim re
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="\["&a&"\](.*?)\[\/"&a&"]"
			strContent=re.Replace(strContent,"")
			set re=Nothing
		End if
		TempCode=strContent
	End Function
	Function BlackTmp(strContent,a)
		If a="" or IsNull(a) Then 
			Exit Function
		Else
			Dim re
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="\["&a&"\](.*?)\[\/"&a&"]"
			strContent=re.Replace(strContent,"$1")
			set re=Nothing
		End if
		BlackTmp = strContent
	End Function

	'�����û�����
	Public Function UserBad(m,s)
		If s="" Or IsNull(s) Then 
			UserBad = m
			Exit Function
		End if
		If m="" Or IsNull(m) Then 
			UserBad = m
			Exit Function
		End if
		Dim i,BadWords
		If team.Club_Class(7)&"" = "" Then
			UserBad = m
			Exit Function
		Else
			If Instr(team.Club_Class(7),Chr(13)&Chr(10))>0 Then 	
				BadWords = Split(team.Club_Class(7),Chr(13)&Chr(10))
				For i = 0 To UBound(BadWords)
					If Trim(s) = Trim(BadWords(i)) Then
						m = "<div style='height:30px;line-height:30px;width:150px;border: 1px solid #6595d6;border-top:3px double #6595d6;text-align : center;color:#00008B;margin :150px 0px 0px 10px;background-color : #e4e8ef;float:right;}.vote{float:left;border-left: 1px solid #6595d6;height:28px;'>���û������Ѿ���ϵͳ����</div>"
					End If
				Next
			Else
				If Trim(s) = Trim(team.Club_Class(7)) Then
					m = "<div style='height:30px;line-height:30px;width:150px;border: 1px solid #6595d6;border-top:3px double #6595d6;text-align : center;color:#00008B;margin :150px 0px 0px 10px;background-color : #e4e8ef;float:right;}.vote{float:left;border-left: 1px solid #6595d6;height:28px;'>���û������Ѿ���ϵͳ����</div>"
				End If
			End If
		End If 
		UserBad = m
	End Function

	'**************************************************
	'��������ChkBadWords
	'��  �ã������ַ�
	'��  ����str-----Ҫ���ε��ַ�
	'����ֵ���滻���κ���ַ�
	'**************************************************
	Public Function ChkBadWords(Str)
		If IsNull(Str) Or str="" Then 
			ChkBadWords = str
			Exit Function
		End if
		If team.Club_Class(5)&""="" Then 
			ChkBadWords = str
			Exit Function
		End if
		Dim i,BadWords,Badtext,aa,bb,ii
		If Instr(team.Club_Class(5),Chr(13)&Chr(10))>0 Then
			BadWords=Split(team.Club_Class(5),Chr(13)&Chr(10))
			For i = 0 To UBound(BadWords)
				Badtext=Split(BadWords(i),"=")
				For ii = 0 To UBound(Badtext)
					If InStr(Str,Badtext(0))>0 Then
						Str = Replace(Str,Badtext(0),Badtext(1))
					End If
				Next
			Next
		Else
			Badtext=Split(team.Club_Class(5),"=")
			For ii = 0 To UBound(Badtext)
				If InStr(Str,Badtext(0))>0 Then
					Str = Replace(Str,Badtext(0),Badtext(1))
				End If
			Next
		End If
		ChkBadWords = Str
	End Function
	'**************************************************
	'��������ReplaceStr
	'��  �ã��滻�ַ� (��ֹ��ֵ)
	'��  ����str   ----��Ҫ�滻���ַ���
	'����ֵ���滻���ַ�
	'**************************************************
	Function ReplaceStr(str,str1,str2)
		If str<>"" Then
			ReplaceStr = Replace(str,str1,str2&"")
		Else
			ReplaceStr=str
		End If
	End Function
	'**************************************************
	'��������strLength
	'��  �ã������ַ���������һ���������ַ���Ӣ����һ���ַ�
	'��  ����str   ----��Ҫ������ַ���
	'����ֵ������ֵ
	'**************************************************
	Function strLength(str)
		ON ERROR RESUME NEXT
		dim WINNT_CHINESE
		WINNT_CHINESE    = (len("��̳")=2)
		if WINNT_CHINESE then
			dim l,t,c
			dim i
			l=len(str)
			t=l
			for i=1 to l
				c=asc(mid(str,i,1))
				if c<0 then c=c+65536
				if c>255 then
					t=t+1
				end if
		    next
			strLength=t
		Else 
			strLength=len(str)
		End if
		if err.number<>0 then err.clear
	End Function
	'**************************************************
	'��������cutStr
	'��  �ã����ַ���������β����".."��.����һ���������ַ���Ӣ����һ���ַ�
	'��  ����str   ----ԭ�ַ���	'strlen ----��ȡ����
	'����ֵ����ȡ����ַ���+.....
	'**************************************************
	Function cutStr(str,strlen)
		dim l,t,c,i
		l=len(str)
		t=0
		for i=1 to l
			c=Abs(Asc(Mid(str,i,1)))
			if c>255 then
				t=t+2
			else
				t=t+1
			end if
			if t>=strlen then
				cutStr=left(str,i)&"..."
				exit for
			else
				cutStr=str
			end if
		next
		cutStr=replace(cutStr,chr(10),"")
	End Function
	'**************************************************
	'��������Emailto
	'��  �ã������ʼ�
	'��  ����--
	'����ֵ��--
	'**************************************************
	Sub Emailto(a,b,c)
		if Not IsValidEmail(a) Then team.error2 " E-mail��ַ��д����"
		Select Case team.Forum_setting(1)
			Case "1"
				Dim JMail
				Set JMail=Server.CreateObject("JMail.Message")
				If -2147221005 = Err Then team.error2 "����������֧�� JMail.Message �����"
				JMail.Charset="gb2312"
				JMail.AddRecipient a
				JMail.Subject = b
				JMail.Body = c
				JMail.From = team.Forum_setting(57)						'�����˵�ַ
				JMail.MailServerUserName = team.Forum_setting(41)		'��������½�û���
				JMail.MailServerPassword = team.Forum_setting(55)		'��������½����
				JMail.Send team.Forum_setting(58)						'��������ַ
				Set JMail=nothing
			Case "2"
				Dim MailObject
				Set MailObject = Server.CreateObject("CDONTS.NewMail")
				If -2147221005 = Err Then team.error2 "����������֧�� CDONTS.NewMail �����"
				MailObject.Send team.Forum_setting(57),a,b,c
				'MailObject.Send "���ͷ��ʼ���ַ","���շ��ʼ���ַ","����","�ʼ�����"
				Set MailObject=nothing
			Case Else
				Exit Sub
		End Select
	End Sub
	'����HTML����
	Public Function Replacehtml(Textstr)
		Dim Str,re
		Str=Textstr
		Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str, "")
			Set Re=Nothing
			Replacehtml=Str
	End Function
	'**************************************************
	'��������GetColor
	'��  �ã��ؼ��ּ�ɫ��ʾ 
	'��  ����str----Ҫ���˵��ַ�   str ----�ؼ���
	'����ֵ�����������Ĺؼ����Զ���ɫ
	'**************************************************
	Function GetColor(str,str1)
		If IsNull(Str) Then Exit Function
		Dim i,Text
		If str="" Or IsNull(str) Then Exit function
		'If InStr(Str,str1)>0 Then		'ע�͵�����ʹ���ݲ��ֱ��Сд
			Str = ReplaceStr(Str,str1,"<Span Style='font-weight: bold;color: #FF0000;'>"&str1&"</Span>")
		'End If
		GetColor = Str
	End Function
	Rem ����½ By DayMoon 05.09.17
	Sub TestUser()
		If Not Team.UserLoginED then 
			Team.Error("<li>����δ<a href=login.asp>��¼</a>��̳")
		End If
	End Sub

	Rem ���ַ��������� By DayMoon 05.10.21
	Function ReadCode(str,Str1)
		dim name
		dim result
		dim i,j,k
		If isnull(str) then
			ReadCode=""
			Exit Function
		End If
		Randomize 
		k=instr(str,"</P>")
		Do while k>0
			result=""
			for i=0 to 19
				j=Int(128 * Rnd)+1
				if j=60 or j=62 then
					j=j+1
				end if
				result =result&chr(j) ' �����������
			next 
			result="<span style='DISPLAY: none'>"&result&"</span>"
			str=replace(str,"</P>",result&"<'/P>",1,1)
			k=instr(str,"</P>")
		loop
		str=replace(str,"<'/P>","</P>")
		k=instr(str,"<BR>")
		Do while k>0
			result=""
			for i=0 to 19
				j=Int(128 * Rnd)+1
				if j=60 or j=62 then
					j=j+1
				end if
				result =result&chr(j) ' �����������
			next 
			result="<span style='DISPLAY: none'>"&result&"</span>"
			str=replace(str,"<BR>",result&"<'BR>",1,1)
			k=instr(str,"<BR>")
		loop
		str=replace(str,"<'BR>","<BR>")
		ReadCode=str&"<div align=right style='color=gray'>[������Ȩ��ԭ���߼�"&Str1&"��ͬӵ�У�ת��������������Ȩ]</div>"
	End Function

	Function UserOnlinetimes(RL_ActTimeT)
		Dim RL_UserClass,RL_NextClassNeed,RL_Str,TempStr,i
		RL_UserClass = 0
		RL_NextClassNeed = 0
		If RL_ActTimeT = "" Or IsNull(RL_ActTimeT) Then Exit Function
		For i=1 to 60
			if RL_ActTimeT \ 60 < 6*i*i + 6*i then 
				RL_NextClassNeed = (6*i*i + 6*i)*60 - RL_ActTimeT
				Exit For
			end if
			RL_UserClass = RL_UserClass + 1
		Next
		RL_Str = ""
		TempStr = "����:"&RL_ActTimeT \ 60&"Сʱ"&RL_ActTimeT mod 60&"����.&#13&#10��������"&RL_NextClassNeed\60&"Сʱ	"&RL_NextClassNeed mod 60&"����.&#13&#10Ŀǰ�ȼ�:"&RL_UserClass&""
		if RL_UserClass = 0 then
			RL_Str = RL_Str & "<img src="""&team.styleurl&"/star1.gif"" alt="""&TempStr&""">"
		end if
		For i=1 to RL_UserClass \ 16
			RL_Str = RL_Str & "<img src="""&team.styleurl&"/star3.gif"" alt="""&TempStr&""">"
		Next
		RL_UserClass = RL_UserClass mod 16
		For i=1 to RL_UserClass \ 4
			RL_Str = RL_Str & "<img src="""&team.styleurl&"/star2.gif"" alt="""&TempStr&""">"
		Next
		RL_UserClass = RL_UserClass mod 4
		For i=1 to RL_UserClass
			RL_Str = RL_Str & "<img src="""&team.styleurl&"/star1.gif"" alt="""&TempStr&""">"
		Next
		UserOnlinetimes = RL_Str
	End Function
	Function UserStar(Level)
		Dim Star,Moon,Sun
		Dim StarCount,MoonCount,SunCount
		Dim i,ImgStr
		Star=1
		Moon= Cid(team.Forum_setting(23))
		Sun= Cid(team.Forum_setting(23) * team.Forum_setting(23))
		SunCount=Level\Sun
		MoonCount=(Level mod Sun)\Moon
		StarCount=Level mod Moon
		for i=1 to SunCount
			ImgStr=ImgStr & "<img src="""&team.styleurl&"/star3.gif"" border=""0"" align=""absmiddle"" alt=""Rank:"&Level&""">"
		Next
		for i=1 to MoonCount
			ImgStr=ImgStr & "<img src="""&team.styleurl&"/star2.gif"" border=""0"" align=""absmiddle"" alt=""Rank:"&Level&""">"
		Next
		for i=0 to StarCount
			ImgStr=ImgStr & "<img src="""&team.styleurl&"/star1.gif"" border=""0"" align=""absmiddle"" alt=""Rank:"&Level&""">"
		Next
		UserStar=ImgStr
	End Function

	Function GetUrlXmls(G)
		Dim Http,XmlDom,tmp,UpTimes,DownTimes,i,NodeGather,tagCount
		Set http = Server.CreateObject("Microsoft.XMLHTTP")
		http.Open "GET", G ,False
		http.send
		Set XmlDom = Server.CreateObject("Microsoft.XmlDom")
		XmlDom.Async = true
		XmlDom.ValidateOnParse = False
		XmlDom.Load( http.ResponseXML)
		set NodeGather = XmlDom.getElementsByTagName("item")
		tagCount = nodeGather.length
		For I = 0 To tagCount-1
			UpTimes = nodeGather(I).getAttribute("UpTimes")
			DownTimes = nodeGather(I).getAttribute("DownTimes")
			If IsDate(UpTimes) And IsDate(DownTimes) Then
				If DateDiff("d",Date(),UpTimes)<=0 And DateDiff("d",Date(),DownTimes)>0 then
					tmp = tmp & nodeGather(I).ChildNodes(0).text
				End If
			End if
		Next
		set NodeGather = Nothing
		GetUrlXmls = tmp
	End Function
	
	Function GetUseSex(a)
		Dim tmp
		If a = 2 Then
			tmp = "<img src="""&team.styleurl&"/female.gif"" border=""0"" align=""absmiddle"" alt=""��Ů����Ů��������Ů��Ů""> "
		Elseif a = 1 Then
			tmp = "<img src="""&team.styleurl&"/Male.gif"" border=""0"" align=""absmiddle"" alt=""ż��˧��!""> "
		Else
			tmp = ""
		End if
		GetUseSex = tmp
	End Function
	
	Function Astro(str)
		Dim a,b,c,d
		str = Trim(str)
		If str&""="" Or IsNull(Str) Then
			Astro=""
			Exit Function
		End If
		If Not pIsDate(str) Then
			Astro=""
			Exit Function			
		End If
		a=Split(str,"-")
		b=a(1)
		c=a(2)
		Select Case b
			Case "1"
				If c>=21 Then
					d="<img src=""images/star/h.gif"" alt=""ˮƿ��"">"
				Else
					d="<img src=""images/star/g.gif"" alt=""ħ����"">"
				End If
			Case "2"
				If c>=20 Then
					d="<img src=""images/star/i.gif"" alt=""˫����"">"
				Else
					d="<img src=""images/star/h.gif"" alt=""ˮƿ��"">"
				End If
			Case "3"
				If c>=21 Then
					d="<img src=""images/star/1.gif"" alt=""������"">"
				Else
					d="<img src=""images/star/i.gif"" alt=""˫����"">"
				End If
			Case "4"
				If c>=21 Then
					d="<img src=""images/star/2.gif"" alt=""��ţ��"">"
				Else
					d="<img src=""images/star/1.gif"" alt=""������"">"
				End If
			Case "5"
				If c>=22 Then
					d="<img src=""images/star/3.gif"" alt=""˫����"">"
				Else
					d="<img src=""images/star/2.gif"" alt=""��ţ��"">"
				End If
			Case "6"
				If c>=22 Then
					d="<img src=""images/star/4.gif"" alt=""��з��"">"
				Else
					d="<img src=""images/star/3.gif"" alt=""˫����"">"
				End If
			Case "7"
				If c>=23 Then
					d="<img src=""images/star/b.gif"" alt=""ʨ����"">"
				Else
					d="<img src=""images/star/4.gif"" alt=""��з��"">"
				End If
			Case "8"
				If c>=24 Then
					d="<img src=""images/star/c.gif"" alt=""��Ů��"">"
				Else
					d="<img src=""images/star/b.gif"" alt=""ʨ����"">"
				End If
			Case "9"
				If c>=24 Then
					d="<img src=""images/star/d.gif"" alt=""�����"">"
				Else
					d="<img src=""images/star/c.gif"" alt=""��Ů��"">"
				End If
			Case "10"
				If c>=24 Then
					d="<img src=""images/star/e.gif"" alt=""��Ы��"">"
				Else
					d="<img src=""images/star/d.gif"" alt=""�����"">"
				End If
			Case "11"
				If c>=23 Then
					d="<img src=""images/star/f.gif"" alt=""������"">"
				Else
					d="<img src=""images/star/e.gif"" alt=""��Ы��"">"
				End If
			Case "12"
				If c>=22 Then
					d="<img src=""images/star/g.gif"" alt=""ħ����"">"
				Else
					d="<img src=""images/star/f.gif"" alt=""������"">"
				End If
			Case Else
				d=""
		End Select
		Astro = d
	End Function

	Function pIsDate(s)
		Dim a
		pIsDate = False
		If s = "" Or IsNull(s) Then
			Exit Function
		End If
		If Not IsDate(s) Then
			Exit Function
		Else
			a = Split(s,"-")
			If UBound(a)<2 Then
				Exit Function
			End if
			If Len(a(0))<>4 Then
				Exit Function
			End If
			If Len(a(1))>2 Or Len(a(1))<1 Then
				Exit Function
			End If
			If Len(a(2))>2 Or Len(a(2))<1 Then
				Exit Function
			End If
		End If
		pIsDate = true
	End function

	Function GetPet(str)
		Dim a,B,C,D
		a=Split(str,"-")
		d = Mid("��ţ���������R����u���i",((a(0)-4) Mod 12)+1,1)
		GetPet = D
	End Function

	Function IIf(ByVal blnBool,ByVal strStr1,ByVal strStr2)
		Dim s
		If blnBool Then
			s=strStr1
		Else
			s=strStr2
		End If
		IIf=s
	End Function
	Function IsValidEmail(email)
		Dim names, name, i, c
		IsValidEmail = True
		names = Split(email, "@")
		If UBound(names) <> 1 Then
			IsValidEmail = False
			Exit Function
		End If
		For Each name In names
			If Len(name) <= 0 Then
				IsValidEmail = False
				Exit Function
			End If
			For i = 1 To Len(name)
				c = Lcase(Mid(name, i, 1))
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) Then
					IsValidEmail = False
					Exit Function
				End If
			Next
			If Left(name, 1) = "." or Right(name, 1) = "." Then
				IsValidEmail = False
				Exit Function
			End If
		Next
		If InStr(names(1), ".") <= 0 Then
			IsValidEmail = False
			Exit Function
		End If
		i = Len(names(1)) - InStrRev(names(1), ".")
		If i <> 2 and i <> 3 Then
			IsValidEmail = False
			Exit Function
		End If
		If InStr(email, "..") > 0 Then
			IsValidEmail = False
		End If
	End Function
%>