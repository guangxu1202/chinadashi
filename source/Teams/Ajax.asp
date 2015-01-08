<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Response.Charset = "gb2312"
team.ChkPost

If Request("checksubmit")="yes" Then
	Select Case Request("action")
		Case "checkusername"
			Call checkusername
		Case "checkseccode"
			Call checkseccode
		Case "checkemail"
			Call checkemail
		Case "smilies"
			Call smilies
		Case "loadtopics"
			loadtopics
	End Select

	Sub loadtopics
		Dim Gs,tid,uName,tmp
		tid = CID(Request("tid"))
		uName = HtmlEncode(Request("uname"))
		Set Gs = team.execute("Select Top 5 ID,Topic,UserName,Views,Replies,Lasttime From ["&Isforum&"Forum] Where deltopic=0 and UserName='"& uName &"' and Not (ID="& tID &") order By Posttime Desc")
		If Gs.Eof And Gs.Bof Then
			Exit sub
		Else
			tmp = "<div class=""textdiv""><ul>"
			Do While not Gs.Eof
				tmp = tmp & IIF(CID(team.Forum_setting(65))=1,"<li><a href=""thread-"&Gs(0)&".html"" target=""_blank"">"& Gs(1) &"</a></li> ","<li><a href=""thread.asp?tid="&Gs(0)&""" target=""_blank"">"& Gs(1) &"</a></li> ")
				Gs.MoveNext
			Loop
			tmp = tmp & " </ul></div> "
		End If
		Echo tmp
	End Sub 

	Sub smilies
		Dim i,page,AllSmNum,u,NextM,NextP,iNum,PNum
		page = CID(Request("page"))
		If page < 1 Then page = 1
		PNum		= 87								'总表情数
		iNum		= CID(Request("inum"))				'分页数
		AllSmNum	= Abs(int(-Abs(PNum/iNum)))			'总分页数
		If Page = 1 Then
			NextP = 1
			NextM = iNum
		Else
			NextM = iNum * Page
			NextP = (NextM - (NextM/Page))+1
			If NextM > PNum Then NextM = PNum
		End If
		Echo "<div class=""smdiv"">"
		For i = NextP To NextM
			Echo "<img src=""images/Emotions/"& i &".gif"" alt=""[em"& i &"]"" id=""smilie_"& i &""" border=""0"" onClick=""insertSmiley("& i &")"" width=""33"" height=""33"" onMouseover=""this.style.cursor = 'hand'""/>"
		Next
		Echo "</div>"
		Echo "<div class=""smpages"">"
		For i =1 To AllSmNum
			Echo "<a onclick=""loadsmile("& iNum &","& i &")"" style=""cursor:hand"">"& i &"</a>"
		Next 
		Echo "</div>"
	End Sub 

	Sub checkusername
		dim username,Rs,SQL,tmp
		UserName=HRF(2,1,"username")
		Set Rs =team.execute( "Select * from [user] where username='"&UserName&"'")
		If Rs.Eof  and Rs.Bof Then
			tmp = "num1"
		Else
			tmp = "num2"
		end If
		rs.close:set rs=Nothing
		tmp = TestUName(UserName,tmp)
		Echo tmp
	End Sub

	function TestUName(s,b)
		Dim tmp,i,tmp1,u,rmps
		rmps = b
		If IsNull(team.Club_Class(25)) Or team.Club_Class(25) = "" Then
			rmps = b
		Else
			If Instr(team.Club_Class(25),Chr(13)&Chr(10))>0 Then 
				tmp = Split(team.Club_Class(25),Chr(13)&Chr(10))
				For i = 0 To UBound(tmp)
					If InStr(tmp(i),"*") > 0 Then
						tmp1 = Split(tmp(i),"*")
						For u=0 To UBound(tmp1)
							If tmp1(u) <> "" Then 
								If InStr(s,tmp1(u)) > 0 Then rmps = "num3"
							End if
						Next
					Else
						If InStr(s,tmp(i)) > 0 Then rmps = "num3"
					End If
				Next 
			Else
				tmp = team.Club_Class(25)
				If InStr(tmp,"*") > 0 Then
					tmp1 = Split(tmp,"*")
					For u=0 To UBound(tmp1)
						If tmp1(u) <> "" Then 
							If InStr(s,tmp1(u)) > 0 Then rmps= "num3"
						End if
					Next
				Else
					If InStr(s,tmp) > 0 Then rmps= "num3"
				End if
			End If
		End If
		TestUName = rmps
	End function


	Sub checkemail
		dim UserMail,Rs
		UserMail=HRF(2,1,"email")
		If Not IsValidEmail(UserMail) Then
			Echo "false"
		Else
			Set Rs =team.execute( "Select * from ["& Isforum &"user] where usermail='"& UserMail &"'")
			If Rs.Eof and Rs.Bof Then
				Echo "true"
			Else
				Echo "false"
			End If	
			Rs.Close: Set Rs = Nothing
		End If 
	End Sub

	Sub checkseccode
		Dim Code
		If team.Forum_setting(48)>=1 Then
			Code=HRF(2,1,"seccodeverify")
			If Len(Code) = 4 Then
				if CodeIsTrue(Code) Then
					Echo "true"
				Else
					Echo "false"
				End If
			Else
				Echo "false"
			End If 
		End If
	End Sub

	Function CodeIsTrue(a)
		Dim CodeStr
		CodeStr=Trim(a)
		CodeStr=Trim(a)
		If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
			CodeIsTrue=True
			Session("GetCode")=empty
		Else
			CodeIsTrue=False
			Session("GetCode")=empty
		End If
	End Function

End If
%>