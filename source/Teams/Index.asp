<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim WordsSize,HotSize,ShowSize,fid
Dim BoardListShows,GifBr,GifNum
'====================����������================================
WordsSize		= "25"		'������ʾ�ĳ��ȡ�
HotSize			= "15"		'�����Ƽ���ʾ������
ShowSize		= "15"		'�����Ƽ���ʾ������
BoardListShows	= "all"		'������Ҫ��ʾ�İ�飬Ĭ������Ϊall�������Ҫָ����Ҫ��ʾ�İ�飬����д����ID���ƣ��������á�,�������֣���1,3,5,8,9 ��
GifBr			= "4"		'ָ��ͼƬ��ʾ�ķ�����
GifNum			= "12"		'ָ����Ҫ��ʾ��������
'����������ô��󣬽��������ݲ�ѯ�������ʾ����BoardListShows����������Ϊ��ֵ�͡�
'====================����������=================================
team.Headers(Team.Club_Class(1))
If CID(team.Forum_setting(111)) = 1 Then
	Call Main()
Else
	Response.Redirect "Default.asp"
End if
team.Footer()

Sub Main()
	Dim tmp
	tmp = Team.IndexHtml (11)
	Tmp = Replace(tmp,"{$uslogin1}",IIF(team.UserLoginED,"Display:none",""))
	Tmp = Replace(tmp,"{$uslogin2}",IIF(team.UserLoginED,"","Display:none"))
	Tmp = Replace(tmp,"{$pass}",IIF(team.Forum_setting(17)>=1,"","Display:none"))
	Tmp = Replace(tmp,"{$username}",tk_UserName)
	If team.UserLoginED Then
		Tmp = Replace(tmp,"{$levelname}",team.Levelname(0))
	Else
		Tmp = Replace(tmp,"{$levelname}","�ο�")
	End if
	Tmp = Replace(tmp,"{$�������}",BoardShow)
	Tmp = Replace(tmp,"{$ͼƬ��ʾ}",BoardPicShow)
	Tmp = Replace(tmp,"{$��תͼƬ}",Advs)
	Tmp = Replace(tmp,"{$��̳����}",Newtopic)
	Tmp = Replace(tmp,"{$���Ű��}",BbslistTop)
	Tmp = Replace(tmp,"{$��������}",ToGodTop)
	Tmp = Replace(tmp,"{$��������}",ToHotTop)
	Tmp = Replace(tmp,"{$�û�����}",ShowMembers)
	Dim Links,ShowLinks,i
	Links = team.Forum_Link_Rs
	ShowLinks = ""
	If isarray(Links) Then
		for i = 0 to Ubound(Links,2)
			ShowLinks = ShowLinks & " [<a href="""& Links(1,i) &""" target=""_blank"" title="""& Links(0,i) &""">"& Links(0,i) &"</a>] "
		Next
	End If 
	tmp = Replace(tmp,"{$��������}",ShowLinks)
	Echo tmp
End sub

Function BoardPicShow()
	Dim Rs,i,t,MyID,U,Tmp
	i = 0 : MyID="" : U=0
	Set Rs = team.execute("Select t.ID,t.topic,t.forumid,t.posttime,u.filename,u.ID From ["&IsForum&"Forum] T Inner Join ["&IsForum&"Upfile] u on t.ID=U.ID Where u.ID>0 Order by t.ID desc ")
	Tmp = "<table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""0""  align=""center""><tr class=""tab4"">"
	Do While Not Rs.Eof
		If InStr(MyID,Rs(0)&"||")<=0 Then
			i = i+1 : U = U+1
			Tmp = Tmp & "<td><table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2""><tr><td class=""tab4""><div class=""pic""><a Href=""Thread.asp?tid="&RS(0)&""" target=""_blank""><img Src=""images/Upfile/"&RS(4)&""" Border=""0"" align=""absmillde"" title="""&Fixjs(RS(1))&""" width=""160"" height=""160""></a></div></td></tr></table></td>"
			If U >= CID(GifBr) Then
				Tmp = Tmp & "</tr><tr>"
				U = 0
			End if
			If i >= CID(GifNum) Then Exit Do
		End if
		MyID = MyID & Rs(5) & "||"
		Rs.moveNext
	Loop
	Tmp = Tmp & "</tr></table>"
	Rs.Close:Set Rs = Nothing
	BoardPicShow = tmp
End Function

function Advs()
	Dim Rs,i,Rs1,t,R1,R2,R3,U,gID
	i = 0: U=0
	Set Rs = team.execute("Select t.ID,t.topic,t.forumid,t.posttime,u.filename,u.ID,u.FileSize From ["&IsForum&"Forum] T Inner Join ["&IsForum&"Upfile] u on t.ID=U.ID Where U.types='gif' or U.types='jpg' and u.ID>0 and t.Deltopic=0 order by t.ID desc ")
	R1 = "" : R2="" : R3=""
	Do While Not Rs.Eof
		If gID <> Rs(0) Then
			i = i+1
			If I = 1 Then
				R1 = R1 & "images/Upfile/"& RS(4) : R2 = R2 & "Thread.asp?tid="& RS(0) : R3 = R3 & Fixjs(RS(1))
			Else
				R1 = R1 & "|" & "images/Upfile/"& RS(4) : R2 = R2 & "|" & "Thread.asp?tid="& RS(0) : R3 = R3 & "|" & Fixjs(RS(1))
			End If 
			If i >=5 Then Exit Do
		End If 
		gID = Rs(0)
		Rs.moveNext
	Loop
	Rs.Close:Set Rs = Nothing
    t = t & "  <script type=""text/javascript"">" & vbcrlf
    t = t & "  var swf_width=290;	 " & vbcrlf
    t = t & "  var swf_height=198;"  & vbcrlf
    t = t & "  var config='5|0xffffff|0x0099ff|50|0xffffff|0x0099ff|0x000000';"  & vbcrlf
    t = t & "  // config ���÷ֱ�Ϊ: �Զ�����ʱ��(��)|������ɫ|���ֱ���ɫ|���ֱ���͸����|��������ɫ|��ǰ����ɫ|��ͨ����ɫ"  & vbcrlf
    t = t & "  var files='"& R1 &"';"  & vbcrlf
    t = t & "  var links='"& R2 &"';"  & vbcrlf
    t = t & "  var texts='"& R3 &"';"  & vbcrlf
	t = t & "   document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ swf_width +'"" height=""'+ swf_height +'"">'); "  & vbcrlf
	t = t & "  document.write('<param name=""movie"" value=""adv/focus.swf"" />'); "  & vbcrlf
	t = t & "  document.write('<param name=""quality"" value=""high"" />');"  & vbcrlf
	t = t & "  document.write('<param name=""menu"" value=""false"" />');"  & vbcrlf
	t = t & "  document.write('<param name=wmode value=""opaque"" />');"  & vbcrlf
	t = t & "  document.write('<param name=""FlashVars"" value=""config='+config+'&bcastr_flie='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'"" />');"  & vbcrlf
	t = t & "  document.write('<embed src=""adv/focus.swf"" wmode=""opaque"" FlashVars=""config='+config+'&bcastr_flie='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'"" menu=""false"" quality=""high"" width=""'+ swf_width +'"" height=""'+ swf_height +'"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />');"  & vbcrlf
	t = t & "  document.write('</object>'); </script>"  & vbcrlf
	Advs = t
End Function

Function Fixjs(Strings)
	Dim Str
	Str = Strings
	str = Replace(str, CHR(39), "\'")
	str = Replace(str, CHR(13), "")
	str = Replace(str, CHR(10), "")
	str = Replace(str, "]]>","]]&gt;")
	Fixjs = str
End Function

'�����Ƽ�
function ToHotTop
	Dim Rs,i,tmp,MyChecks
	MyChecks = team.BoardList
	tmp = "<table width=""98%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2"">"
	tmp = tmp & "  <tr class=""a1""><td>  �����Ƽ� </td></tr>"
	Set Rs = team.execute("Select Top "&CID(HotSize)&" ID,topic,forumid,posttime From ["&IsForum&"Forum] Where deltopic=0 Order by Replies Desc")
	Do While Not Rs.Eof
		tmp = tmp & "<tr class=""a4""><td> "
		If IsArray(MyChecks) Then
			For i = 0 To UBound(MyChecks,2)
				If Int(MyChecks(0,i)) = Int(Rs(2)) Then
					tmp = tmp & " <a Href=""Forums.asp?fid="&MyChecks(0,i)&""" target=""_blank"">["&MyChecks(1,i)&"]</a>"
				End If
			Next
		End if
		tmp = tmp & " <a Href=""Thread.asp?tid="&RS(0)&""" title=""����ʱ�䣺"&FormatDatetime(RS(3),1)&""" target=""_blank"">"&Cutstr(RS(1),18)&"</a> </td></tr> "
		Rs.MoveNext
	Loop
	tmp = tmp & "</table>"
	Rs.Close:Set Rs=Nothing
	ToHotTop = tmp
End function

'�������Ƽ�
function ToGodTop()
	Dim Rs,i,tmp,MyChecks
	MyChecks = team.BoardList
	tmp = "<table width=""98%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2"">"
	tmp = tmp &"  <tr class=""tab1""><td>  �����Ƽ� </td></tr>"
	Set Rs = team.execute("Select Top "&CID(ShowSize)&" ID,topic,forumid,posttime From ["&IsForum&"Forum] Where goodtopic=1 Order by Posttime Desc")
	Do While Not Rs.Eof
		tmp = tmp & "<tr class=""a4""><td> "
		If IsArray(MyChecks) Then
			For i = 0 To UBound(MyChecks,2)
				If Int(MyChecks(0,i)) = Int(Rs(2)) Then
					tmp = tmp & " <a Href=""Forums.asp?fid="&MyChecks(0,i)&""" target=""_blank"">["&MyChecks(1,i)&"]</a>"
				End If
			Next
		End if
		tmp = tmp & " <a Href=""Thread.asp?tid="&RS(0)&""" title=""����ʱ�䣺"&FormatDatetime(RS(3),1)&""" target=""_blank"">"&Cutstr(RS(1),18)&"</a> </td></tr> "
		Rs.MoveNext
	Loop
	tmp = tmp & " </table>"
	Rs.Close:Set Rs=Nothing
	ToGodTop = tmp
End Function

function ShowMembers
	Dim Rs,tmp,R
	R = 0
	tmp = "<table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2""><tr><td class=""a1"" colspan=""3""><B>��̳�û�TOP��</B></td></tr>"
	Set Rs = team.execute("Select Top 10 UserName,postrevert,posttopic From ["&IsForum&"User] Order By postrevert+posttopic desc")
	Do While Not Rs.Eof
		R = R+1
		tmp = tmp & "<tr class=""tab4""><td width=""20%""><img src="""& team.Styleurl &"/rank_"& R &".gif""></td><td><a href=""Profile.asp?username="&Rs(0)&""" target=""_blank"">"&Rs(0)&"</a> </td><td> [��������"& RS(1) + Int(Rs(2)) &"]</td></tr>"
		Rs.movenext
	Loop
	Rs.Close:Set Rs = Nothing
	tmp = tmp & "</table>"
	ShowMembers = tmp
End function

'�����ʾ
function BoardShow()
	Dim u,i,tmp,MyChecks
	MyChecks = BoardList
	If IsArray(MyChecks) Then
		u = 0
		tmp = tmp & "<table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""0""  align=""center""><tr>" & vbCrlf
		For i = 0 To UBound(MyChecks,2)	
			If Int(MyChecks(3,i)) <> 0 Then 
					u = u+1
					tmp = tmp & "<td width=""33%"" valign=""top""><table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2""><tr ><td class=""a1""> <a href=Forums.asp?fid="&MyChecks(0,i)&"> "&MyChecks(1,i)&" </a> </td></tr><tr class=""a4""><td>" & vbCrlf
					tmp = tmp & Miniboard(MyChecks(0,i))
					tmp = tmp & "</td></tr></table></td>" & vbCrlf
					If u >2 Then
						tmp = tmp & " </tr><tr>" & vbCrlf
						u = 0
					End If
			End If
		Next
		tmp = tmp &"</table>"
	End If
	BoardShow = tmp
End Function

Function BoardList()	
	Dim Rs,Moderuser
	Cache.Reloadtime = 30
	Cache.Name = "myBoardLists"
	If Cache.ObjIsEmpty() Then
		If BoardListShows = "all" Then
			Set Rs=team.Execute("Select ID,bbsname,Board_Model,Followid,Readme,today,toltopic,tolrestore,icon,Board_Last,Pass,Board_URL From ["&IsForum&"Bbsconfig] Where Hide=0 Order By toltopic desc,tolrestore desc")
		Else
			Set Rs=team.Execute("Select ID,bbsname,Board_Model,Followid,Readme,today,toltopic,tolrestore,icon,Board_Last,Pass,Board_URL From ["&IsForum&"Bbsconfig] Where Hide=0 and ID In ("&BoardListShows&") Order By toltopic desc,tolrestore desc")
		End if
	   	If RS.Eof Then
			Exit Function
		Else
	      	Cache.Value = Rs.GetRows(-1)
	   	End If
		Rs.Close:Set Rs=Nothing
	End If
	BoardList = Cache.Value
End Function

Function Miniboard(a)
	Dim Rs,tmp
	Cache.Reloadtime = 0
	Cache.Name = "ForumPostLists_"&a
	If Cache.ObjIsEmpty() Then
		Set Rs = team.execute("Select Top 10 ID,topic,forumid,posttime From ["&IsForum&"Forum] Where forumid="&Int(a)&" and deltopic=0 Order by Lasttime Desc")
		Do While Not Rs.Eof
			tmp = tmp & "<a Href=""Thread.asp?tid="&RS(0)&""" title=""����ʱ�䣺"&Rs(3)&""" target=""_blank"">"&Cutstr(RS(1),CID(WordsSize))&"</a> <BR>" & vbCrlf
			Rs.Movenext
		Loop
		Cache.Value = tmp
		Rs.Close:Set Rs = Nothing
	End If
	Miniboard = Cache.Value
End Function

Function Newtopic()
	Dim Rs,i,tmp,MyChecks
	MyChecks = team.BoardList
	tmp = tmp & "<table width=""98%"" border=""0"" cellpadding=""3"" cellspacing=""1""  align=""center"" class=""a2"">"
	tmp = tmp & " <tr class=""a1""><td colspan=""2""> ��̳���� </td></tr>"
	Set Rs = team.execute("Select Top 20 ID,topic,forumid,posttime,UserName From ["&IsForum&"Forum] Where deltopic=0  Order by Posttime Desc")
	Do While Not Rs.Eof
		tmp = tmp & "<tr class=""a4""><td><li>"
		If IsArray(MyChecks) Then
			For i = 0 To UBound(MyChecks,2)
				If Int(MyChecks(0,i)) = Int(Rs(2)) Then
					tmp = tmp & " <a Href=""Forums.asp?fid="&MyChecks(0,i)&""" target=""_blank"">["&MyChecks(1,i)&"]</a>"
				End If
			Next
		End if
		tmp = tmp & " <a Href=""Thread.asp?tid="&RS(0)&""" target=""_blank"" title=""�����ߣ�"&RS(4)&""">"&CutStr(RS(1),30)&"</a></li></td><td align=""center"">["&FormatDatetime(RS(3),1)&"]</td></tr>"
		Rs.MoveNext
	Loop
	tmp = tmp & "</table>"
	Rs.Close:Set Rs=Nothing
	Newtopic = tmp
End Function

Function BbslistTop()
	Dim tmp		
	Dim Rs,R
	tmp = tmp & "<table width=""100%""  border=""0"" cellpadding=""3"" cellspacing=""1"" align=""center"" class=""a2"">"
	tmp = tmp & "<tr class=""a1""><td colspan=""3"">��̳���а�</td></tr>"
	tmp = tmp & "<tr class=""tab3""><td width=""10%""></td><td>�������</td><td>����</td></tr>"
	R = 0
	Set Rs = team.execute("Select Top 9 ID,bbsName,today From ["&IsForum&"bbsConfig] Where hide=0 and followid>0 Order By today Desc")
	Do While Not Rs.Eof
		R = R+1
		tmp = tmp & "<tr class=""tab4""><td> <img src="""& team.Styleurl &"/rank_"& R &".gif""> </td><td><a href=""Forums.asp?fid="&RS(0)&"""> "& Rs(1) &"</a> </td><td> "& Rs(2) &" </td></tr>"
		If R = 10 Then Exit Do
		Rs.moveNext
	Loop
	Rs.Close:Set Rs=Nothing	
	tmp = tmp & " </table>"
	BbslistTop = tmp
End Function
%>

