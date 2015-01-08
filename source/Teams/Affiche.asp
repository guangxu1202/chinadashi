<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim x1,x2,fID
team.Headers(Team.Club_Class(1) &" - 公告栏")
Call Main
team.Footer()
Sub Main
	Dim tmp,AfficheTitle,i,AffTemp
	x1 = "<a href=""Affiche.asp"">公告栏</a>"
	tmp = Replace(Team.ElseHtml (3),"{$weburl}",team.MenuTitle)
	tmp = Replace(tmp,"{$affadv}",team.AdvShows(7,0))
	AfficheTitle = Team.Affiche()
	If IsArray(AfficheTitle) Then
		For i=0 To Ubound(AfficheTitle,2)
			AffTemp = AffTemp & "<table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""98%"" align=""center"" class=""a2"">" & vbCrlf
			AffTemp = AffTemp & " <tr class=""tab1""> " & vbCrlf
			AffTemp = AffTemp & "    <td colspan=""2""> <a name="""&AfficheTitle(0,i)&"""></a> <img src="""&team.styleurl&"/tip.gif"" align=""absmiddle"" border=""0""> <Span Style='"&AfficheTitle(5,i)&"'> "&AfficheTitle(1,i)&" </Span> <img src="""&team.styleurl&"/tip.gif"" align=""absmiddle"" border=""0""></td>" & vbCrlf
			AffTemp = AffTemp & " </tr>" & vbCrlf
			AffTemp = AffTemp & " <tr class=""a4"">" & vbCrlf
			AffTemp = AffTemp & "    <td colspan=""2"" height=""50""> "&UBB_Code(AfficheTitle(2,i))&"</td>" & vbCrlf
			AffTemp = AffTemp & " </tr>" & vbCrlf
			AffTemp = AffTemp & " <tr class=""a3"">" & vbCrlf
			AffTemp = AffTemp & "    <td width=""50%"">作者："&AfficheTitle(3,i)&"</td><td align=""right"">发布时间："&AfficheTitle(4,i)&"</td>" & vbCrlf
			AffTemp = AffTemp & " </tr> " & vbCrlf
			AffTemp = AffTemp & "</table><br>" & vbCrlf
		Next
		Tmp = Replace(Tmp,"{$roundaff}",AffTemp)
	Else
		Tmp = Replace(Tmp,"{$roundaff}",AfficheTitle)
	End If
	Echo tmp
End Sub
%>
