<%
Response.Write " <link rel=""stylesheet"" rev=""stylesheet"" href=""skins/teams/bbs.css"" type=""text/css"" media=""all"" />"
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""8"" width=""80%"" align=""center"" class=""a2""><tr class=""tab1""><td>"
Select Case Request("action")
	Case "ipclose"
		Response.Write "<H3>您的IP位于受限IP段,您没有查看论坛的权限.</H3>"
	Case "upower"
		Response.Write "<H3>您所在的组没有查看论坛的权限</H3>"
	Case Else
		Response.Write "<H3>您没有查看论坛的权限</H3>"
End Select
Response.Write "</td></tr></table>"
%>

