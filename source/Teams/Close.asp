<%
Response.Write " <link rel=""stylesheet"" rev=""stylesheet"" href=""skins/teams/bbs.css"" type=""text/css"" media=""all"" />"
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""8"" width=""80%"" align=""center"" class=""a2""><tr class=""tab1""><td>"
Select Case Request("action")
	Case "ipclose"
		Response.Write "<H3>����IPλ������IP��,��û�в鿴��̳��Ȩ��.</H3>"
	Case "upower"
		Response.Write "<H3>�����ڵ���û�в鿴��̳��Ȩ��</H3>"
	Case Else
		Response.Write "<H3>��û�в鿴��̳��Ȩ��</H3>"
End Select
Response.Write "</td></tr></table>"
%>

