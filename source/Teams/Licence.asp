<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim X1,X2,Fid,acc,Webname,WebUrl
X1="�汾˵��"
team.Headers(Team.Club_Class(1))
Webname = Server.UrlEncode(Team.Club_Class(1))
WebUrl = Server.UrlEncode(Request.ServerVariables("server_name"))
If IsSqlDataBase = 1 Then
	acc="SQL"
Else
	acc="ACC"
End If
Echo team.MenuTitle
Echo "<table cellspacing=""1"" cellpadding=""10"" width=""98%"" align=""center"" class=""a2"">"
Echo " <tr class=""a6"" align=""center"">"
Echo " 	<td> <a href=""http://www.team5.cn"">TEAM��̳�汾ָ�� </a></td>"
Echo " </tr>"
Echo " <tr class=""a4"">"
Echo " 	<td>"
Echo " 	<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%""> "
Echo " 	<tr class=""tab4"">"
Echo " 		<td width=""50%"">��վ����:</TD>"
Echo " 		<td align=""left""> " & Team.Club_Class(1) &" </TD> "
Echo " 	</tr>"
Echo " 	<tr class=""tab4"">"
Echo " 		<td>��վ��ַ:</TD> "
Echo " 		<td align=""left""> " & Team.Club_Class(2) &" </TD> "
Echo " 	</tr>"
Echo " 	<tr class=""tab4"">"
Echo " 		<td>��̳�汾:</TD> "
Echo " 		<td align=""left""><a href=""http://www.team5.cn""> team 2.0.5 (Build "& team.iBuild &") - "& acc &" </a></TD>"
Echo " 	</tr> "
Echo " 	<tr class=""tab4"">"
Echo " 		<td>ע������:</TD> "
Echo " 		<td align=""left""><span id=""regdate"">Loading...</span></TD> "
Echo " 	</tr> "
Echo " 	</table>"
Echo " 	</TD>"
Echo " </tr></table><br><center><input onclick=""history.back(-1)"" type=""submit"" value=""�� �� �� һ ҳ"" name=""Submit""></center><script src=""http://www.team5.cn/ck/webreg.asp?id=regdate&webname="& Webname&"&url="&WebUrl&"""></script> "
Team.footer
%>