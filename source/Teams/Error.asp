<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Response.Buffer = false
Dim Message,Message1,Message2
Dim X1,X2,Fid,Tmp
Message= Fixjs(Request.QueryString("message"))
Message1= Fixjs(Request.QueryString("message1"))
Message2= Fixjs(Request.QueryString("message2"))
team.Headers("信息提示页面")
X1="论坛提示信息"
X2=""
Echo team.MenuTitle
tmp = Team.ElseHtml (6)
tmp = HtmlEncode(tmp)
If Message&""="" Then
	tmp = TempCode(tmp,"hide")
Else
	tmp = BlackTmp(tmp,"hide")
End If
If Message1&""="" Then
	tmp = TempCode(tmp,"clear")
Else
	tmp = BlackTmp(tmp,"clear")
End if
If Message2&""="" Then
	tmp = TempCode(tmp,"msg")
Else
	tmp = BlackTmp(tmp,"msg")
End if
tmp = iHtmlEncode(tmp)
tmp = Replace(tmp,"{$clubname}",Team.Club_Class(1))
tmp = Replace(tmp,"{$message}",Message)
tmp = Replace(tmp,"{$message1}",Message1)
tmp = Replace(tmp,"{$message2}",Message2)
Echo tmp
If Message2&"" = "" Then 
	Team.Footer
End If
%>