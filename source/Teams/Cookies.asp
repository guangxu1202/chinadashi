<%
Select Case Request("action")
	Case "skins"
		Response.Cookies("Style")("skins")=Request("styleid")
		Response.Cookies("Style").expires= date+365
		Response.Redirect Request.ServerVariables("http_referer")
End Select
%>

