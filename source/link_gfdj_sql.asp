<!--#include file="DA_CHMRW/CHMRWB_index.asp" -->
<%
if request("act")="addj" then
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from khly"
	rs.open sql,conn,1,3
	rs.addnew
		rs("Uname")=request.Form("Uname")
		rs("Uxb")=request.Form("Uxb")
		rs("Utel")=request.Form("Utel")
		rs("Udz")=request.Form("Udz")
		rs("Uyb")=request.Form("Uyb")
		rs("Umail")=request.Form("Umail")
		rs("Place")=request.Form("Place")
		rs("zone")=request.Form("zone")
		rs("Utime")=request.Form("Utime")
		rs("Udsyz")=request.Form("Udsyz")
		rs("Ugfjl")=request.Form("Ugfjl")
		rs("Ubz")=request.Form("Ubz")
		rs("sendtime")=now
	rs.update
	rs.close
	set rs=nothing
	response.Redirect("link_gfdj.asp?msg=1")
end if
%>
