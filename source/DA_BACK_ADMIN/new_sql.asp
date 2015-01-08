<!--#include file="../DA_CHMRW/CHMRWB.asp" -->
<!--#include file="../DA_CHMRW/replace.asp" -->
<%
title=request.Form("title")
content=replace_t(request.Form("content"))
id=request.Form("id")
tag=request.Form("tag")
act=request.Form("act")

if act = "add" then
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from new"
	rs.open sql,conn,3,3
	rs.addnew
		rs("title")=title
		rs("content")=content
		rs("tag")=tag
		rs("sendtime")=now()
		rs("lrr")=session("love_uname")
	rs.update
	rs.close
	set rs=nothing
	
end if

if act = "edit" then
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from new where id="&id
	rs.open sql,conn,3,3
		rs("title")=title
		rs("tag")=tag
		rs("content")=content
	rs.update
	rs.close
	set rs=nothing
end if

response.Redirect("new_main.asp?mm=1")
%>