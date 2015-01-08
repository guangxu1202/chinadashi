<!--#include file="../DA_CHMRW/CHMRWB.asp" -->
<!--#include file="../DA_CHMRW/MD5.asp" -->

<%
oldmm=md5(trim(request.Form("oldmm")))
newmm=md5(trim(request.Form("newmm")))
newcmm=md5(trim(request.Form("newcmm")))
set rs=server.CreateObject("adodb.recordset")
    sql="select * from users where id="&session("love_id")
      rs.open sql,conn,3,3
        if oldmm<>rs("upwd") then
		    %><script>alert("密码错误!");history.go(-1)</script><%
			response.End()
		end if
		
		if newmm<>newcmm then
		    %><script>alert("新密码不一致!");history.go(-1)</script><%
			response.End()
		end if
	rs("upwd")=newmm
	rs.update
	rs.close
	set rs=nothing
	%>
	<script>
	alert("密码修改成功!下次登录请使用新密码!!")
	window.location.href("default.asp")
	</script>
