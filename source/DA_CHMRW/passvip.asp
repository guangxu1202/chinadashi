<!--#include file="CHMRWB.asp" -->
<!--#include file="MD5.asp" -->

<%
vip_kh=trim(request.Form("vip_kh"))
vip_mm=md5(trim(request.Form("vip_mm")))


set rs=server.CreateObject("adodb.recordset")
sql="select * from cl where cl_zx <>1 "
rs.open sql,conn,1,1
	for i = 1 to rs.recordcount
		if rs("cl_zh")=vip_kh then
			cl_mc=rs("cl_mc")
			cl_id=rs("id")
			cl_pwd=rs("cl_pwd")
		else
		end if
	if rs.eof then exit for
	rs.movenext
	next
rs.close
set rs=nothing

if cl_mc="" or cl_pwd<>vip_mm then
%>
<script>
alert("帐号或密码错误，请重新输入！")
history.back()
</script>
<%
else

		session("client_vip")=vip_kh
		session("client_id")=cl_id


		response.Redirect("../index.asp")

end if

%>