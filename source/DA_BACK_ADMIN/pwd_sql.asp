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
		    %><script>alert("�������!");history.go(-1)</script><%
			response.End()
		end if
		
		if newmm<>newcmm then
		    %><script>alert("�����벻һ��!");history.go(-1)</script><%
			response.End()
		end if
	rs("upwd")=newmm
	rs.update
	rs.close
	set rs=nothing
	%>
	<script>
	alert("�����޸ĳɹ�!�´ε�¼��ʹ��������!!")
	window.location.href("default.asp")
	</script>
