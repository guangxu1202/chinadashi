<!--#include file="../DA_CHMRW/CHMRWB.asp"-->

<%

id=request.QueryString("id")
fdown=request.QueryString("filedown")
set rs=server.CreateObject("adodb.recordset")
sql="select * from project where id="&id
rs.open sql,conn,1,1
if not rs.bof and not rs.eof then

	filename=rs("filename4")
	filedown=rs("filedown4")
	filetype=rs("filetype4")
end if
rs.close
set rs=nothing


	n=split(filename,",")
	a=split(filedown,",")
	b=split(filetype,",")
	for i=lbound(n) to ubound(n)-1
		if a(i)=fdown then

		else
			fd=fd&a(i)&","
			ft=ft&b(i)&","
			fn=fn&n(i)&","
		end if
	next
	
	
	
set rs=server.CreateObject("adodb.recordset")
sql="select * from project where id="&id
rs.open sql,conn,1,3
if not rs.bof and not rs.eof then

	rs("filename4")=fn
	rs("filedown4")=fd
	rs("filetype4")=ft

end if
rs.update
rs.close
set rs=nothing

response.redirect "pic_xmhx.asp?id="&id

%>