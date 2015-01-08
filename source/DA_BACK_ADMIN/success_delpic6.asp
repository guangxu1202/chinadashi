<!--#include file="../DA_CHMRW/CHMRWB.asp"-->

<%

id=request.QueryString("id")
fdown=request.QueryString("filedown")
set rs=server.CreateObject("adodb.recordset")
sql="select * from project where id="&id
rs.open sql,conn,1,1
if not rs.bof and not rs.eof then

	filename=rs("filename6")
	filedown=rs("filedown6")
	filetype=rs("filetype6")
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

	rs("filename6")=fn
	rs("filedown6")=fd
	rs("filetype6")=ft

end if
rs.update
rs.close
set rs=nothing

response.redirect "pic_cpsc.asp?id="&id

%>