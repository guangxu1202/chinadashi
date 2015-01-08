<!--#include file="../DA_CHMRW/CHMRWB.asp"-->

<%

act = request.form("act")

if act="del" then

   for each obj in request.Form
       'response.write "obj:"&obj&"----reqeust.form(obj):"&request.Form(obj)&"<br>"
      if obj = request.Form(obj) then
	     sql = "delete from khly where id="&obj
		 conn.execute(sql)
	  end if
       
   next

response.redirect "cl_main.asp?mm=2"
end if
%>