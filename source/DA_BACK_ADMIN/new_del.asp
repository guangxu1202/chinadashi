<!--#include file="../DA_CHMRW/CHMRWB.asp"-->

<%

act = request.form("act")

if act="del" then

   for each obj in request.Form
       'response.write "obj:"&obj&"----reqeust.form(obj):"&request.Form(obj)&"<br>"
      if obj = request.Form(obj) then
	     sql = "delete from new where id="&obj
		 conn.execute(sql)
	  end if
       
   next

response.redirect "new_main.asp?mm=1"
end if
%>