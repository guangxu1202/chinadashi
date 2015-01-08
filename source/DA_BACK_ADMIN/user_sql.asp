<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<!--#include file="../DA_CHMRW/MD5.asp" -->
<%
function zero(text)
    zero=text
 if text="" then zero="0"
end function
uid=trim(request.form("uid"))
upwd=md5(trim(request.form("upwd")))
uname=trim(request.form("uname"))
xwzx=zero(request.form("xwzx"))
zx=request.form("zx")
rcgl=zero(request.form("rcgl"))
ulevel=zero(request.form("ulevel"))
act=request("act")
if act="add" then

   set rs = server.CreateObject("adodb.recordset")
      sql = "select id from users where uid='"&uid&"'"
	  rs.open sql,conn,1,1
	  if not rs.bof and not rs.eof then
		 %>
		 <script language="JavaScript" type="text/JavaScript">
		   alert("用户名已存在！");
           history.go(-1);		 
		 </script>
		 <%
		 response.end
	  end if	
   rs.close
   set rs = nothing


   set rs= server.CreateObject("adodb.recordset")
       sql = "select * from users"
	   rs.open sql,conn,3,3
	   rs.addnew
	  rs("uid")=uid
	  rs("upwd")=upwd
	  rs("uname")=uname
	  rs("ulevel")=ulevel
	  rs.update
      rs.close
      set rs = nothing
end if

if act="edit" then
     id=request.form("id")
	
	 set rs = server.CreateObject("adodb.recordset")
      sql = "select id from users  where uname='"&uname&"' and id<>"&id
	  rs.open sql,conn,1,1
	  if not rs.bof and not rs.eof then
		 %>
		 <script language="JavaScript" type="text/JavaScript">
		   alert("昵称已存在！");
           history.go(-1);		 
		 </script>
		 <%
		 response.end
	  end if	
   rs.close
   set rs = nothing
	 
	 set rs = server.CreateObject("adodb.recordset")
     sql = "select * from users where id="&id
	  rs.open sql,conn,3,3
	  rs("uname")=uname
	  if zx<>"" then
		rs("zx")=zx
	  end if
	  rs("ulevel")=ulevel
	  rs.update
      rs.close
      set rs = nothing

end if

if act="dele" then 
id=request.QueryString("id")
  
  sql="delete * from users where id="&id
    conn.execute (sql)
end if

response.redirect "user_main.asp"
%>