<%
set conn=Server.CreateObject("adodb.connection")
   'connString="Provider=sqloledb;User Id=sa;PASSWORD=2004;Initial Catalog=dllnb;Data source=jta"
   
   connString ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("DINSXC/DA_afkpuz.mdb")
	'connString="dsn=dllnb;uid=jta;pwd=2006;"
	
    conn.open connString
%>