<style>
#reset a:link{ color:#000000;}
#reset a:hover{ color: #FF9900; text-decoration:underline;}
#reset a:visited{ color: #FF9900; text-decoration:underline;}
</style>
<%
if session("love_id")="" then 
   response.write "<span id='reset'>ÇëÖØĞÂ<a href='login.asp'>µÇÂ½</a>£¡</span>"
   response.end
end if
%>