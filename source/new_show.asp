<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="DA_CHMRW/CHMRWB_index.asp" -->
<%
if request("id")="" then
	response.Write("�����������������ϵ��վ������Ա!")
	response.End()
end if
set rs=server.CreateObject("adodb.recordset")
sql="select * from new where id="&request("id")
rs.open sql,conn,1,1
if not rs.bof and not rs.eof then
	tag=rs("tag")
	title=rs("title")
	content=rs("content")
else
	response.Write("��¼�Ѿ���ɾ�����뷵��!")
	response.End()
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="��ʵ��������,�����˴�ִ��̤ʵ��������ʵ��ҵ�������޹�˾�ĺ��Ĳ�ҵ�Ƿ��ز�����,�ѿ������ɵ���ɽׯ�����俥�����������⡢���ֺ�ɽ��סլС��������30����ƽ����">
<meta name="keywords" content="��ʵ���ţ���ʵ�����ֺ�ɽ������¥�̣��������ز������ز�" >
<title><%=title&"-"%>��ʵ����</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
<script src="scripts/FancyZoom.js" language="JavaScript" type="text/javascript"></script>
<script src="scripts/FancyZoomHTML.js" language="JavaScript" type="text/javascript"></script>

</head>

<body onLoad="setupZoom();">

<script type="text/javascript" src="scripts/banner.js"></script>
<script type="text/javascript" src="dhtml.js"></script>
<script src="scripts/AC_RunActiveContent.js" type="text/javascript"></script>

<div id="DsTop">
	<img src="images/logo.gif" />
    <ul>
    	<li><a href="default.asp" id="menu1"><img src="images/top1.jpg" /></a></li>
        <li><a href="news.asp" id="menu2"><img src="images/top2.jpg" /></a></li>
        <li><a href="about.asp" id="menu3"><img src="images/top3.jpg" /></a></li>
        <li><a href="brand_list.asp" id="menu4"><img src="images/top4.jpg" /></a></li>
        <li><a href="estate.asp" id="menu5"><img src="images/top5.jpg" /></a></li>
        <li><a href="teams/default.asp" id="menu8"><img src="images/top8.jpg" /></a></li>
        <li><a href="join.asp" id="menu6"><img src="images/top6.jpg" /></a></li>
        <li><a href="contact.asp" id="menu7"><img src="images/top7.jpg" /></a></li>
    </ul>
</div>
<script type="text/javascript" src="dropdown_initialize.js"></script>
<div id="DsBanner">
  

<script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','992','height','312','style','z-index:0;','src','images/Inbanner5','wmode','transparent','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','images/Inbanner5' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="992" height="312" style="z-index:0;">
    <param name="movie" value="images/Inbanner5.swf" />
    <param name="quality" value="high" />
	<param name="wmode" value="transparent">
    <embed src="images/Inbanner5.swf" wmode="transparent" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="992" height="312"></embed>
  </object></noscript></noscript>
  
  
  
</div>

 
<script>
window.onload=function(){
 init();
 setupZoom();
}
</script>

<div id="big"></div>

<div id="Main">
	<div id="Main_left">
    	<em><img src="images/news.jpg" /></em>
        <ul>
        	<li><a href="news.asp">��˾����</a></li>
            <li><a href="news.asp?act=2">��ҵ����</a></li>
            <li><a href="news.asp?act=3">ý�屨��</a></li>
        </ul>
    </div>
    <div id="Main_right">
    <%if tag=1 then%>
   		 <img  src="images/news1.jpg" class="guide" />
    <%elseif tag=2 then%>
    	<img  src="images/news2.jpg" class="guide" />
    <%elseif tag=3 then%>
    	<img  src="images/news3.jpg" class="guide" />
    <%end if%>
     	
        <div class="Main_right_content">
        	<h1><%=title%></h1>
			<div class="newcon">
            <%=content%>
            </div>
        </div>
    </div>
</div>

















<!--#include file="bottom.asp" -->

</body>
</html>