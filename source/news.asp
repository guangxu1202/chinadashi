<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="��ʵ��������,�����˴�ִ��̤ʵ��������ʵ��ҵ�������޹�˾�ĺ��Ĳ�ҵ�Ƿ��ز�����,�ѿ������ɵ���ɽׯ�����俥�����������⡢���ֺ�ɽ��סլС��������30����ƽ����">
<meta name="keywords" content="��ʵ���ţ���ʵ�����ֺ�ɽ������¥�̣��������ز������ز�" >
<title>��ʵ����-��������</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
<script src="scripts/FancyZoom.js" language="JavaScript" type="text/javascript"></script>
<script src="scripts/FancyZoomHTML.js" language="JavaScript" type="text/javascript"></script>
<!--#include file="DA_CHMRW/CHMRWB_index.asp" -->
<%
Page=Request.QueryString("page")
If Page=Empty then Page=1
%>
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
    <%if request("act")="" then%>
   		 <img  src="images/news1.jpg" class="guide" />
    <%elseif request("act")=2 then%>
    	<img  src="images/news2.jpg" class="guide" />
    <%elseif request("act")=3 then%>
    	<img  src="images/news3.jpg" class="guide" />
    <%end if%>
     	
        <div class="Main_right_content">
        	<ul class="news">
            <%
			
			set rs=server.CreateObject("adodb.recordset")
			
			if request("act")="" then
				sql="select * from new where tag=1 order by id desc"
			else
				sql="select * from new where tag="&request("act")&" order by id desc"
			end if
			rs.open sql,conn,1,1
			if not rs.bof and not rs.eof then
				 sum = rs.recordcount
				 rs.pagesize = 8
				 last = rs.pagecount
				 if cint(page) > last  then page = last
				 rs.AbsolutePage = page
			for i = 1 to rs.pagesize
			%>
            	<li> <img src="images/point.gif" style="float:left; margin-top:6px;" /><a href="new_show.asp?id=<%=rs("id")%>"><%=rs("title")%></a><em><%=year(rs("sendtime"))&"-"&month(rs("sendtime"))&"-"&day(rs("sendtime"))%></em></li>
            <%
			rs.movenext
			if rs.eof then exit for
			next
			else%>
				<li> ���޼�¼</li>
			<%
			end if
			rs.close
			set rs=nothing
			%>
            </ul>
			<div>
            	<%if sum>8 then%>
                      &nbsp;<a href="news.asp?page=1&id=<%=request("id")%>"><img src="images/doubleleft.gif" alt="��һҳ" class="nborder" /></a>&nbsp;<a href="news.asp?page=<%=Page-1%>&id=<%=request("id")%>"><img src="images/singleleft.gif" alt="��һҳ" class="nborder" /></a>
                      <%
	for j = 1 to page-1%>
                      <%if page-j>5 then%>
                      <%else%>
                      &nbsp;<a href="news.asp?page=<%=j%>&id=<%=request("id")%>"><%=j%></a>
                      <%end if%>
                      <%next%>
                      &nbsp;<a style="font-weight:bold;"><%=page%></a>
                      <%for h = page+1 to page+5%>
                      <%if h>last then%>
                      <%else%>
                      &nbsp;<a href="news.asp?page=<%=h%>&id=<%=request("id")%>"><%=h%></a>
                      <%end if%>
                      <%next%>
                      <a href="news.asp?page=<%=Page+1%>&id=<%=request("id")%>"><img src="images/singleright.gif" alt="��һҳ" class="nborder"/></a>&nbsp;<a href="news.asp?page=<%=last%>&id=<%=request("id")%>"><img src="images/doubleright.gif" alt="���һҳ" class="nborder"/></a>
                     ��ҳ��<%=page%>/<%=last%>
                    <%end if%>
            </div>
        </div>
    </div>
</div>

















<!--#include file="bottom.asp" -->

</body>
</html>