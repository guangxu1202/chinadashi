<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="大实集团理念,有容乃大、执诚踏实。大连大实企业集团有限公司的核心产业是房地产开发,已开发建成叠翠山庄、叠翠骏景、泊林阳光、泊林和山等住宅小区，共计30多万平方米">
<meta name="keywords" content="大实集团，大实，泊林和山，大连楼盘，大连房地产，房地产" >
<title>大实集团-产品品牌</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
<script src="scripts/FancyZoom.js" language="JavaScript" type="text/javascript"></script>
<script src="scripts/FancyZoomHTML.js" language="JavaScript" type="text/javascript"></script>
</head>

<body>
<!--#include file="DA_CHMRW/CHMRWB_index.asp" -->
<!--#include file="top.asp" -->

<div id="big"></div>

<div id="Main">
	<div id="Main_left">
    	<em><img src="images/dspp.jpg" /></em>
        <ul>
        	<li><a href="brand.asp">大实理念</a></li>
            <li><a href="brand.asp?act=2">大实之道</a></li>
            <li><a href="brand.asp?act=3">大实文化</a></li>
            <li><a href="brand_list.asp">产品品牌</a></li>
        </ul>
    </div>
    <%
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from project where id="&request("id")
	rs.open sql,conn,1,1
	%>
    <div id="Main_right">
     	<img  src="images/estate_cypp.jpg" class="guide" />
        <div class="Main_right_brand">
        	<div class="Main_right_back"><a href="#" onclick="javascript:history.go(-1)"><img src="images/back.jpg" /></a></div>
            <div class="Main_right_guide">&nbsp;&nbsp;+<%=rs("xmmc")%></div>
			<div class="Main_right_pic">
            <%if rs("filename2")<>"" then%>
            	<img src="../upload/<%=rs("filedown2")&rs("filetype2")%>" width="702" height="220" />
            <%else%>
                <img src="images/brand_plys_pic.jpg" />
            <%end if%>
          </div>
            <div class="Main_right_guide">&nbsp;&nbsp;-项目概况</div>
          <ul>
           	<li><img src="images/point.gif" /><strong>售楼热线：</strong><%=rs("slrx")%></li>
                <li><img src="images/point.gif" /><strong>产品价格：</strong><%=rs("cpjg")%></li>
                <li><img src="images/point.gif" /><strong>占地面积：</strong><%=rs("zdmj")%></li>
                <li><img src="images/point.gif" /><strong>建筑面积：</strong><%=rs("jzmj")%></li>
                <li><img src="images/point.gif" /><strong>规划用途：</strong><%=rs("ghyt")%></li>
                <li><img src="images/point.gif" /><strong>户型面积：</strong><%=rs("hxmj")%></li>
                <li><img src="images/point.gif" /><strong>入住时间：</strong><%=rs("rzsj")%></li>
                <li><img src="images/point.gif" /><strong>公交线路：</strong><%=rs("gjxl")%></li>
                <li><img src="images/point.gif" /><strong>项目区位：</strong><%=rs("xmqw")%></li>
                <li><img src="images/point.gif" /><strong>周边配套：</strong><%=rs("zbpt")%></li>
                <li><img src="images/point.gif" /><strong>项目介绍：</strong><%=rs("xmjs")%></li>
                <li><img src="images/point.gif" /><strong>产品信息：</strong><%=rs("cpxx")%></li>
          </ul>
            <div class="Main_right_guide">&nbsp;&nbsp;-前沿动态</div>
            <%if rs("qydt")<>"" then%>
            <ul>
            	<li><%=rs("qydt")%></li>
            </ul>
            <%end if%>
            <div class="Main_right_guide">&nbsp;&nbsp;-精品图片</div>
            <div class="Main_right_Picguide">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/point.gif" /><strong>项目区位</strong></div>
            <script language=JavaScript src="images/scroll.js"></script>
            	<div class="Main_right_xmqw">
                <%if rs("filename5")<>"" then%>
                	<img src="../upload/<%=rs("filedown5")&rs("filetype5")%>" />
                <%end if%>
                </div>
            <div class="Main_right_Picguide">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/point.gif" /><strong>项目实景</strong></div>
            	<div class="Main_right_xmsj">
                
                <%if rs("filename3")<>"" then
                	n=split(rs("filename3"),",")
				
                    a=split(rs("filedown3"),",")
                    b=split(rs("filetype3"),",")
                    for i=lbound(n) to ubound(n)-1
					%>
                    	
                      <a href="../upload/<%=a(i)&b(i)%>"><img src="../upload/images/<%=a(i)&b(i)%>"  class="imgg"/></a>
                    <%
                    next
				%>
                <%end if%>
                 </div>
            <div class="Main_right_Picguide">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/point.gif" /><strong>项目户型</strong></div>
            	<div class="Main_right_xmhx">
                <%if rs("filename4")<>"" then
                	n=split(rs("filename4"),",")
				
                    a=split(rs("filedown4"),",")
                    b=split(rs("filetype4"),",")
                    for i=lbound(n) to ubound(n)-1
					%>
                    	
                      <a href="../upload/<%=a(i)&b(i)%>"><img src="../upload/images/<%=a(i)&b(i)%>"  class="imgg"/></a>
                    <%
                    next
				%>
                <%end if%>
                </div>
        </div>
    </div>
    <%
	rs.close
	set rs=nothing
	%>
</div>

















<!--#include file="bottom.asp" -->

</body>
</html>