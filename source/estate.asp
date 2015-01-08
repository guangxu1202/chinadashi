<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="大实集团理念,有容乃大、执诚踏实。大连大实企业集团有限公司的核心产业是房地产开发,已开发建成叠翠山庄、叠翠骏景、泊林阳光、泊林和山等住宅小区，共计30多万平方米">
<meta name="keywords" content="大实集团，大实，泊林和山，大连楼盘，大连房地产，房地产" >
<title>大实集团-大实产业</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
</head>

<body>

<!--#include file="top.asp" -->



<div id="Main">
	<div id="Main_left">
    	<em><img src="images/estate.jpg" /></em>
        <ul>
        	<li><a href="estate.asp">房地产开发</a></li>
            <li><a href="estate.asp?act=1">其他产业</a></li>
        </ul>
    </div>
    <div id="Main_right">
     <%if request("act")="" then%>
    		<img  src="images/estate1.jpg" class="guide" />
	<%elseif request("act")=1 then%>
    		<img  src="images/estate2.jpg" class="guide" />
    <%end if%>
        <div class="Main_right_content">
        <%if request("act")="" then%>
      <p> <strong>核心产业</strong></p>
      <p>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;大连大实企业集团有限公司的核心产业是房地产开发。集团下辖大连大实企业集团房屋开发有限公司是房地产开发的核心企业。其成立9年，已开发建成叠翠山庄、叠翠骏景、泊林阳光、泊林和山等住宅小区，共计30多万平方米。在建泊林映山住宅小区10万平方米，规划建设景山“海·风·景”18万平方米。外埠开发分别在四川乐山和辽宁沈阳两地，共计开发30万平米。伴随大实房地产开发一起成长的是大连弘实物业管理有限公司，它担负着大实集团房屋开发有限公司开发的所有住宅小区的物业管理。</p><br />
<p> <strong>企业产品</strong></p>
<p> 
1.	叠翠山庄小区<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积8万平米，2004年建设完工。<br />
2.	叠翠骏景小区<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积9万平米，2006年建设完工。<br />
3.	泊林阳光小区<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积6万平米，2007年建设完工。<br />
4.	泊林和山小区<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积7万平米，2008年建设完工。<br />
5.	泊林映山小区<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积10万平米，2009年开工在建。<br />
6.	黑石礁星海湾畔——海·风·景<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积18万平米，筹划中。<br />
7．外埠开发<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;（1）四川乐山“名雅花园”<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积近10万平米，2001年完工。<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;（2）沈阳“大华·水岸福邸”<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建筑面积近22万平米，2007年完工。<br />
</p>
<%elseif request("act")=1 then%>

<p>大连大实汇友商贸有限公司</p>
<p>大连万嘉机电安装有限公司</p>
<p>大实德国中国中心有限公司</p>
<p>大连弘实物业管理有限公司</p>

<%end if%>
        </div>
    </div>
</div>


















<!--#include file="bottom.asp" -->

</body>
</html>