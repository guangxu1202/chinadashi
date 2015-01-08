<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="DA_CHMRW/CHMRWB_index.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="大实集团理念,有容乃大、执诚踏实。大连大实企业集团有限公司的核心产业是房地产开发,已开发建成叠翠山庄、叠翠骏景、泊林阳光、泊林和山等住宅小区，共计30多万平方米">
<meta name="keywords" content="大实集团，大实，泊林和山，大连楼盘，大连房地产，房地产" >
<title>大实集团-快速链接-产品视窗</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<link href="css/pic.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
<script type="text/javascript" src="dhtml.js"></script>

</head>

<body  onload=init();>
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



<div id="link_cpsc">

	<div id="link_spsc_left">

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="148px" class="link_tab">
          <tr>
            <td width="1" rowspan="2"><img src="images/pic_left.jpg" /></td>
            <td>&nbsp;</td>
            <td height="26">建筑</td>
            <td>&nbsp;</td>
            <td width="1" rowspan="2"><img src="images/pic_right.jpg" /></td>
          </tr>
          <tr>
            <td width="33" align="center" valign="middle"><img src="images/pic_left.gif" width="16" height="25" /></td>
            <td height="123">
            
            
            
            
            
            
            
            
              <div id=demo style="OVERFLOW: hidden; WIDTH: 180px; align: center">
  <table cellspacing="0" cellpadding="0" align="center" 
border="0">
    <tbody>
      <tr>
        <td id="marquePic1" valign="top">
	<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from project"
rs.open sql,conn,1,1
if not rs.bof and not rs.eof then
	if rs("filename6")<>"" then
	
	for iii=1 to rs.recordcount
		filename6=filename6&rs("filename6")
		filedown6=filedown6&rs("filedown6")
		filetype6=filetype6&rs("filetype6")
	rs.movenext
	if rs.eof then exit for
	next

	n=split(filename6,",")
	a=split(filedown6,",")
	b=split(filetype6,",")
	for nn=0 to ubound(n)
		if InStr(n(nn),"建筑")=1 then
			tag1=tag1&nn&","
			num1=num1+1
		end if
		if InStr(n(nn),"景观")=1 then
			tag2=tag2&nn&","
			num2=num2+1
		end if
		if InStr(n(nn),"细节")=1 then
			tag3=tag3&nn&","
			num3=num3+1
		end if
	next

	end if
end if
rs.close
set rs=nothing

tag01=split(tag1,",")
tag02=split(tag2,",")
tag03=split(tag3,",")
max=num1-1
half=int(max/2)

%>	
<table width="<%=cint(num1/2)*60%>" height="105" border="0" cellpadding="0" cellspacing="0" style="color:#999999;">

  <tr>
  <%
	for i=lbound(n) to half

%>
    <td  align="center">
<a href="link_cpsc.asp?tag=<%=a(tag01(i))%>&i=<%=tag01(i)%>"><img src="../upload/images/<%=a(tag01(i))&b(tag01(i))%>" width="50" height="40" /></a>
</td><%
	next
	%>
    </tr>
    <tr>
     <%
	for i= half+1 to max

%>
    <td  align="center">
<a href="link_cpsc.asp?tag=<%=a(tag01(i))%>&i=<%=tag01(i)%>"><img src="../upload/images/<%=a(tag01(i))&b(tag01(i))%>" width="50" height="40" /></a>
</td><%
	next
	%>
  </tr>

</table>		
		
		</td>
        <td id="marquePic2" valign="top"></td>
      </tr>
    </tbody>
  </table>
</div>
<script type=text/javascript> 
var speed=30 
marquePic2.innerHTML=marquePic1.innerHTML 
function Marquee(){ 
if(demo.scrollLeft>=marquePic1.scrollWidth){ 
demo.scrollLeft=0 
}else{ 
demo.scrollLeft++ 
}} 
var MyMar=setInterval(Marquee,speed) 
demo.onmouseover=function() {clearInterval(MyMar)} 
demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)} 
</script>
            
            
            
            
            
            
            
            
            
            
            
            
            
            </td>
            <td width="33" align="center" valign="middle"><img src="images/pic_right.gif" width="16" height="25" /></td>
    </tr>
        </table>
		
        
        
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="148px" class="link_tab">
          <tr>
            <td width="1" rowspan="2"><img src="images/pic_left.jpg" /></td>
            <td>&nbsp;</td>
            <td height="26">景观</td>
            <td>&nbsp;</td>
            <td width="1" rowspan="2"><img src="images/pic_right.jpg" /></td>
          </tr>
          <tr>
            <td width="33" align="center" valign="middle"><img src="images/pic_left.gif" width="16" height="25" /></td>
            <td height="123"><div id="demo1" style="OVERFLOW: hidden; WIDTH: 180px; align: center">
                <table cellspacing="0" cellpadding="0" align="center" 
border="0">
                  <tbody>
                    <tr>
                      <td id="marquePic11" valign="top"><%

max=num2-1
half=int(max/2)
%>
                          <table width="<%=cint(num2/2)*60%>" height="105" border="0" cellpadding="0" cellspacing="0" style="color:#999999;">
                            <tr>
                              <%
	for i=lbound(n) to half

%>
                              <td  align="center"><a href="link_cpsc.asp?tag=<%=a(tag02(i))%>&amp;i=<%=tag02(i)%>"><img src="../upload/images/<%=a(tag02(i))&b(tag02(i))%>" width="50" height="40" /></a> </td>
                              <%
	next
	%>
                            </tr>
                            <tr>
                              <%
	for i= half+1 to max

%>
                              <td  align="center"><a href="link_cpsc.asp?tag=<%=a(tag02(i))%>&amp;i=<%=tag02(i)%>"><img src="../upload/images/<%=a(tag02(i))&b(tag02(i))%>" width="50" height="40" /></a> </td>
                              <%
	next
	%>
                            </tr>
                        </table></td>
                      <td id="marquePic21" valign="top"></td>
                    </tr>
                  </tbody>
                </table>
            </div>
                <script type="text/javascript"> 
var speed=30 
marquePic21.innerHTML=marquePic11.innerHTML 
function Marquee1(){ 
if(demo1.scrollLeft>=marquePic11.scrollWidth){ 
demo1.scrollLeft=0 
}else{ 
demo1.scrollLeft++ 
}} 
var MyMar1=setInterval(Marquee1,speed) 
demo1.onmouseover=function() {clearInterval(MyMar1)} 
demo1.onmouseout=function() {MyMar1=setInterval(Marquee1,speed)} 
          </script>
            </td>
            <td width="33" align="center" valign="middle"><img src="images/pic_right.gif" width="16" height="25" /></td>
          </tr>
        </table>
        
        
        
        
        
        
        
        
        
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="148px" class="link_tab">
          <tr>
            <td width="1" rowspan="2"><img src="images/pic_left.jpg" /></td>
            <td>&nbsp;</td>
            <td height="26">细节</td>
            <td>&nbsp;</td>
            <td width="1" rowspan="2"><img src="images/pic_right.jpg" /></td>
          </tr>
          <tr>
            <td width="33" align="center" valign="middle"><img src="images/pic_left.gif" width="16" height="25" /></td>
            <td height="123">
            
            
            
            
            
            <div id="demo3" style="OVERFLOW: hidden; WIDTH: 180px; align: center">
                <table cellspacing="0" cellpadding="0" align="center" 
border="0">
                  <tbody>
                    <tr>
                      <td id="marquePic3" valign="top"><%

max=num3-1

half=int(max/2)


%>
                          <table width="<%=cint(num3/2)*60%>" height="105" border="0" cellpadding="0" cellspacing="0" style="color:#999999;">
                            <tr>
                              <%
	for i=lbound(n) to half

%>
                              <td  align="center"><a href="link_cpsc.asp?tag=<%=a(tag03(i))%>&amp;i=<%=tag03(i)%>"><img src="../upload/images/<%=a(tag03(i))&b(tag03(i))%>" width="50" height="40" /></a> </td>
                              <%
	next
	%>
                            </tr>
                            <tr>
                              <%
	for i= half+1 to cint(max)

%>
                              <td  align="center"><a href="link_cpsc.asp?tag=<%=a(tag03(i))%>&amp;i=<%=tag03(i)%>"><img src="../upload/images/<%=a(tag03(i))&b(tag03(i))%>" width="50" height="40" /></a> </td>
                              <%
	next
	%>
                            </tr>
                        </table></td>
                      <td id="marquePic4" valign="top"></td>
                    </tr>
                  </tbody>
                </table>
            </div>
                <script type="text/javascript"> 
var speed=30 
marquePic4.innerHTML=marquePic3.innerHTML 
function Marquee3(){ 
if(demo3.scrollLeft>=marquePic3.scrollWidth){ 
demo3.scrollLeft=0 
}else{ 
demo3.scrollLeft++ 
}} 
var MyMar3=setInterval(Marquee3,speed) 
demo3.onmouseover=function() {clearInterval(MyMar3)} 
demo3.onmouseout=function() {MyMar3=setInterval(Marquee3,speed)} 
          </script>
            </td>
            <td width="33" align="center" valign="middle"><img src="images/pic_right.gif" width="16" height="25" /></td>
          </tr>
        </table>
  </div>
    
    <div id="link_spsc_right">
    	<img src="images/link_cpsc.jpg" />
    	<div class="pic_bigmain">
        <%
		tag=request("tag")
		if tag="" then
		%>
        <img src="../upload/<%=a(max)&b(max)%>"  />
        <%
		else
		%>
        <img src="../upload/<%=a(request("i"))&b(request("i"))%>"  />
        <%
		end if
		%>
        </div>
    </div>



</div>





<div id="bottom">
	<div id="bottom_content">
    <select>
    	<option selected="selected">----------友情链接----------</option>
    </select>
    <em>
    	<a href="#">联系我们 |</a>
        <a href="#">在线统计 |</a>
        <a href="#">网站地图 |</a>
        <a href="#">法律声明</a>
    </em>
    </div>
</div>
