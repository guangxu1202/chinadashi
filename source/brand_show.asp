<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="��ʵ��������,�����˴�ִ��̤ʵ��������ʵ��ҵ�������޹�˾�ĺ��Ĳ�ҵ�Ƿ��ز�����,�ѿ������ɵ���ɽׯ�����俥�����������⡢���ֺ�ɽ��סլС��������30����ƽ����">
<meta name="keywords" content="��ʵ���ţ���ʵ�����ֺ�ɽ������¥�̣��������ز������ز�" >
<title>��ʵ����-��ƷƷ��</title>
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
        	<li><a href="brand.asp">��ʵ����</a></li>
            <li><a href="brand.asp?act=2">��ʵ֮��</a></li>
            <li><a href="brand.asp?act=3">��ʵ�Ļ�</a></li>
            <li><a href="brand_list.asp">��ƷƷ��</a></li>
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
            <div class="Main_right_guide">&nbsp;&nbsp;-��Ŀ�ſ�</div>
          <ul>
           	<li><img src="images/point.gif" /><strong>��¥���ߣ�</strong><%=rs("slrx")%></li>
                <li><img src="images/point.gif" /><strong>��Ʒ�۸�</strong><%=rs("cpjg")%></li>
                <li><img src="images/point.gif" /><strong>ռ�������</strong><%=rs("zdmj")%></li>
                <li><img src="images/point.gif" /><strong>���������</strong><%=rs("jzmj")%></li>
                <li><img src="images/point.gif" /><strong>�滮��;��</strong><%=rs("ghyt")%></li>
                <li><img src="images/point.gif" /><strong>���������</strong><%=rs("hxmj")%></li>
                <li><img src="images/point.gif" /><strong>��סʱ�䣺</strong><%=rs("rzsj")%></li>
                <li><img src="images/point.gif" /><strong>������·��</strong><%=rs("gjxl")%></li>
                <li><img src="images/point.gif" /><strong>��Ŀ��λ��</strong><%=rs("xmqw")%></li>
                <li><img src="images/point.gif" /><strong>�ܱ����ף�</strong><%=rs("zbpt")%></li>
                <li><img src="images/point.gif" /><strong>��Ŀ���ܣ�</strong><%=rs("xmjs")%></li>
                <li><img src="images/point.gif" /><strong>��Ʒ��Ϣ��</strong><%=rs("cpxx")%></li>
          </ul>
            <div class="Main_right_guide">&nbsp;&nbsp;-ǰ�ض�̬</div>
            <%if rs("qydt")<>"" then%>
            <ul>
            	<li><%=rs("qydt")%></li>
            </ul>
            <%end if%>
            <div class="Main_right_guide">&nbsp;&nbsp;-��ƷͼƬ</div>
            <div class="Main_right_Picguide">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/point.gif" /><strong>��Ŀ��λ</strong></div>
            <script language=JavaScript src="images/scroll.js"></script>
            	<div class="Main_right_xmqw">
                <%if rs("filename5")<>"" then%>
                	<img src="../upload/<%=rs("filedown5")&rs("filetype5")%>" />
                <%end if%>
                </div>
            <div class="Main_right_Picguide">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/point.gif" /><strong>��Ŀʵ��</strong></div>
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
            <div class="Main_right_Picguide">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/point.gif" /><strong>��Ŀ����</strong></div>
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