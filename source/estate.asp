<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="��ʵ��������,�����˴�ִ��̤ʵ��������ʵ��ҵ�������޹�˾�ĺ��Ĳ�ҵ�Ƿ��ز�����,�ѿ������ɵ���ɽׯ�����俥�����������⡢���ֺ�ɽ��סլС��������30����ƽ����">
<meta name="keywords" content="��ʵ���ţ���ʵ�����ֺ�ɽ������¥�̣��������ز������ز�" >
<title>��ʵ����-��ʵ��ҵ</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
</head>

<body>

<!--#include file="top.asp" -->



<div id="Main">
	<div id="Main_left">
    	<em><img src="images/estate.jpg" /></em>
        <ul>
        	<li><a href="estate.asp">���ز�����</a></li>
            <li><a href="estate.asp?act=1">������ҵ</a></li>
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
      <p> <strong>���Ĳ�ҵ</strong></p>
      <p>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ʵ��ҵ�������޹�˾�ĺ��Ĳ�ҵ�Ƿ��ز�������������Ͻ������ʵ��ҵ���ŷ��ݿ������޹�˾�Ƿ��ز������ĺ�����ҵ�������9�꣬�ѿ������ɵ���ɽׯ�����俥�����������⡢���ֺ�ɽ��סլС��������30����ƽ���ס��ڽ�����ӳɽסլС��10��ƽ���ף��滮���辰ɽ�������硤����18��ƽ���ס��Ⲻ�����ֱ����Ĵ���ɽ�������������أ����ƿ���30��ƽ�ס������ʵ���ز�����һ��ɳ����Ǵ�����ʵ��ҵ�������޹�˾���������Ŵ�ʵ���ŷ��ݿ������޹�˾����������סլС������ҵ����</p><br />
<p> <strong>��ҵ��Ʒ</strong></p>
<p> 
1.	����ɽׯС��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������8��ƽ�ף�2004�꽨���깤��<br />
2.	���俥��С��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������9��ƽ�ף�2006�꽨���깤��<br />
3.	��������С��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������6��ƽ�ף�2007�꽨���깤��<br />
4.	���ֺ�ɽС��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������7��ƽ�ף�2008�꽨���깤��<br />
5.	����ӳɽС��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������10��ƽ�ף�2009�꿪���ڽ���<br />
6.	��ʯ���Ǻ����ϡ��������硤��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������18��ƽ�ף��ﻮ�С�<br />
7���Ⲻ����<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��1���Ĵ���ɽ�����Ż�԰��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���������10��ƽ�ף�2001���깤��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��2���������󻪡�ˮ����ۡ��<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���������22��ƽ�ף�2007���깤��<br />
</p>
<%elseif request("act")=1 then%>

<p>������ʵ������ó���޹�˾</p>
<p>������λ��簲װ���޹�˾</p>
<p>��ʵ�¹��й��������޹�˾</p>
<p>������ʵ��ҵ�������޹�˾</p>

<%end if%>
        </div>
    </div>
</div>


















<!--#include file="bottom.asp" -->

</body>
</html>