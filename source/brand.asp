<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="��ʵ��������,�����˴�ִ��̤ʵ��������ʵ��ҵ�������޹�˾�ĺ��Ĳ�ҵ�Ƿ��ز�����,�ѿ������ɵ���ɽׯ�����俥�����������⡢���ֺ�ɽ��סլС��������30����ƽ����">
<meta name="keywords" content="��ʵ���ţ���ʵ�����ֺ�ɽ������¥�̣��������ز������ز�" >
<title>��ʵ����-��ʵƷ��</title>
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
  <%if request("act")=3 then%>
  <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','992','height','312','style','z-index:0;','src','images/Inbanner7','wmode','transparent','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','images/Inbanner7' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="992" height="312" style="z-index:0;">
    <param name="movie" value="images/Inbanner7.swf" />
    <param name="quality" value="high" />
	<param name="wmode" value="transparent">
    <embed src="images/Inbanner7.swf" wmode="transparent" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="992" height="312"></embed>
  </object></noscript></noscript>
  
  
	<%else%>
<script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','992','height','312','style','z-index:0;','src','images/Inbanner4','wmode','transparent','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','images/Inbanner4' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="992" height="312" style="z-index:0;">
    <param name="movie" value="images/Inbanner4.swf" />
    <param name="quality" value="high" />
	<param name="wmode" value="transparent">
    <embed src="images/Inbanner4.swf" wmode="transparent" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="992" height="312"></embed>
  </object></noscript></noscript>
  
  <%end if%>
  

  
  
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
    	<em><img src="images/dspp.jpg" /></em>
        <ul>
        	<li><a href="brand.asp">��ʵ����</a></li>
            <li><a href="brand.asp?act=2">��ʵ֮��</a></li>
            <li><a href="brand.asp?act=3">��ʵ�Ļ�</a></li>
            <li><a href="brand_list.asp">��ƷƷ��</a></li>
        </ul>
    </div>
    <div id="Main_right">
    <%if request("act")="" then%>
   		 <img  src="images/Dsln.jpg" class="guide" />
    <%elseif request("act")=2 then%>
    	<img  src="images/brand2.jpg" class="guide" />
    <%elseif request("act")=3 then%>
    	<img  src="images/brand3.jpg" class="guide" />
    <%end if%>
     	
        <div class="Main_right_content">
        
        <%if request("act")="" then%>
                <h4 style="font-size:16px"><strong>�����˴�  ִ��̤ʵ</strong></h4>
<br />
<p>���壺</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ɽ���Ǳ����������ƶˣ�������ɰٴ����������ġ��Ⱥ��������գ����������˵��Ļ�����Ҫ��˼���������Ƶļ�����������չ˫�ᣬ����ɶ�߾��ܷɶ�ߡ���Ҫ����������խ�յ�С�ݣ����������ܵ�װ���������̣������ϱ��С����ݵö࣬����ǿ�����ɵö࣬���ܴ�������޵Ŀռ䲻�ǿտյ��������ǳ������ں���</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��չֻ��վ��û���յ㡣��չ�ĵ�·������������ƽ̹���������Ϣ������ֹ������������������ܷ���������ִ�����������ֵ�����еĴʻ㡣</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ϵ�԰�ز���֦��Ҷï����ϵ�֦農���˶�����ۣ����죬����������ϵ����ӣ����죬�����ջ���ϵĹ�ʵ����ϵĹ�ʵ�����ǹ�ͬ����</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ӥ��׽��Զ��Ŀ���Ҫ��ɣ���Ҫ����һ���켣����Ҫ��Ϊ��û���������㼣����˵����õ�������ס��������ýŲ������ɹ����ۻ����������ó��ȥ���ܷܶ�����ζ�����Ƕ�Ҫһ��һ���ߣ�һ��һ��ɣ�</p>

    <%elseif request("act")=2 then%>
    <br />
<strong>ƾ������ ����ʵҵ</strong><br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������ǵĸ������ǵĸ��Ǵ����������������淢չ�ĸ������ڡ����ǵ������Ǵ������裬���ǵ����ԴӴ�����ȡ�������������������ױȵĵ�λ��ʹ����ӵ�����ҷ�չ��Ϊ�������ٳ��е��������������ĵ�������������ׯ�ϵؽ�ȡ����ΰ�����ϡ�<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ʵ���Ǵ�ʵ�˵�����֮����Ǧ��ϴ����������ɬ����ѿ�ѽ�ɽ������Ĺ�ʵ�����궯��ƽʵ����ӳ�����ǵĳɳ������ǵĳ��졣̤̤ʵʵ����ʵ�޻���������ǣ�ķ��ݣ������޸��ĸ�Ƽ���������ɭ�֣��Ѹ�����������أ��ڴ������ΰ����<br />
<br />
<strong>������� ����̤ʵ</strong><br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ʵ���ŵķ�չת���߹�35�ꡣ��Щ���ʵ�����߹���һ����ѧʽ�ķ�չ�켣�����޵��У���ǿԽǿ��һ���˵ĳ�����˼��ĳ��죬һ����ҵ�ĳ���������ҵ�Ļ�������ĳ��졣���������ĳ������Ǵ�ʵ��ҵ�Ļ�����Դ��<br />

<br />
<strong>�����Ϊ��ҵ����֮�� ��ʵ��Ϊ��˾�������</strong><br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���ڴ�ʵ������������ҵ�ɳ�����û��ʲô���������֧�ָ���Ҫ���ˡ�������ˡ�ʵ�������Ǵ�ʵ���Ŷ�ÿһ��Ա���������Ҫ��Ҳ����ҵ�ܰ�����������չ׳��ĸ�����ʱ����䣬������䣬����ʵ�˵�����ԭ�������ı䣡<br />
<br />
<strong>���� ���� ����</strong><br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ÿһ����ʵ�˶�����ϵģ�ÿһ����ʵ�˶��������ġ���ˮ���㣬�����ԣ��Ӵ�ʵ��һשһ�ߡ���ʵ�˵�һ��һ�У���ᷢ�֣���������׷����Զ�Ǵ�ʵ��ҵ���ڵ������Ŀ�ġ�<br />

<br />
<br />


    <%elseif request("act")=3 then%>
    <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;һȺ��ʵ��ũ�񣬴���������ʮ��������﷫������Ҳ��������֮����û���뵽��ҵ�ᷢչ���������ģ��������ֻ��ʹ�Լ��������һЩ��������ǰ�������ࡣȻ���������������ע����������ʩչ��ƽ̨����ʵ��׷����ʱ���Ĵ���ײ����ȫ�ı�������ǰ���ķ��������ÿ���������Ĺ켣���ĸ￪�ŵĴ󳱣����羭�÷�չ�Ĵ󳱣��������ܾ����ܾ����˵ĺ����Ĺ�Ю�£����Ǿ����Ŵ�����ɳ���ɹ��Ĵ�ϲ��ʧ���Ĵ󱯣��ջ�ĸ��𣬴��۵Ŀ�ɬ������֧�Ų�ס����������������ȴ��������������ѣ���͹��崣��ѽ��п�����ҵ���ˣ���ҵ��ң���ҵ�������ʮ��ļ��飬��ʮ��ĺ�������ʮ����ջ񡭡�</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ʵ���ŵķ�չת���߹�35�ꡣ����������������վ�ڴ����ĺ��ߣ������Ŵ�֮�����ǣ�ֻ���﷫�߽��󺣵��ػ���������ͨ�����Ѫ����Ѱ�������׷���һ��ʼҲ�������ɵģ�����������һ�ô�������������ŷ���һ�����硣���ܵȴ������з��꣬���׵磬�о��ˣ���������һ���﷫����һ����ǰ�����������ã����һ�����࣬��ļ���ƣ����������ı������׽�Ʒ����ĥ���д����ζ��ʲô�Ǵ������ʿռ����չ������˼������ķ��ӡ���ֻ����Ϧ�ذνڣ�����ѭ�򽥽���������</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ǰ�еĽ�ӡ����Ѱ��ʷ�Ĵ𰸣��������Ĺ��ܴ����й��ɳ���������ԡ�</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������ǵĸ������ǵĸ��Ǵ����������������淢չ�ĸ������ڡ����ǵ������Ǵ������裬���ǵ����ԴӴ�����ȡ�������������������ױȵĵ�λ��ʹ����ӵ�����ҷ�չ��Ϊ�������ٳ��е��������������ĵ�������������ׯ�ϵؽ�ȡ����ΰ�����ϡ�</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ʵ�������ǵ�Ϊ������֮����Ǧ��ϴ����������ɬ����ѿ�ѽ�ɽ������Ĺ�ʵ�����궯��ƽʵ����ӳ�����ǵĳɳ������ǵĳ��졣̤̤ʵʵ����ʵ�޻���������ǣ�ķ��ݣ������޸��ĸ�Ƽ���������ɭ�֣��Ѹ�����������أ��ڴ������ΰ����</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ƾ������������ʵҵ��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����Ǵ󺣵�������ǳ����Ӵ�����õ���ʾ��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����ѧϰ�����ɰٴ��Ĳ����ػ����㽻���ѣ�������ʿ��Я��ǰ�С�</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����ѧϰ�󺣻�ˮ��Ԩ��̤ʵ���񣬲�ͼ��������ͷ��ɣ��ɾ�ΰҵ��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�Ӹ��ʵĽ��ּ����ܶ������ǲ�и��׷��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�Ӵ��õĵ�������ܿ������Ǽ�ʵ�Ļ���</p>

    <%end if%>
      

        </div>
    </div>
</div>

















<!--#include file="bottom.asp" -->

</body>
</html>