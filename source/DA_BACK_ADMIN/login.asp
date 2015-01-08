<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台登录</title>
<%
ex=request.QueryString("ex")
if ex=1 then
	session("love_id")=""
end if
%>
<style type="text/css">
<!--
body {
	margin:0px;
	padding:0px;
}
input.smallInput{background: #FFFFFF;border-bottom-color:#CCCCCC; border-bottom-width:1px;border-top-width:0px;border-left-width:0px;border-right-width:0px; solid #ff6633; color: #000000; FONT-SIZE: 9pt; color: #000000; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; width:200px; LINE-HEIGHT: normal}
input.buttonface{BACKGROUND: #1D1E19; border:1 solid #ff6633; COLOR: #CCCCCC; FONT-SIZE: 9pt; HEIGHT:20px; width:50px; LINE-HEIGHT: normal}
.STYLE4 {color: #666666}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--

function submitOk(){
	var uid=document.getElementById("uid").value;
	var pwd=document.getElementById("pwd").value;
	
	if(uid==''){
      alert("帐号不能为空!");
	  return false;
	}
	if(pwd==''){
      alert("密码不能为空!");
	  return false;
   }
     document.form1.action="../DA_CHMRW/pass.asp";
	 document.form1.submit();
}

//-->
</script>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body onkeydown="if(event.keyCode==13)submitOk()">
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td height="120"><img src="images/logo.jpg"  /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center" >&nbsp;</td>
    <td align="center" style="border-bottom:1px solid #CCCCCC;"> <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0','width','270','height','270','src','images/shu','quality','high','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','images/shu' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="270" height="270">
      <param name="movie" value="images/shu.swf" />
      <param name="quality" value="high" />
      <embed src="images/shu.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="270" height="270"></embed>
    </object></noscript>      &nbsp;&nbsp;</td>
    <td align="center" >&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="100"><p>&nbsp;</p>
      <form id="form1" name="form1" method="post" action="" onSubmit="return doLoginNow(this);">
      
      <table width="320" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>&nbsp;</td>
        <td height="30">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="30" align="center" valign="middle"><span class="STYLE4">用户名:</span></td>
        <td height="30" align="center" valign="middle"><input name="uid" type="text" class="smallInput" id="uid" /></td>
        <td align="center" valign="middle">&nbsp;</td>
      </tr>
      <tr>
        <td height="30" align="center" valign="middle" class="STYLE4">密&nbsp;&nbsp;码:</td>
        <td height="30" align="center" valign="middle"><input name="pwd" type="password" class="smallInput" id="pwd" /></td>
        <td align="center" valign="middle"><img src="images/Lmenu.jpg" width="20" height="17" onclick="submitOk()" id="menu" style="cursor:pointer;" /></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td height="10">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="30" colspan="3" align="center">&nbsp;</td>
      </tr>
    </table>
	
	<script>
	var form=document.getElementById('menu'); 

form.onkeydown=function(evt){ 
    evt=evt?evt:window.evt; 
    if(13==evt.keyCode){//监视是否按下'Enter'键 
        fsubmit(form); 
    } 
}	
	</script>
      </form></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
