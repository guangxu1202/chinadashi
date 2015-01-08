<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理-用户管理</title>
<script type="text/javascript" src="../images/nav.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function doLoginNow(loginform){

   if(loginform.oldmm.value==''){
      alert("旧密码不能为空!");
	  loginform.oldmm.focus();
	  return false;
   }
    if(loginform.newmm.value==''){
      alert("密码不能为空!");
	  loginform.newmm.focus();
	  return false;
   }
   if(loginform.newcmm.value==''){
      alert("密码不能为空!");
	  loginform.newcmm.focus();
	  return false;
   }
  loginform.action="pwd_sql.asp";
}

//-->
</script>
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="top.asp" -->
<table width="1000" height="600" border="0" align="center" cellpadding="0" cellspacing="0" class="bxline">
  <tr>
    <td>&nbsp;</td>
    <td valign="top" class="leftnav">&nbsp;</td>
    <td>&nbsp;</td>
    <td valign="top" class="right">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td valign="top" class="leftnav">&nbsp;</td>
    <td>&nbsp;</td>
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">密码修改</span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="50">&nbsp;</td>
    <td width="164" valign="top" class="leftnav"><!--#include file="left_default.asp" --></td>
    <td width="25" class="leftline">&nbsp;</td>
    <td valign="top" class="right"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="d_right">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
      <tr>
        <td height="500" valign="top"><br><br><br>
		 <form action="" method="post" name="form1" id="form1" onsubmit="return doLoginNow(this);">
           <table width="437" border="0" align="center" cellpadding="0" cellspacing="0">
             <tr>
               <td width="87" align="right">请输入旧密码：</td>
               <td width="350" height="30"><input name="oldmm" type="password" id="oldmm" onblur="showHint(this.value)">
                   <span id="txtHint"></span></td>
             </tr>
             <tr>
               <td align="right">请输入新密码：</td>
               <td height="30"><input name="newmm" type="password" id="newmm" />               </td>
             </tr>
             <tr>
               <td align="right">请重复新密码：</td>
               <td height="30"><input name="newcmm" type="password" id="newcmm" /></td>
             </tr>
             <tr>
               <td align="right">&nbsp;</td>
               <td height="30"><input type="submit" name="Submit" value="修改" />
                   <input name="Submit2" type="button" onclick="MM_goToURL('parent','default.asp');return document.MM_returnValue" value="返回" /></td>
             </tr>
           </table>
		   </form>
		 <p>
		   <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0','border','0','width','307','height','238','style','float: right; display:block; top:0; position: relative; z-index: 1; bottom:0px; left: 0;','src','images/po','pluginspage','http://www.macromedia.com/go/getflashplayer','quality','High','wmode','transparent','movie','images/po' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11CF-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0" border="0" width="307" height="238" style="float: right; display:block; top:0; position: relative; z-index: 1; bottom:0px; left: 0;">
             <param name="movie" value="images/po.swf" />
             <param name="quality" value="High" />
             <param name="wmode" value="transparent" />
             <embed src="images/po.swf" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="307" height="238" quality="High" wmode="transparent"> </embed>
           </object></noscript>
		 </p>          </td>
      </tr>
    </table></td>
    <td width="6">&nbsp;</td>
  </tr>
    <tr>
    <td>&nbsp;</td>
    <td valign="top" class="leftnav">&nbsp;</td>
    <td>&nbsp;</td>
    <td valign="top" class="right">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="BFBFBF">
  <tr>
    <td height="1"></td>
  </tr>
</table>
</body>
</html>
