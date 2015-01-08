<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理-添加管理员</title>
<script type="text/javascript" src="../images/nav.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function SubmitOk()
{

 var uid =document.form1.uid.value;
 var upwd =document.form1.upwd.value;
 if (uid=='')
    {
	alert("用户名不能为空！");
	return false;
	}
 if (upwd=='')
    {
	alert("密码不能为空！");
	return false;
	}
 if (confirm("确定要添加吗？"))
     document.form1.submit();
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">管理员列表</span></td>
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
        <td height="500" valign="top">
		
		 <br>
		 <br><br><br>
		 <form action="user_sql.asp" method="post" name="form1" id="form1">
           <table width="541" border="0" align="center" cellpadding="0" cellspacing="0">
             <tr>
               <td height="30" bgcolor="#D8D8D8" class="STYLE7" >基本信息</td>
               <td height="30" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE">&nbsp;</td>
             </tr>
             <tr>
               <td width="85" align="right" bgcolor="#EEEEEE">&nbsp;</td>
               <td width="85" height="25" align="left" bgcolor="#EEEEEE">用户名：</td>
               <td width="144" height="25" align="left" bgcolor="#EEEEEE"><input name="uid" type="text" class="input" id="uid" size="10" /></td>
               <td width="96" height="25" align="left" bgcolor="#EEEEEE">密码：</td>
               <td width="131" height="25" align="left" bgcolor="#EEEEEE"><input name="upwd" type="password" class="input" id="upwd" size="10" /></td>
             </tr>
             <tr>
               <td align="right" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" align="left" bgcolor="#EEEEEE">姓名：</td>
               <td height="25" align="left" bgcolor="#EEEEEE"><input name="uname" type="text" class="input" id="uname" size="10" /></td>
               <td height="25" align="left" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" align="left" bgcolor="#EEEEEE">&nbsp;</td>
             </tr>
             <tr>
               <td height="20" bgcolor="#EEEEEE" class="STYLE7">&nbsp;</td>
               <td height="21" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="21" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="21" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="21" bgcolor="#EEEEEE" >&nbsp;</td>
             </tr>
             <tr>
               <td height="20" class="STYLE7">&nbsp;</td>
               <td height="21" >&nbsp;</td>
               <td height="21" >&nbsp;</td>
               <td height="21" >&nbsp;</td>
               <td height="21" >&nbsp;</td>
             </tr>
             <tr>
               <td height="30" bgcolor="#D8D8D8" class="STYLE7">权限设置</td>
               <td height="30" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE" >&nbsp;</td>
             </tr>
            
             <tr>
               <td bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" align="left" bgcolor="#EEEEEE">普通管理员：</td>
               <td height="25" align="left" bgcolor="#EEEEEE"><input name="ulevel" type="radio" value="1" checked="checked" /></td>
               <td height="25" bgcolor="#EEEEEE">超级管理员：</td>
               <td height="25" bgcolor="#EEEEEE"><input type="radio" name="ulevel" value="9" /></td>
             </tr>
             <tr>
               <td bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" align="right" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" align="right" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" bgcolor="#EEEEEE">&nbsp;</td>
             </tr>
             <tr>
               <td>&nbsp;</td>
               <td height="25" align="right">&nbsp;</td>
               <td height="25" align="right">&nbsp;</td>
               <td height="25">&nbsp;</td>
               <td height="25">&nbsp;</td>
             </tr>
             <tr>
               <td>&nbsp;</td>
               <td height="25" align="right"><input name="add" type="button" class="button" id="add" value="添加" onclick="SubmitOk()" /></td>
               <td height="25" align="right"><input name="Submit" type="button" onclick="MM_goToURL('parent','user_main.asp');return document.MM_returnValue" value="返回" /></td>
               <td height="25">&nbsp;</td>
               <td height="25"><input name="act" type="hidden" id="act" value="add" /></td>
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
