<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理-资料修改</title>

<script language="JavaScript" type="text/JavaScript">
<!--

function SubmitOk()
{
 var uname =document.form1.uname.value;
 if (uname=='')
    {
	alert("昵称不能为空！");
	return false;
	}
 if (confirm("确定要修改吗？"))
     document.form1.submit();
}
//-->
</script>
<script type="text/javascript" src="../images/nav.js"></script>
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body>
<%
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
   sql="select * from users where id="&id
   rs.open sql,conn,1,1
   if not rs.bof and not rs.eof then
      uid =rs("uid")
	  upwd =rs("upwd")
	  uname =rs("uname")
	  zx =rs("zx")
	  ulevel =rs("ulevel")
   end if
   rs.close
   set rs=nothing
function checked(str)
 if str="1" then checked="checked"
end function
%>
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
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">资料修改</span></td>
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
               <td height="30" bgcolor="#D8D8D8" class="STYLE7">基本信息</td>
               <td height="30" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE" >&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="30" bgcolor="#EEEEEE">&nbsp;</td>
             </tr>
             <tr>
               <td width="85" align="right" bgcolor="#EEEEEE">&nbsp;</td>
               <td width="85" height="25" align="left" bgcolor="#EEEEEE">用户名：</td>
               <td width="144" height="25" align="left" bgcolor="#EEEEEE"><%=uid%></td>
               <td width="96" height="25" align="left" bgcolor="#EEEEEE">&nbsp;</td>
               <td width="131" height="25" align="left" bgcolor="#EEEEEE">&nbsp;</td>
             </tr>
             <tr>
               <td align="right" bgcolor="#EEEEEE">&nbsp;</td>
               <td height="25" align="left" bgcolor="#EEEEEE">姓名：</td>
               <td height="25" align="left" bgcolor="#EEEEEE"><input name="uname" type="text" class="input" id="uname" value="<%=uname%>" size="10" /></td>
               <td height="25" align="left" bgcolor="#EEEEEE"><%if zx=1 then%>解除注销：<%else%>注销：<%end if%></td>
               <td height="25" align="left" bgcolor="#EEEEEE"><%if zx=1 then%><input name="zx" type="checkbox" id="zx" value="0" ><%else%><input name="zx" type="checkbox" id="zx" value="1"><%end if%></td>
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
               <td height="25" align="left" bgcolor="#EEEEEE"><input type="radio" name="ulevel" value="1" <%=checked(ulevel)%> /></td>
               <td height="25" bgcolor="#EEEEEE">超级管理员：</td>
               <td height="25" bgcolor="#EEEEEE"><input type="radio" name="ulevel" value="9" <%if ulevel=9 then response.Write("checked")%> /></td>
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
               <td height="25" align="right"><input name="add2" type="button" class="button" id="add2" value="修改" onclick="SubmitOk()" /></td>
               <td height="25" align="right"><input name="Submit22" type="button" class="button"  onclick="history.back()" value="返回" /></td>
               <td height="25"><input name="id" type="hidden" id="id" value="<%=id%>" /></td>
               <td height="25"><input name="act" type="hidden" id="act" value="edit" /></td>
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
