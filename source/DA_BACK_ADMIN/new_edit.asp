<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<%
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sql="select * from new where id="&id
rs.open sql,conn,1,1
	title=rs("title")
	content=rs("content")
	tag=rs("tag")
	id=rs("id")
rs.close
set rs=nothing
%>
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理-用户管理</title>

<script type="text/javascript" src="../images/nav.js"></script>
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
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">新闻动态修改</span></td>
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
        <td height="500" valign="top"><p>&nbsp;</p>
		  <form id="form1" name="form1" method="post" action="new_sql.asp">
        <table width="700" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#F0F0F0">
          
          <tr bgcolor="#EEFEE0" height="18">
            <td width="92" height="25" align="center" bgcolor="#F9F9F9">标题：</td>
            <td width="502" height="30" bgcolor="#F9F9F9">
                  <input name="title" type="text" id="title" value="<%=title%>" size="30" />              </td>
          </tr>
         
          <tr bgcolor="#EEFEE0" height="18">
            <td height="25" align="center" bgcolor="#F9F9F9">分类：</td>
            <td><select name="tag" id="tag">
              <option value="1" <%if tag=1 then response.Write("selected")%>>公司新闻</option>
              <option value="2" <%if tag=2 then response.Write("selected")%>>产业新闻</option>
              <option value="3" <%if tag=3 then response.Write("selected")%>>媒体报道</option>
            </select></td>
          </tr>
          <tr bgcolor="#EEFEE0" height="18">
            <td height="25" align="center" bgcolor="#F9F9F9">内容：</td>
            <td><input name="content" type="hidden" id="content" value="<%Response.Write Server.HTMLEncode(content)%>" />
                <iframe id="eWebEditor1" src="../ewebeditor/ewebeditor.asp?id=content&amp;style=s_coolblue" frameborder="1" scrolling="No" width="600" height="380">&nbsp;</iframe></td>
          </tr>
		   <tr bgcolor="#EEFEE0" height="18">
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9"><input type="submit" name="Submit" value=" 修改 " />
              <input type="reset" name="Submit2" value=" 重置 " />
              <input type="button" name="Submit3" value=" 返回 " onclick="javascript:history.back()" />
              <input name="act" type="hidden" id="act" value="edit" />
              <input name="id" type="hidden" id="id" value="<%=id%>" /></td>
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
    <td valign="top">&nbsp;</td>
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
