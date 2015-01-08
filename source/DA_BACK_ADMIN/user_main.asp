<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理-用户管理</title>

<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="top.asp" -->
<table width="1000" height="600" border="0" align="center" cellpadding="0" cellspacing="0" class="bxline">
  <tr>
    <td>&nbsp;</td>
    <td valign="top">&nbsp;</td>
    <td>&nbsp;</td>
    <td valign="top" class="right">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td valign="top">&nbsp;</td>
    <td>&nbsp;</td>
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">管理员列表</span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="50">&nbsp;</td>
    <td width="164" valign="top"><!--#include file="left_default.asp" --></td>
    <td width="25" class="leftline">&nbsp;</td>
    <td valign="top" class="right"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="d_right">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
      <tr>
        <td height="500" valign="top">
		
		 <br>
		 <br><br><br>
		<form>
          <table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="44" align="center" bgcolor="#999999">序号</td>
              <td width="153" height="30" align="center" bgcolor="#999999">用户姓名</td>
              <td width="166" align="center" bgcolor="#999999">用户帐户</td>
              <td width="97" align="center" bgcolor="#999999">状态</td>
              <td width="186" align="center" bgcolor="#999999" class="addnew"><a href="user_add.asp">添加用户</a></td>
            </tr>
            <tr>
              <td width="44" align="center" valign="top" background="../images/line.gif"></td>
              <td width="153" align="center" valign="top" background="../images/line.gif"></td>
              <td width="166" align="center" valign="top" background="../images/line.gif"></td>
              <td width="97" align="center" valign="top" background="../images/line.gif"></td>
              <td width="186" align="center" valign="top" background="../images/line.gif"></td>
            </tr>
          </table>
          <table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
            <% 
		      set rs=server.CreateObject("adodb.recordset")
			    sql="select * from users where tag=0"
				i=0
				rs.open sql,conn,1,1
				  while not rs.eof 
				  i=i+1
		  %>
            <tr>
              <td width="45" bgcolor="#F5F5F5" <%if i/2=1 then%>class="line"<%end if%>><div align="center"><%=i%></div></td>
              <td width="156" align="center" bgcolor="#F5F5F5" <%if i/2=1 then%>class="line"<%end if%>>&nbsp;<%=rs("uname")%></td>
              <td width="168" height="20" align="center" bgcolor="#F5F5F5" <%if i/2=1 then%>class="line"<%end if%>>&nbsp;<%=rs("uid")%></td>
              <td width="100" align="center"  bgcolor="#F5F5F5" <%if i/2=1 then%>class="line"<%end if%>><%if rs("zx")=0 then response.Write("<span style=' color:#009900'>正常</span>") else response.Write("<span style=' color:#999999'>已注销</span>")%></td>
              <td width="100" align="center" bgcolor="#F5F5F5" <%if i/2=1 then%>class="line"<%end if%>><a href="user_edit.asp?id=<%=rs("id")%>">编辑用户</a></td>
              <td width="85" align="center" bgcolor="#F5F5F5" <%if i/2=1 then%>class="line"<%end if%>><a href="#" onclick="javascript:if (confirm('确定要删除吗？')){window.location.href='user_sql.asp?id='+<%=rs("id")%>+'&act=dele'}">删除用户</a></td>
            </tr>
            <% 
	     rs.movenext
	     wend
	 
	  rs.close
	  set rs=nothing
	  %>
          </table>
        </form>          <p>
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
