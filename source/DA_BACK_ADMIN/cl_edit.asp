<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<%id=request.QueryString("id")
%>
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理-用户管理</title>


<script src="../scripts/AC_RunActiveContent.js" type="text/javascript"></script>
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
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">详细登记信息</span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="50">&nbsp;</td>
    <td width="164" valign="top" class="leftnav"><!--#include file="left_client.asp" --></td>
    <td width="25" class="leftline">&nbsp;</td>
    <td valign="top" class="right"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="d_right">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
      <tr>
        <td height="500" valign="top"><p>&nbsp;</p>

         <%
		 set rs=server.CreateObject("adodb.recordset")
		 sql="select * from khly where id="&id
		 rs.open sql,conn,1,1
		 if not rs.bof and not rs.eof then

	

		 %>
          <table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#F0F0F0" bgcolor="#F9F9F9" id="main">
          <tr>
            <td height="28" align="center" bgcolor="#999999" style="color:#FFFFFF; font-weight:bold;">详细信息：</td>
            <td height="28" bgcolor="#999999">&nbsp;</td>
          </tr>
          <tr>
            <td width="13%" height="28" align="center">姓名：</td>
            <td width="87%" height="28">&nbsp;<%=rs("Uname")%></td>
          </tr>
          <tr>
            <td height="28" align="center">性别：</td>
            <td height="28">&nbsp;<%=rs("Uxb")%></td>
          </tr>
          <tr>

            <td height="28" align="center">联系电话：</td>
            <td height="28">&nbsp;<%=rs("Utel")%></td>
          </tr>
          <tr>
            <td height="28" align="center">地址：</td>
            <td height="28">&nbsp;<%=rs("Udz")%></td>
          </tr>
          <tr>
            <td height="28" align="center">邮编：</td>
            <td height="28">&nbsp;<%=rs("Uyb")%></td>
          </tr>
          <tr>
            <td height="28" align="center">电子邮箱：</td>
            <td height="28">&nbsp;<%=rs("Umail")%></td>
          </tr>
          <tr>
            <td height="28" align="center">计划购买业态：</td>
            <td height="28">&nbsp;
			<%
			if rs("Place")=0 then
				response.Write("未选择")
			elseif rs("Place")=1 then
				response.Write("大连-"&rs("zone"))
			elseif rs("Place")=2 then
				response.Write("沈阳-"&rs("zone"))
			elseif rs("Place")=3 then
				response.Write("四川-"&rs("zone"))
			end if
			%>
			
			</td>
          </tr>
          <tr>
            <td height="28" align="center">计划购买时间：</td>
            <td height="28">&nbsp;<%=rs("Utime")%></td>
          </tr>
          <tr>
            <td height="28" align="center">大实业主：</td>
            <td height="28">&nbsp;<%=rs("Udsyz")%></td>
          </tr>
          <tr>
            <td height="28" align="center">购房经历：</td>
            <td height="28">&nbsp;<%=rs("Ugfjl")%></td>
          </tr>
          <tr>
            <td height="28" align="center">登记时间：</td>
            <td height="28">&nbsp;<%=rs("sendtime")%></td>
          </tr>
          <tr>
            <td height="28" align="center">备注：</td>
            <td height="28">&nbsp;<%=rs("Ubz")%></td>
          </tr>
          <%
		  end if
		  rs.close
		  set rs=nothing
		  %>
          <tr>
            <td height="80" colspan="2"><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="submit" name="Submit3" value=" 返回 " onclick="javascript:history.back()" />
              <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0','border','0','width','307','height','238','style','float: right; display:block; top:0; z-index: 1; bottom:0px; left: 0;','src','images/po','pluginspage','http://www.macromedia.com/go/getflashplayer','quality','High','wmode','transparent','movie','images/po' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11CF-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0" border="0" width="307" height="238" style="float: right; display:block; top:0; z-index: 1; bottom:0px; left: 0;">
                <param name="movie" value="images/po.swf" />
                <param name="quality" value="High" />
                <param name="wmode" value="transparent" />
                <embed src="images/po.swf" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="307" height="238" quality="High" wmode="transparent"> </embed>
              </object></noscript></td>
            </tr>
        </table>   
          <p>&nbsp;</p>          </td>
      </tr>
    </table></td>
    <td width="6">&nbsp;</td>
  </tr>
        <tr>
    <td valign="top">&nbsp;</td>
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
