
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理-用户管理</title>
<style type="text/css" media="all">
body table tr td{font:normal normal normal 12px/1.5em Simsun,Arial, "Arial Unicode MS", Mingliu, Helvetica;text-align: inherit;height:100%;word-break : break-all;}
</style>

<script type="text/javascript" src="../images/nav.js"></script>

  <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="top.asp" -->
<%
id=request("id")
set rs=server.CreateObject("adodb.recordset")
sql="select * from project where id="&id
rs.open sql,conn,1,1
if not rs.bof and not rs.eof then
%>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" class="bxline">
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
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">项目添加</span></td>
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
          <form action="success_sql.asp" method="post" enctype="multipart/form-data" name="form1" id="form1">
          
          <table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#F0F0F0">
          
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9">&nbsp;</td>
          </tr>
          <tr>
            <td width="11%" height="25" align="center" bgcolor="#F9F9F9">项目名称：</td>
            <td width="89%" height="25" bgcolor="#F9F9F9"><input name="xmmc" type="text" id="xmmc" value="<%=rs("xmmc")%>">
              *</td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">售楼热线：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="slrx" type="text" id="slrx" value="<%=rs("slrx")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">产品价格：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="cpjg" type="text" id="cpjg" value="<%=rs("cpjg")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">占地面积：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="zdmj" type="text" id="zdmj" value="<%=rs("zdmj")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">建筑面积：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="jzmj" type="text" id="jzmj" value="<%=rs("jzmj")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">规划用途：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="ghyt" type="text" id="ghyt" value="<%=rs("ghyt")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">户型面积：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="hxmj" type="text" id="hxmj" value="<%=rs("hxmj")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">入住时间：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="rzsj" type="text" id="rzsj" value="<%=rs("rzsj")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">公交线路：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="gjxl" type="text" id="gjxl" value="<%=rs("gjxl")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">项目区位：</td>
            <td height="25" bgcolor="#F9F9F9"><input name="xmqw" type="text" id="xmqw" value="<%=rs("xmqw")%>"></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">周边配套：</td>
            <td height="25" bgcolor="#F9F9F9"><textarea name="zbpt" cols="50" rows="5" id="zbpt"><%=rs("zbpt")%></textarea></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">项目介绍：</td>
            <td height="25" bgcolor="#F9F9F9"><textarea name="xmjs" cols="50" rows="5" id="xmjs"><%=rs("xmjs")%></textarea></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">产品信息：</td>
            <td height="25" bgcolor="#F9F9F9"><textarea name="cpxx" cols="50" rows="5" id="cpxx"><%=rs("cpxx")%></textarea></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">前沿动态：</td>
            <td height="25" bgcolor="#F9F9F9"><textarea name="qydt" cols="50" rows="5" id="qydt"><%=rs("qydt")%></textarea></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9">&nbsp;</td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">项目LOGO:</td>
            <td height="25" bgcolor="#F9F9F9"><input type="file" name="file1" id="file2">
              
              <a href="../upload/<%=rs("filedown1")&rs("filetype1")%>" target="_blank"><%=rs("filename1")%></a>
              <span class="STYLE4">*LOGO大小为110*73像素</span></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">旗帜图片：</td>
            <td height="25" bgcolor="#F9F9F9"><input type="file" name="file2" id="file3">
              <a href="../upload/<%=rs("filedown2")&rs("filetype2")%>" target="_blank"><%=rs("filename2")%></a><span class="STYLE4">*旗帜图片大小为702*220像素</span></td>
          </tr>
		   

          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9">&nbsp;</td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9"><input type="submit" name="Submit2" value=" 提交 "  />
            &nbsp;&nbsp;
              <input type="reset" name="Submit22" value=" 重设 " />
              &nbsp;&nbsp;
              <input type="button" name="Submit222" value=" 返回 " onClick="javascript:history.back()" />
              <input name="act" type="hidden" id="act" value="edit" />
              <input name="id" type="hidden" id="id" value="<%=id%>" /></td>
          </tr>
        </table>   
          </form>     <p>
            <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0','border','0','width','307','height','238','style','float: right; display:block; top:0; position: relative; z-index: 1; left: 0;','src','images/po','pluginspage','http://www.macromedia.com/go/getflashplayer','quality','High','wmode','transparent','movie','images/po' ); //end AC code
</script>
            <noscript><object classid="clsid:D27CDB6E-AE6D-11CF-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0" border="0" width="307" height="238" style="float: right; display:block; top:0; position: relative; z-index: 1; left: 0;">
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
<%
end if
rs.close
set rs=nothing
%>
</body>
</html>
