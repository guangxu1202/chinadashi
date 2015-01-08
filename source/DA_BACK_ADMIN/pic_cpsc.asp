
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
<script type="text/javascript"> 
	
	
	
	var i=0;
        function AddMore(){ 
            var more = document.getElementById("file"); 
            var br = document.createElement("br"); 
            var input = document.createElement("input"); 
            var button = document.createElement("input"); 
			
			i=i+1;
            if (i>5){
			alert("最多添加6张图片")
			return false;
			}
            input.type = "file"; 
            input.name = "file"+i; 
            
            button.type = "button"; 
            button.value = "删除"; 
            
            more.appendChild(br); 
            more.appendChild(input); 
            more.appendChild(button); 
            
            button.onclick = function(){ 
                more.removeChild(br); 
                more.removeChild(input); 
                more.removeChild(button); 
            }; 
        } 
    </script> 
  <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body>
<!--#include file="top.asp" -->
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
    <td valign="top" class="right">您的位置&gt;&gt;后台管理&gt;&gt;<span class="tag">产品视窗图</span></td>
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
          <form action="pic_sql.asp" method="post" enctype="multipart/form-data" name="form1" id="form1">
          
          <table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#F0F0F0">
          
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">图片类型：</td>
            <td height="25" bgcolor="#F9F9F9">
            <select name="piclx" id="piclx">
              <option value="建筑">建筑</option>
              <option value="景观">景观</option>
              <option value="细节">细节</option>
            </select>
            </td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9">
            <%
				
				id = request("id")
				set rs=server.CreateObject("adodb.recordset")
				sql="select * from project where id="&id
				rs.open sql,conn,1,1
				
					if rs("filename6")<>"" then
					
					%>
					<%
				n=split(rs("filename6"),",")
				a=split(rs("filedown6"),",")
				b=split(rs("filetype6"),",")
				
				for i=lbound(n) to ubound(n)-1
				  response.Write("【"&left(n(i),2)&"】&nbsp;&nbsp;"&"<a href='../upload/"&a(i)&b(i)&"'  target='_blank'>"&right(n(i),len(n(i))-2)&"</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='success_delpic6.asp?filedown="&a(i)&"&id="&id&"'>删除</a><br /><br />")
				next
				%>
                    <%
					end if
				rs.close
				set rs=nothing
			
			%>            </td>
          </tr>
          <tr>
            <td width="11%" height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td width="89%" height="25" bgcolor="#F9F9F9">&nbsp;</td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">产品视窗图:</td>
            <td height="25" bgcolor="#F9F9F9"><input type="file" name="filex" id="file9">
              <span class="STYLE4">*产品视窗图大小为600*400像素</span></td>
          </tr>
		   

          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9"><input type="button"  onClick="AddMore()" name="Submit" value=" 继续添加 " />
              <span class="STYLE4">*每次最多可以添加6张图片，上传多张图片可能需要一段时间，不要关闭网页</span></td>
          </tr>
          <tr>
            <td height="25" align="center" bgcolor="#F9F9F9">&nbsp;</td>
            <td height="25" bgcolor="#F9F9F9" id="file">&nbsp;</td>
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
              <input name="act" type="hidden" id="act" value="cpsc" />
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
</body>
</html>
