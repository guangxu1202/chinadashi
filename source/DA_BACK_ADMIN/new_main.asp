<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
manage_end=request.QueryString("manage_end")
q=request("q")
search=request("search")
if search="" then
	search=1
end if
if q<>"" then
	q=replace(q," ","%")
	keyword= "%"&q&"%"
end if


%>
<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<%
Page=Request.QueryString("page")
If Page=Empty then Page=1
%>
<script language="JavaScript">
function del()
{
if (confirm("ȷ��Ҫɾ����"))
  {
   document.form1.submit();
  }
}
</script>
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����-�û�����</title>

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
    <td valign="top" class="right">����λ��&gt;&gt;��̨����&gt;&gt;<span class="tag">���Ŷ�̬����</span></td>
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
        <td height="500" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#F9F9F9">
        </table> 
		
		
		
		
		
		
		
		
		
		
		       
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="main">
            <form action="new_del.asp" method="post" name="form1" id="form1">
              <tr>
                <td align="center" bgcolor="#F5F5F5">&nbsp;</td>
                <td height="30" align="center" bgcolor="#F5F5F5"><strong>��������</strong></td>
                <td align="center" bgcolor="#F5F5F5"><strong>���</strong></td>
                <td align="center" bgcolor="#F5F5F5"><strong>¼����</strong></td>
                <td align="center" bgcolor="#F5F5F5"><strong>¼��ʱ��</strong></td>
                <td align="center" bgcolor="#F5F5F5">&nbsp;</td>
                <td align="center" bgcolor="#F5F5F5"><a href="new_add.asp?mm=1">���</a></td>
              </tr>
              <%
		set rs=server.createobject("adodb.recordset")
			sql="select * from new order by sendtime desc"
	
          rs.open sql,conn,1,1
		  if not rs.eof and not rs.bof then
		     sum = rs.recordcount
			 rs.pagesize = 16
			 last = rs.pagecount
			 if cint(page) > last  then page = last
			 rs.AbsolutePage = page
			 for i = 1 to rs.pagesize
			     check=1
				 
			    text=""
				text=rs("title")
				length=len(text)
				if length<38 then
				   title=text
				else
				   title=left(text,38)&"....."
				end if
		%>
              <tr>
                <td width="29" align="center" <%if i mod 2 =0 then%> bgcolor="#F5F5F5"<%end if%>><div align="center"><%=i%></div></td>
                <td width="323" align="center" <%if i mod 2 =0 then%> bgcolor="#F5F5F5"<%end if%>><%=title%></td>
                <td width="143" align="center" <%if i mod 2 =0 then%> bgcolor="#F5F5F5"<%end if%>>
                <%
				if rs("tag")=1 then
					response.Write("��˾����")
				elseif rs("tag")=2 then
					response.Write("��ҵ����")
				elseif rs("tag")=3 then
					response.Write("ý�屨��")
				end if
				%>
                </td>
                <td width="80" align="center" <%if i mod 2 =0 then%> bgcolor="#F5F5F5"<%end if%>><%=rs("lrr")%></td>
                <td width="100" height="25" align="center" <%if i mod 2 =0 then%> bgcolor="#F5F5F5"<%end if%>><%=year(rs("sendtime"))&"-"&month(rs("sendtime"))&"-"&day(rs("sendtime"))%></td>
                
                <td width="60" align="center" <%if i mod 2 =0 then%> bgcolor="#F5F5F5" <%end if%>><a href="new_edit.asp?id=<%=rs("id")%>">�޸�</a></td>
                <td width="86" align="center" <%if i mod 2 =0 then%> bgcolor="#F5F5F5"<%end if%>><input name="<%=rs("id")%>" type="checkbox" id="<%=rs("id")%>" value="<%=rs("id")%>" /></td>
              </tr>
              <tr>
                <td width="29" align="center" valign="top" background="../images/line.gif"></td>
                <td colspan="5" align="center" valign="top" background="../images/line.gif"></td>
                <td align="center" valign="top" background="../images/line.gif"></td>
              </tr>
              <%
        rs.movenext
		if rs.eof  then exit for
	      next	   
		 rs.close
		 set rs = nothing
	   %>
              <tr>
                <td width="29" height="32">&nbsp;</td>
                <td colspan="5" align="left">&nbsp;</td>
                <td align="center"><font color="#FF6600">
                  <input name="act" type="hidden" id="act" value="del" />
                  &nbsp;</font>
                    <input name="Submit2" type="button" class="button" value="ɾ��" onclick="del()" /></td>
              </tr>
              <%
		end if
			  %>
            </form>
          </table>
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
          <%if sum>16 then%>
          <table width="468" border="0" align="center" cellpadding="0" cellspacing="0">
            <%If Page=1 Then%>
            <tr>
              <td width="57">&nbsp;</td>
              <td width="59">&nbsp;</td>
              <td width="57"><div align="center"><a href="new_main.asp?mm=2&page=<%=Page+1%>">[��һҳ]</a></div></td>
              <td width="73"><div align="center"><a href="new_main.asp?mm=2&page=<%=Last%>">[���һҳ]</a></div></td>
              <td width="84"><div align="center"><font color="#666666">ҳ����
                <%response.Write page&"/"&last%>
                ҳ</font></div></td>
              <td width="73"><div align="center"><font color="#666666">�ܹ�:<%=sum%>��¼</font></div></td>
              <td width="65">��ǰҳ<%=page%></td>
            </tr>
            <% ElseIf Cint(Page)=Last Then%>
            <tr>
              <td width="57">&nbsp;</td>
              <td width="59">&nbsp;</td>
              <td width="57"><div align="center"><a href="new_main.asp?mm=2&page=1">[��һҳ]</a></div></td>
              <td width="73"><div align="center"><a href="new_main.asp?mm=2&page=<%=Page-1%>">[��һҳ]</a></div></td>
              <td><div align="center"><font color="#666666">ҳ����
                <%response.Write page&"/"&last%>
                ҳ</font></div></td>
              <td><div align="center"><font color="#666666">�ܹ�:<%=sum%>��¼</font></div></td>
              <td>��ǰҳ<%=page%></td>
            </tr>
            <%Else%>
            <tr>
              <td width="57"><div align="center"><a href="new_main.asp?mm=2&page=1">[��һҳ]</a></div></td>
              <td width="59"><div align="center"><a href="new_main.asp?mm=2&page=<%=Page-1%>">[��һҳ]</a></div></td>
              <td width="57"><div align="center"><a href="new_main.asp?mm=2&page=<%=Page+1%>">[��һҳ]</a></div></td>
              <td width="73"><div align="center"><a href="new_main.asp?mm=2&page=<%=Last%>">[���һҳ]</a></div></td>
              <td><div align="center"><font color="#666666">ҳ����
                <%response.Write page&"/"&last%>
                ҳ</font></div></td>
              <td><div align="center"><font color="#666666">�ܹ�:<%=sum%>��¼</font></div></td>
              <td>��ǰҳ<%=page%></td>
            </tr>
            <%End if%>
          </table>
          <%end if%><p><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
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
