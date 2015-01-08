<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理首页</title>
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<%

Response.Buffer = true

' 声明待检测数组
Dim ObjTotest(26,4)

ObjTotest(0,0) = "MSWC.AdRotator"
ObjTotest(1,0) = "MSWC.BrowserType"
ObjTotest(2,0) = "MSWC.NextLink"
ObjTotest(3,0) = "MSWC.Tools"
ObjTotest(4,0) = "MSWC.Status"
ObjTotest(5,0) = "MSWC.Counters"
ObjTotest(6,0) = "IISSample.ContentRotator"
ObjTotest(7,0) = "IISSample.PageCounter"
ObjTotest(8,0) = "MSWC.PermissionChecker"
ObjTotest(9,0) = "Scripting.FileSystemObject"
	ObjTotest(9,1) = "(FSO 文本文件读写)"
ObjTotest(10,0) = "adodb.connection"
	ObjTotest(10,1) = "(ADO 数据对象)"
	
ObjTotest(11,0) = "SoftArtisans.FileUp"
	ObjTotest(11,1) = "(SA-FileUp 文件上传)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
	ObjTotest(12,1) = "(SoftArtisans 文件管理)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
	ObjTotest(13,1) = "(刘云峰的文件上传组件)"
ObjTotest(14,0) = "Persits.Upload.1"
	ObjTotest(14,1) = "(ASPUpload 文件上传)"
ObjTotest(15,0) = "w3.upload"
	ObjTotest(15,1) = "(Dimac 文件上传)"

ObjTotest(16,0) = "JMail.SmtpMail"
	ObjTotest(16,1) = "(Dimac JMail 邮件收发) <a href='http://www.ajiang.net'>中文手册下载</a>"
ObjTotest(17,0) = "CDONTS.NewMail"
	ObjTotest(17,1) = "(虚拟 SMTP 发信)"
ObjTotest(18,0) = "Persits.MailSender"
	ObjTotest(18,1) = "(ASPemail 发信)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
	ObjTotest(19,1) = "(ASPmail 发信)"
ObjTotest(20,0) = "DkQmail.Qmail"
	ObjTotest(20,1) = "(dkQmail 发信)"
ObjTotest(21,0) = "Geocel.Mailer"
	ObjTotest(21,1) = "(Geocel 发信)"
ObjTotest(22,0) = "IISmail.Iismail.1"
	ObjTotest(22,1) = "(IISmail 发信)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(23,1) = "(SmtpMail 发信)"
	
ObjTotest(24,0) = "SoftArtisans.ImageGen"
	ObjTotest(24,1) = "(SA 的图像读写组件)"
ObjTotest(25,0) = "W3Image.Image"
	ObjTotest(25,1) = "(Dimac 的图像读写组件)"

public IsObj,VerObj,TestObj
public okOS,okCPUS,okCPU

'检查预查组件支持情况及版本

dim i
for i=0 to 25
	on error resume next
	IsObj=false
	VerObj=""
	'dim TestObj
	TestObj=""
	set TestObj=server.CreateObject(ObjTotest(i,0))
	If -2147221005 <> Err then		'感谢网友iAmFisher的宝贵建议
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'检查组件是否被支持及组件版本的子程序
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	TestObj=""
	set TestObj=server.CreateObject (strObj)
	If -2147221005 <> Err then		'感谢网友iAmFisher的宝贵建议
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if	
End sub
%>

<script type="text/javascript" src="../images/nav.js"></script>
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body><!--#include file="top.asp" -->

<table width="1000" height="600" border="0" align="center" cellpadding="0" cellspacing="0" class="bxline">
  <tr>
    <td valign="top">&nbsp;</td>
    <td valign="top">&nbsp;</td>
    <td>&nbsp;</td>
    <td valign="top" class="right">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td valign="top">&nbsp;</td>
    <td valign="top">&nbsp;</td>
    <td>&nbsp;</td>
    <td valign="top" class="right"><strong>您的位置</strong>&gt;&gt;<span class="tag">后台管理首页</span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="50" valign="top">&nbsp;</td>
    <td width="164" valign="top"><!--#include file="left_default.asp" --></td>
    <td width="25" class="leftline">&nbsp;</td>
    
    <td width="780" valign="top" class="right"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="d_right">
      <tr>
        <td height="15">&nbsp;</td>
      </tr>
      <tr>
        <td height="20"><font class="fonts">是否支持ASP</font></td>
      </tr>
      <tr>
        <td>&nbsp;&nbsp;出现以下情况即表示您的空间不支持ASP： <br />
&nbsp;&nbsp;1、访问本文件时提示下载。 <br />
&nbsp;&nbsp;2、访问本文件时看到类似“&lt;%@ Language=&quot;VBScript&quot; %&gt;”的文字。</td>
      </tr>
      <tr>
        <td height="20"><font class="fonts">服务器的有关参数</font></td>
      </tr>
      <tr>
        <td><table width="450" border="1" cellpadding="0" cellspacing="0" bordercolor="#F0F0F0">
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器名</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器IP</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器端口</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器时间</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=now%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;IIS版本</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;脚本超时时间</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Server.ScriptTimeout%> 秒</td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;本文件路径</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器解译引擎</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
            </tr>
            <%getsysinfo()  '获得服务器数据%>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器CPU数量</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=okCPUS%> 个</td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器CPU详情</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=okCPU%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;服务器操作系统</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=okOS%></td>
            </tr>
          </table></td>
      </tr>
      <tr>
        <td height="25"><font class="fonts">组件支持情况</font>
          <%
Dim strClass
	strClass = Trim(Request.Form("classname"))
	If "" <> strClass then
	Response.Write "<br>您指定的组件的检查结果："
	Dim Verobj1
	ObjTest(strClass)
	  If Not IsObj then 
		Response.Write "<br><font color=red>很遗憾，该服务器不支持 " & strclass & " 组件！</font>"
	  Else
		if VerObj="" or isnull(VerObj) then 
			Verobj1="无法取得该组件版本"
		Else
			Verobj1="该组件版本是：" & VerObj
		End If
		Response.Write "<br><font class=fonts>恭喜！该服务器支持 " & strclass & " 组件。" & verobj1 & "</font>"
	  End If
	  Response.Write "<br>"
	end if
	%></td>
      </tr>
      <tr>
        <td>■ IIS自带的ASP组件
          <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
            <tr height="18" class="backs" align="center">
              <td width="320" bgcolor="#666666">组 件 名 称</td>
              <td width="130" bgcolor="#666666">支持及版本</td>
            </tr>
            <%For i=0 to 10%>
            <tr height="18" class="backq">
              <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
              <td align="left" bgcolor="#F9F9F9">&nbsp;
                  <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
            </tr>
            <%next%>
          </table></td>
      </tr>
      <tr>
        <td><br />
■ 常见的文件上传和管理组件
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
    <tr height="18" class="backs" align="center">
      <td width="320" bgcolor="#666666">组 件 名 称</td>
      <td width="130" bgcolor="#666666">支持及版本</td>
    </tr>
    <%For i=11 to 15%>
    <tr height="18" class="backq">
      <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
      <td align="left" bgcolor="#F9F9F9">&nbsp;
          <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
    </tr>
    <%next%>
  </table></td>
      </tr>
      <tr>
        <td>■ 常见的收发邮件组件
          <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
            <tr height="18" class="backs" align="center">
              <td width="320" bgcolor="#666666">组 件 名 称</td>
              <td width="130" bgcolor="#666666">支持及版本</td>
            </tr>
            <%For i=16 to 23%>
            <tr height="18" class="backq">
              <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
              <td align="left" bgcolor="#F9F9F9">&nbsp;
                  <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
            </tr>
            <%next%>
          </table></td>
      </tr>
      <tr>
        <td><br />
■ 图像处理组件
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
    <tr height="18" class="backs" align="center">
      <td width="320" bgcolor="#666666">组 件 名 称</td>
      <td width="130" bgcolor="#666666">支持及版本</td>
    </tr>
    <%For i=24 to 25%>
    <tr height="18" class="backq">
      <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
      <td align="left" bgcolor="#F9F9F9">&nbsp;
          <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		Else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
    </tr>
    <%next%>
  </table>
  <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0','border','0','width','307','height','238','style','float: right; display:block; top:0; position: relative; z-index: 1; left: 0;','src','images/po','pluginspage','http://www.macromedia.com/go/getflashplayer','quality','High','wmode','transparent','movie','images/po' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11CF-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0" border="0" width="307" height="238" style="float: right; display:block; top:0; position: relative; z-index: 1; left: 0;">
    <param name="movie" value="images/po.swf" />
    <param name="quality" value="High" />
    <param name="wmode" value="transparent" />
    <embed src="images/po.swf" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="307" height="238" quality="High" wmode="transparent"> </embed>
  </object></noscript></td>
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
