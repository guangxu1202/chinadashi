<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨������ҳ</title>
<link href="images/admin.css" type="text/css" rel="stylesheet" />
<%

Response.Buffer = true

' �������������
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
	ObjTotest(9,1) = "(FSO �ı��ļ���д)"
ObjTotest(10,0) = "adodb.connection"
	ObjTotest(10,1) = "(ADO ���ݶ���)"
	
ObjTotest(11,0) = "SoftArtisans.FileUp"
	ObjTotest(11,1) = "(SA-FileUp �ļ��ϴ�)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
	ObjTotest(12,1) = "(SoftArtisans �ļ�����)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
	ObjTotest(13,1) = "(���Ʒ���ļ��ϴ����)"
ObjTotest(14,0) = "Persits.Upload.1"
	ObjTotest(14,1) = "(ASPUpload �ļ��ϴ�)"
ObjTotest(15,0) = "w3.upload"
	ObjTotest(15,1) = "(Dimac �ļ��ϴ�)"

ObjTotest(16,0) = "JMail.SmtpMail"
	ObjTotest(16,1) = "(Dimac JMail �ʼ��շ�) <a href='http://www.ajiang.net'>�����ֲ�����</a>"
ObjTotest(17,0) = "CDONTS.NewMail"
	ObjTotest(17,1) = "(���� SMTP ����)"
ObjTotest(18,0) = "Persits.MailSender"
	ObjTotest(18,1) = "(ASPemail ����)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
	ObjTotest(19,1) = "(ASPmail ����)"
ObjTotest(20,0) = "DkQmail.Qmail"
	ObjTotest(20,1) = "(dkQmail ����)"
ObjTotest(21,0) = "Geocel.Mailer"
	ObjTotest(21,1) = "(Geocel ����)"
ObjTotest(22,0) = "IISmail.Iismail.1"
	ObjTotest(22,1) = "(IISmail ����)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(23,1) = "(SmtpMail ����)"
	
ObjTotest(24,0) = "SoftArtisans.ImageGen"
	ObjTotest(24,1) = "(SA ��ͼ���д���)"
ObjTotest(25,0) = "W3Image.Image"
	ObjTotest(25,1) = "(Dimac ��ͼ���д���)"

public IsObj,VerObj,TestObj
public okOS,okCPUS,okCPU

'���Ԥ�����֧��������汾

dim i
for i=0 to 25
	on error resume next
	IsObj=false
	VerObj=""
	'dim TestObj
	TestObj=""
	set TestObj=server.CreateObject(ObjTotest(i,0))
	If -2147221005 <> Err then		'��л����iAmFisher�ı�����
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'�������Ƿ�֧�ּ�����汾���ӳ���
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	TestObj=""
	set TestObj=server.CreateObject (strObj)
	If -2147221005 <> Err then		'��л����iAmFisher�ı�����
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
    <td valign="top" class="right"><strong>����λ��</strong>&gt;&gt;<span class="tag">��̨������ҳ</span></td>
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
        <td height="20"><font class="fonts">�Ƿ�֧��ASP</font></td>
      </tr>
      <tr>
        <td>&nbsp;&nbsp;���������������ʾ���Ŀռ䲻֧��ASP�� <br />
&nbsp;&nbsp;1�����ʱ��ļ�ʱ��ʾ���ء� <br />
&nbsp;&nbsp;2�����ʱ��ļ�ʱ�������ơ�&lt;%@ Language=&quot;VBScript&quot; %&gt;�������֡�</td>
      </tr>
      <tr>
        <td height="20"><font class="fonts">���������йز���</font></td>
      </tr>
      <tr>
        <td><table width="450" border="1" cellpadding="0" cellspacing="0" bordercolor="#F0F0F0">
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;��������</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;������IP</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;�������˿�</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;������ʱ��</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=now%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;IIS�汾</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;�ű���ʱʱ��</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Server.ScriptTimeout%> ��</td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;���ļ�·��</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;��������������</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
            </tr>
            <%getsysinfo()  '��÷���������%>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;������CPU����</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=okCPUS%> ��</td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;������CPU����</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=okCPU%></td>
            </tr>
            <tr bgcolor="#EEFEE0" height="18">
              <td align="left" bgcolor="#F9F9F9">&nbsp;����������ϵͳ</td>
              <td bgcolor="#F9F9F9">&nbsp;<%=okOS%></td>
            </tr>
          </table></td>
      </tr>
      <tr>
        <td height="25"><font class="fonts">���֧�����</font>
          <%
Dim strClass
	strClass = Trim(Request.Form("classname"))
	If "" <> strClass then
	Response.Write "<br>��ָ��������ļ������"
	Dim Verobj1
	ObjTest(strClass)
	  If Not IsObj then 
		Response.Write "<br><font color=red>���ź����÷�������֧�� " & strclass & " �����</font>"
	  Else
		if VerObj="" or isnull(VerObj) then 
			Verobj1="�޷�ȡ�ø�����汾"
		Else
			Verobj1="������汾�ǣ�" & VerObj
		End If
		Response.Write "<br><font class=fonts>��ϲ���÷�����֧�� " & strclass & " �����" & verobj1 & "</font>"
	  End If
	  Response.Write "<br>"
	end if
	%></td>
      </tr>
      <tr>
        <td>�� IIS�Դ���ASP���
          <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
            <tr height="18" class="backs" align="center">
              <td width="320" bgcolor="#666666">�� �� �� ��</td>
              <td width="130" bgcolor="#666666">֧�ּ��汾</td>
            </tr>
            <%For i=0 to 10%>
            <tr height="18" class="backq">
              <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
              <td align="left" bgcolor="#F9F9F9">&nbsp;
                  <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
            </tr>
            <%next%>
          </table></td>
      </tr>
      <tr>
        <td><br />
�� �������ļ��ϴ��͹������
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
    <tr height="18" class="backs" align="center">
      <td width="320" bgcolor="#666666">�� �� �� ��</td>
      <td width="130" bgcolor="#666666">֧�ּ��汾</td>
    </tr>
    <%For i=11 to 15%>
    <tr height="18" class="backq">
      <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
      <td align="left" bgcolor="#F9F9F9">&nbsp;
          <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
    </tr>
    <%next%>
  </table></td>
      </tr>
      <tr>
        <td>�� �������շ��ʼ����
          <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
            <tr height="18" class="backs" align="center">
              <td width="320" bgcolor="#666666">�� �� �� ��</td>
              <td width="130" bgcolor="#666666">֧�ּ��汾</td>
            </tr>
            <%For i=16 to 23%>
            <tr height="18" class="backq">
              <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
              <td align="left" bgcolor="#F9F9F9">&nbsp;
                  <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
            </tr>
            <%next%>
          </table></td>
      </tr>
      <tr>
        <td><br />
�� ͼ�������
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#F0F0F0" width="450">
    <tr height="18" class="backs" align="center">
      <td width="320" bgcolor="#666666">�� �� �� ��</td>
      <td width="130" bgcolor="#666666">֧�ּ��汾</td>
    </tr>
    <%For i=24 to 25%>
    <tr height="18" class="backq">
      <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
      <td align="left" bgcolor="#F9F9F9">&nbsp;
          <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
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
