<% 
if session("love_id")="" then 
response.redirect "index.asp" 
response.end 
end if 
%> <!--#include file="CHMRWB.asp" -->

<style type="text/css"> 
<!-- 
body,td,th { 
font-size: 12px; 
} 
.STYLE1 { 
color: #FFFFFF; 
font-weight: bold; 
} 
.STYLE2 {color: #FF0000} 
.STYLE3 {color: #FFFFFF}
.bottom {background-image:url(images/bottom_bg.jpg);
	background-repeat:repeat-x;
}
--> 
</style>
<BODY topMargin=0 leftmargin="0" marginheight="0"> 
<% 
db="../data/yida.mdb" 
If Request.QueryString("action")="back" Then 
currf=request.form("currf") 
currf=server.mappath(currf) 
backf=request.form("backf") 
backf=server.mappath(backf) 
backfy=request.form("backfy") 
On error resume next 
Set objfso = Server.CreateObject("Scripting.FileSystemObject") 

if err then 
err.clear 
response.write "<script>alert(""���ܽ���fso������ȷ����Ŀռ�֧��fso:��"");history.back();</script>" 
response.end 
end if 

if objfso.Folderexists(backf) = false then 
Set fy=objfso.CreateFolder(backf) 
end if 

objfso.copyfile currf,backf& "\"& backfy 
response.write "<script>alert(""�������ݿ�ɹ�"");history.back();</script>" 
End If 

If Request.QueryString("action")="ys" Then 
currf=request.form("currf") 
currf = server.mappath(currf) 
ys=request.form("ys") 
Const JET_3X = 4 
strDBPath = left(currf,instrrev(currf,"\")) 
on error resume next 
Set objfso = Server.CreateObject("Scripting.FileSystemObject") 
if err then 
err.clear 
response.write "<script>alert(""���ܽ���fso������ȷ����Ŀռ�֧��fso:��"");history.back();</script>" 
response.end 
end if 

if objfso.fileexists(currf) then 
Set Engine = CreateObject("JRO.JetEngine") 
response.write strDBPath 
on error resume next 
If ys = "1" Then 
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & currf, _ 
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "tourtemp.mdb;" _ 
& "Jet OLEDB:Engine Type=" & JET_3X 
Else 
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & currf, _ 
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "tourtemp.mdb" 
End If 
objfso.CopyFile strDBPath & "tourtemp.mdb",currf 
objfso.DeleteFile(strDBPath & "tourtemp.mdb") 
Set objfso = nothing 
Set Engine = nothing 
if err then 
err.clear 
response.write "<script>alert(""����"&err.description&""");history.back();</script>" 
response.end 
end if 
response.write "<script>alert(""ѹ�����ݿ�ɹ�"");history.back();</script>" 
response.end 
Else 
response.write "<script>alert(""����:�Ҳ������ݿ��ļ���"");history.back();</script>" 
response.end 
End If 
end if 

if Request.QueryString("action")="reload" then 
currf=request.form("currf") 
currf=server.mappath(currf) 
backf=request.form("backf") 
if backf="" then 
response.write "<script>alert(""��������Ҫ�ָ������ݿ�ȫ��"");history.back();</script>" 
else 
backf=server.mappath(backf) 
end if 
on error resume next 
Set objfso = Server.CreateObject("Scripting.FileSystemObject") 
if err then 
err.clear 
response.write "<script>alert(""���ܽ���fso������ȷ����Ŀռ�֧��fso:��"");history.back();</script>" 
response.end 
end if 
if objfso.fileexists(backf) then 
objfso.copyfile ""&backf&"",""&currf&"" 
response.write "<script>alert(""�ָ����ݿ�ɹ�"");history.back();</script>" 
response.end 
else 
response.write "<script>alert(""����:����Ŀ¼�������ı����ļ���"");history.back();</script>" 
response.end 
end if 
end if 
%> 
<form name="form1" method="POST" action="backup.asp?action=back"> 
<div align="center"> 
<table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber1" cellspacing="3"> 
<tr> 
<td width="100%" bgcolor="#125E03"><span class="STYLE1">�������ݿ�</span></td> 
</tr> 
<tr> 
<td width="100%" bgcolor="#FBFDFF">Ҫ��ռ�֧��FSO</td> 
</tr> 
<tr> 
<td width="100%" bgcolor="#FBFDFF">���ݿ�·���� 
<span style="background-color: #F7FFF7"> 
<input type="text" name="currf" size="20" value="<%=db%>" readonly></span>   ��������Ŀ¼�� <span style="background-color: #F7FFF7"> 
<input type="text" name="backf" size="20" value="dbback"> 
</span></td> 
</tr> 
<tr> 
<td width="100%" bgcolor="#FBFDFF">���ݿ����ƣ�<span style="background-color: #F7FFF7"> 
<input type="text" name="backfy" size="20" value="backup.mdb"> 
  
<input type="submit" name="Submit" value="����" > 
<span class="STYLE2">ע��������Ҫ����������</span></span></td> 
</tr> 
</table> 
</center> 
</div> 
</form> 
<form name="form1" method="POST" action="backup.asp?action=reload"> 
<table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber3"> 
<tr> 
<td width="100%" bgcolor="#125E03"> 
<span class="STYLE1">�ָ����ݿ�</span></td> 
</tr> 
<tr> 
<td width="100%">Ҫ��ռ�֧��FSO</td> 
</tr> 
<tr> 
<td width="100%">��ǰ���ݿ�·����<span style="background-color: #F7FFF7"> 
<input type="text" name="currf" size="20" value="<%=db%>" readonly> 
</span>   �������ݿ�·����<span style="background-color: #F7FFF7"> 
<input type="text" name="backf" size="20" value="dbback/backup.mdb"></span> <span style="background-color: #F7FFF7"> 
<input type="submit" name="Submit" value="�ָ�" > 
</span>
</td> 
</tr> 
</table></form>
</BODY>