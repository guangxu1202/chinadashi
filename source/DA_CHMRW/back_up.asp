<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>数据备份与恢复</title>
<style type="text/css">
<!--
.navlink {display:block;
	width:69px;
	height:30px;
}


#ccc { background:url(images/nav1.jpg) no-repeat 0 0;}
#bbb{ background:url(images/nav1.jpg) no-repeat 0 -32px;}
.bbb { background:url(images/nav1.jpg) no-repeat 0 -32px;}
.ccc { background:url(images/nav1.jpg) no-repeat 0 0;}
.nav_new1{ background:url(../ART_BACK_ADMIN/images/nav_client.jpg) no-repeat 0 0;}
.nav_new2{ background:url(../ART_BACK_ADMIN/images/nav_client.jpg) no-repeat 0 -32px;}
#nav_new1{ background:url(../ART_BACK_ADMIN/images/nav_client.jpg) no-repeat 0 0;}
body{
	margin:0px;
	padding:0px;
	font-size:12px;
}
.STYLE2 {color: #FF6600}
.backq {	BACKGROUND-COLOR: #EEFEE0
}
.backs {	
	BACKGROUND-COLOR: #3F8805;
	COLOR: #ffffff;
}
.fonts {	COLOR: #3F8805
}
.leftnav {	border:1px solid #C2C2C2;
	background-color:#3C3B37;
}
.right {	border:1px solid #C2C2C2;
}
#d_right {	padding-left:10px;
}
.bottom {background-image:url(images/bottom_bg.jpg);
	background-repeat:repeat-x;
}
.STYLE1 {color: #FFFFFF; 
font-weight: bold; 
}

-->
</style>
<% 
db="../DINSXC/DA_afkpuz.mdb" 
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
response.write "<script>alert(""不能建立fso对象，请确保你的空间支持fso:！"");history.back();</script>" 
response.end 
end if 

if objfso.Folderexists(backf) = false then 
Set fy=objfso.CreateFolder(backf) 
end if 

objfso.copyfile currf,backf& "\"& backfy 
response.write "<script>alert(""备份数据库成功"");history.back();</script>" 
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
response.write "<script>alert(""不能建立fso对象，请确保你的空间支持fso:！"");history.back();</script>" 
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
response.write "<script>alert(""错误："&err.description&""");history.back();</script>" 
response.end 
end if 
response.write "<script>alert(""压缩数据库成功"");history.back();</script>" 
response.end 
Else 
response.write "<script>alert(""错误:找不到数据库文件！"");history.back();</script>" 
response.end 
End If 
end if 

if Request.QueryString("action")="reload" then 
currf=request.form("currf") 
currf=server.mappath(currf) 
backf=request.form("backf") 
if backf="" then 
response.write "<script>alert(""请输入您要恢复的数据库全名"");history.back();</script>" 
else 
backf=server.mappath(backf) 
end if 
on error resume next 
Set objfso = Server.CreateObject("Scripting.FileSystemObject") 
if err then 
err.clear 
response.write "<script>alert(""不能建立fso对象，请确保你的空间支持fso:！"");history.back();</script>" 
response.end 
end if 
if objfso.fileexists(backf) then 
objfso.copyfile ""&backf&"",""&currf&"" 
response.write "<script>alert(""恢复数据库成功"");history.back();</script>" 
response.end 
else 
response.write "<script>alert(""错误:备份目录下无您的备份文件！"");history.back();</script>" 
response.end 
end if 
end if 
%> 
</head>

<body>
<table width="803" height="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="6" valign="top">&nbsp;</td>
    <td width="11">&nbsp;</td>
    <td width="780" align="left" valign="top" class="right"><form action="backup.asp?action=back" method="post" name="form1" id="form1">

        <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber1" cellspacing="3">
          <tr>
            <td width="100%" bgcolor="#125E03"><span class="STYLE1">备份数据库</span></td>
          </tr>
          <tr>
            <td width="100%" bgcolor="#FBFDFF">要求空间支持FSO</td>
          </tr>
          <tr>
            <td width="100%" bgcolor="#FBFDFF">数据库路径： <span style="background-color: #F7FFF7">
              <input type="text" name="currf" size="20" value="<%=db%>" readonly="readonly" />
              </span> 备份数据目录： <span style="background-color: #F7FFF7">
              <input type="text" name="backf" size="20" value="dbback" />
            </span></td>
          </tr>
          <tr>
            <td width="100%" bgcolor="#FBFDFF">数据库名称：<span style="background-color: #F7FFF7">
              <input type="text" name="backfy" size="20" value="backup.mdb" />
              <input type="submit" name="Submit" value="备份">
              <span class="STYLE2">注：尽量不要更改以上项</span></span></td>
          </tr>
        </table>
     

    </form>
      <form action="backup.asp?action=reload" method="post" name="form1" id="form1">
        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber3">
          <tr>
            <td width="100%" bgcolor="#125E03"><span class="STYLE1">恢复数据库</span></td>
          </tr>
          <tr>
            <td width="100%">要求空间支持FSO</td>
          </tr>
          <tr>
            <td width="100%">当前数据库路径：<span style="background-color: #F7FFF7">
              <input type="text" name="currf2" size="20" value="<%=db%>" readonly="readonly" />
              </span> 备份数据库路径：<span style="background-color: #F7FFF7">
              <input type="text" name="backf2" size="20" value="dbback/backup.mdb" />
                </span> <span style="background-color: #F7FFF7">
              <input type="submit" name="Submit2" value="恢复">
            </span> </td>
          </tr>
        </table>
    </form></td>
    <td width="6">&nbsp;</td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="10">&nbsp;</td>
  </tr>
</table>
</body>
</html>
