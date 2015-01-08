<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<!--#include file="../DA_CHMRW/sjcatstudio.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>大实集团</title>
</head>
<body>
<%

	

  set upload = new sjCat_Upload ''建立上传对象
  set uploadFile1 = upload.file("file1")
  set uploadFile2 = upload.file("file2")
  
 if upload.form("xmmc")="" then
	response.Write("<script>alert('项目名称不能为空');history.back()</script>")
	response.End()
 end if
 
 if upload.form("act")="add" then
	  if uploadFile1.filename="" then
		response.Write("<script>alert('项目LOGO不能为空');history.back()</script>")
		response.End()
	 end if
	 
	  if uploadFile2.filename="" then
		response.Write("<script>alert('旗帜广告不能为空');history.back()</script>")
		response.End()
	 end if
  end if


  filename1 = uploadFile1.filename
  filename2 = uploadFile2.filename

  filepath = server.MapPath("../upload")
  
  
  if filename1<>"" then
  filename1=replace(filename1,",","_")
	filedown1 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_1"
	filetype1 = strReverse(left(strReverse(filename1),instr(strReverse(filename1),".")))
	
    uploadFile1.Save2File filepath&"\"&filedown1&filetype1
	
  end if 
  if filename2<>"" then
  filename2=replace(filename2,",","_")
	filedown2 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_2"
	filetype2 = strReverse(left(strReverse(filename2),instr(strReverse(filename2),".")))
    uploadFile2.Save2File filepath&"\"&filedown2&filetype2
  end if 
  
  
  
  
  
  if upload.form("act")="add" then
  
  
  set rs=server.CreateObject("adodb.recordset")
  sql="select * from project"
  rs.open sql,conn,1,3
  rs.addnew
  	  rs("xmmc")=upload.form("xmmc")
	  rs("slrx")=upload.form("slrx")
	  rs("cpjg")=upload.form("cpjg")
	  rs("zdmj")=upload.form("zdmj")
	  rs("jzmj")=upload.form("jzmj")
	  rs("ghyt")=upload.form("ghyt")
	  rs("hxmj")=upload.form("hxmj")
	  rs("rzsj")=upload.form("rzsj")
	  rs("gjxl")=upload.form("gjxl")
	  rs("xmqw")=upload.form("xmqw")
	  rs("zbpt")=upload.form("zbpt")
	  rs("xmjs")=upload.form("xmjs")
	  rs("cpxx")=upload.form("cpxx")
	  rs("qydt")=upload.form("qydt")
	  rs("sendtime")=now
	  rs("filename1")=filename1
	  rs("filedown1")=filedown1
	  rs("filetype1")=filetype1
	  rs("filename2")=filename2
	  rs("filedown2")=filedown2
	  rs("filetype2")=filetype2
  rs.update
  rs.close
  set rs=nothing
  

  
  end if
  
  
  if upload.form("act")="edit" then
  id = upload.form("id")
  
  set rs=server.CreateObject("adodb.recordset")
  sql="select * from project where id="&id
  rs.open sql,conn,1,3
  	  rs("xmmc")=upload.form("xmmc")
	  rs("slrx")=upload.form("slrx")
	  rs("cpjg")=upload.form("cpjg")
	  rs("zdmj")=upload.form("zdmj")
	  rs("jzmj")=upload.form("jzmj")
	  rs("ghyt")=upload.form("ghyt")
	  rs("hxmj")=upload.form("hxmj")
	  rs("rzsj")=upload.form("rzsj")
	  rs("gjxl")=upload.form("gjxl")
	  rs("xmqw")=upload.form("xmqw")
	  rs("zbpt")=upload.form("zbpt")
	  rs("xmjs")=upload.form("xmjs")
	  rs("cpxx")=upload.form("cpxx")
	  rs("qydt")=upload.form("qydt")
	  rs("sendtime")=now
	  if filename1<>"" then
		  rs("filename1")=filename1
		  rs("filedown1")=filedown1
		  rs("filetype1")=filetype1
	  end if
	  if filename2<>"" then
		  rs("filename2")=filename2
		  rs("filedown2")=filedown2
		  rs("filetype2")=filetype2
	  end if
  rs.update
  rs.close
  set rs=nothing
  

  
  end if
  

  
  
  response.Redirect("success_main.asp?mm=2")


%>


</body>
</html>