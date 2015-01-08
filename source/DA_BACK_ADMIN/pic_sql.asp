<!--#include file="../DA_CHMRW/CHMRWB.asp"-->
<!--#include file="../DA_CHMRW/sjcatstudio.inc" -->
<%

  set upload = new sjCat_Upload ''建立上传对象
  
if upload.form("act")="xmqw" then
  set uploadFile5 = upload.file("file5")
  filename5 = uploadFile5.filename
  filepath = server.MapPath("../upload")
  
  if filename5<>"" then
  filename5=replace(filename5,",","_")
	filedown5 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_5"
	filetype5 = strReverse(left(strReverse(filename5),instr(strReverse(filename5),".")))
    uploadFile5.Save2File filepath&"\"&filedown5&filetype5
  end if 
	id=upload.form("id")
		set rs=server.CreateObject("adodb.recordset")
		sql="select * from project where id="&id
		rs.open sql,conn,1,3
			if filename5<>"" then
				rs("filename5")=filename5
				rs("filedown5")=filedown5
				rs("filetype5")=filetype5
			end if
		rs.update
		rs.close
		set rs=nothing
	response.Redirect("success_main.asp?mm=2")
end if


if upload.form("act")="xmsj" then
	
  set uploadFilex = upload.file("filex")
  
  set uploadFile1 = upload.file("file1")
  set uploadFile2 = upload.file("file2")
  set uploadFile3 = upload.file("file3")
  set uploadFile4 = upload.file("file4")
  set uploadFile5 = upload.file("file5")
  filenamex = uploadFilex.filename
  
  filename1 = uploadFile1.filename
  filename2 = uploadFile2.filename
  filename3 = uploadFile3.filename
  filename4 = uploadFile4.filename
  filename5 = uploadFile5.filename
  filepath = server.MapPath("../upload")
  filepath_sm = server.MapPath("../upload/images")


Set Jpeg = Server.CreateObject("Persits.Jpeg") '调用组件 




  if filenamex<>"" then
  filenamex=replace(filenamex,",","_")
	filedownx = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_x"
	filetypex = strReverse(left(strReverse(filenamex),instr(strReverse(filenamex),".")))
    uploadFilex.Save2File filepath&"\"&filedownx&filetypex
	
	Path = Server.MapPath("../upload")&"\"&filedownx&filetypex '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedownx&filetypex 
  end if 
  if filename1<>"" then
  filename1=replace(filename1,",","_")
	filedown1 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_1"
	filetype1 = strReverse(left(strReverse(filename1),instr(strReverse(filename1),".")))
    uploadFile1.Save2File filepath&"\"&filedown1&filetype1
	Path = Server.MapPath("../upload")&"\"&filedown1&filetype1 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown1&filetype1 
  end if 
  if filename2<>"" then
  filename2=replace(filename2,",","_")
	filedown2 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_2"
	filetype2 = strReverse(left(strReverse(filename2),instr(strReverse(filename2),".")))
    uploadFile2.Save2File filepath&"\"&filedown2&filetype2
	Path = Server.MapPath("../upload")&"\"&filedown2&filetype2 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown2&filetype2 
  end if 
  if filename3<>"" then
  filename3=replace(filename3,",","_")
	filedown3 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_3"
	filetype3 = strReverse(left(strReverse(filename3),instr(strReverse(filename3),".")))
    uploadFile3.Save2File filepath&"\"&filedown3&filetype3
	Path = Server.MapPath("../upload")&"\"&filedown3&filetype3 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown3&filetype3 
  end if 
  if filename4<>"" then
  filename4=replace(filename4,",","_")
	filedown4 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_4"
	filetype4 = strReverse(left(strReverse(filename4),instr(strReverse(filename4),".")))
    uploadFile4.Save2File filepath&"\"&filedown4&filetype4
	Path = Server.MapPath("../upload")&"\"&filedown4&filetype4 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown4&filetype4 
  end if 
  if filename5<>"" then
  filename5=replace(filename5,",","_")
	filedown5 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_5"
	filetype5 = strReverse(left(strReverse(filename5),instr(strReverse(filename5),".")))
    uploadFile5.Save2File filepath&"\"&filedown5&filetype5
	Path = Server.MapPath("../upload")&"\"&filedown5&filetype5 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown5&filetype5 
  end if 


	
	filename=filenamex&","&filename1&","&filename2&","&filename3&","&filename4&","&filename5&","
	filedown=filedownx&","&filedown1&","&filedown2&","&filedown3&","&filedown4&","&filedown5&","
	filetype=filetypex&","&filetype1&","&filetype2&","&filetype3&","&filetype4&","&filetype5&","
	
	
	a=split(filename,",")
	b=split(filedown,",")
	c=split(filetype,",")
	 
	 
	for i=LBound(a) to UBound(a)
		if a(i)<>"" then
		xx=xx&a(i)&","
		end if
	next
	for j=LBound(b) to UBound(b)
		if b(j)<>"" then
		yy=yy&b(j)&","
		end if
	next
	for k=LBound(c) to UBound(c)
		if c(k)<>"" then
		zz=zz&c(k)&","
		end if
	next

	set rs=server.CreateObject("adodb.recordset")
	sql="select * from project where id="&upload.form("id")
	rs.open sql,conn,1,3
		if xx<>"" then
			rs("filename3")=rs("filename3")&xx
			rs("filedown3")=rs("filedown3")&yy
			rs("filetype3")=rs("filetype3")&zz
		end if
	rs.update
	rs.close
	set rs=nothing
	
	
	response.Redirect("success_main.asp?mm=2")
end if



if upload.form("act")= "xmhx" then

 set uploadFilex = upload.file("filex")
  
  set uploadFile1 = upload.file("file1")
  set uploadFile2 = upload.file("file2")
  set uploadFile3 = upload.file("file3")
  set uploadFile4 = upload.file("file4")
  set uploadFile5 = upload.file("file5")
  filenamex = uploadFilex.filename
  
  filename1 = uploadFile1.filename
  filename2 = uploadFile2.filename
  filename3 = uploadFile3.filename
  filename4 = uploadFile4.filename
  filename5 = uploadFile5.filename
  filepath = server.MapPath("../upload")
  filepath_sm = server.MapPath("../upload/images")


Set Jpeg = Server.CreateObject("Persits.Jpeg") '调用组件 




  if filenamex<>"" then
  filenamex=replace(filenamex,",","_")
	filedownx = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_x"
	filetypex = strReverse(left(strReverse(filenamex),instr(strReverse(filenamex),".")))
    uploadFilex.Save2File filepath&"\"&filedownx&filetypex
	
	Path = Server.MapPath("../upload")&"\"&filedownx&filetypex '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedownx&filetypex 
  end if 
  if filename1<>"" then
  filename1=replace(filename1,",","_")
	filedown1 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_1"
	filetype1 = strReverse(left(strReverse(filename1),instr(strReverse(filename1),".")))
    uploadFile1.Save2File filepath&"\"&filedown1&filetype1
	Path = Server.MapPath("../upload")&"\"&filedown1&filetype1 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown1&filetype1 
  end if 
  if filename2<>"" then
  filename2=replace(filename2,",","_")
	filedown2 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_2"
	filetype2 = strReverse(left(strReverse(filename2),instr(strReverse(filename2),".")))
    uploadFile2.Save2File filepath&"\"&filedown2&filetype2
	Path = Server.MapPath("../upload")&"\"&filedown2&filetype2 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown2&filetype2 
  end if 
  if filename3<>"" then
  filename3=replace(filename3,",","_")
	filedown3 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_3"
	filetype3 = strReverse(left(strReverse(filename3),instr(strReverse(filename3),".")))
    uploadFile3.Save2File filepath&"\"&filedown3&filetype3
	Path = Server.MapPath("../upload")&"\"&filedown3&filetype3 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown3&filetype3 
  end if 
  if filename4<>"" then
  filename4=replace(filename4,",","_")
	filedown4 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_4"
	filetype4 = strReverse(left(strReverse(filename4),instr(strReverse(filename4),".")))
    uploadFile4.Save2File filepath&"\"&filedown4&filetype4
	Path = Server.MapPath("../upload")&"\"&filedown4&filetype4 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown4&filetype4 
  end if 
  if filename5<>"" then
  filename5=replace(filename5,",","_")
	filedown5 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_5"
	filetype5 = strReverse(left(strReverse(filename5),instr(strReverse(filename5),".")))
    uploadFile5.Save2File filepath&"\"&filedown5&filetype5
	Path = Server.MapPath("../upload")&"\"&filedown5&filetype5 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 150
	Jpeg.Height = 100
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown5&filetype5 
  end if 


	
	filename=filenamex&","&filename1&","&filename2&","&filename3&","&filename4&","&filename5&","
	filedown=filedownx&","&filedown1&","&filedown2&","&filedown3&","&filedown4&","&filedown5&","
	filetype=filetypex&","&filetype1&","&filetype2&","&filetype3&","&filetype4&","&filetype5&","
	
	
	a=split(filename,",")
	b=split(filedown,",")
	c=split(filetype,",")
	 
	 
	for i=LBound(a) to UBound(a)
		if a(i)<>"" then
		xx=xx&a(i)&","
		end if
	next
	for j=LBound(b) to UBound(b)
		if b(j)<>"" then
		yy=yy&b(j)&","
		end if
	next
	for k=LBound(c) to UBound(c)
		if c(k)<>"" then
		zz=zz&c(k)&","
		end if
	next

	set rs=server.CreateObject("adodb.recordset")
	sql="select * from project where id="&upload.form("id")
	rs.open sql,conn,1,3
		if xx<>"" then
			rs("filename4")=rs("filename4")&xx
			rs("filedown4")=rs("filedown4")&yy
			rs("filetype4")=rs("filetype4")&zz
		end if
	rs.update
	rs.close
	set rs=nothing
	
	
	response.Redirect("success_main.asp?mm=2")

end if







if upload.form("act")= "cpsc" then

 set uploadFilex = upload.file("filex")
  
  set uploadFile1 = upload.file("file1")
  set uploadFile2 = upload.file("file2")
  set uploadFile3 = upload.file("file3")
  set uploadFile4 = upload.file("file4")
  set uploadFile5 = upload.file("file5")
  filenamex = uploadFilex.filename
  
  filename1 = uploadFile1.filename
  filename2 = uploadFile2.filename
  filename3 = uploadFile3.filename
  filename4 = uploadFile4.filename
  filename5 = uploadFile5.filename
  filepath = server.MapPath("../upload")
  filepath_sm = server.MapPath("../upload/images")


Set Jpeg = Server.CreateObject("Persits.Jpeg") '调用组件 




  if filenamex<>"" then
  filenamex=replace(filenamex,",","_")
	filedownx = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_x"
	filetypex = strReverse(left(strReverse(filenamex),instr(strReverse(filenamex),".")))
    uploadFilex.Save2File filepath&"\"&filedownx&filetypex
	
	Path = Server.MapPath("../upload")&"\"&filedownx&filetypex '待处理图片路径 
	Jpeg.Open Path '打开图片 
	Jpeg.Width = 50
	Jpeg.Height = 40
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedownx&filetypex 
  end if 
  if filename1<>"" then
  filename1=replace(filename1,",","_")
	filedown1 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_1"
	filetype1 = strReverse(left(strReverse(filename1),instr(strReverse(filename1),".")))
    uploadFile1.Save2File filepath&"\"&filedown1&filetype1
	Path = Server.MapPath("../upload")&"\"&filedown1&filetype1 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 50
	Jpeg.Height = 40
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown1&filetype1 
  end if 
  if filename2<>"" then
  filename2=replace(filename2,",","_")
	filedown2 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_2"
	filetype2 = strReverse(left(strReverse(filename2),instr(strReverse(filename2),".")))
    uploadFile2.Save2File filepath&"\"&filedown2&filetype2
	Path = Server.MapPath("../upload")&"\"&filedown2&filetype2 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 50
	Jpeg.Height = 40
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown2&filetype2 
  end if 
  if filename3<>"" then
  filename3=replace(filename3,",","_")
	filedown3 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_3"
	filetype3 = strReverse(left(strReverse(filename3),instr(strReverse(filename3),".")))
    uploadFile3.Save2File filepath&"\"&filedown3&filetype3
	Path = Server.MapPath("../upload")&"\"&filedown3&filetype3 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 50
	Jpeg.Height = 40
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown3&filetype3 
  end if 
  if filename4<>"" then
  filename4=replace(filename4,",","_")
	filedown4 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_4"
	filetype4 = strReverse(left(strReverse(filename4),instr(strReverse(filename4),".")))
    uploadFile4.Save2File filepath&"\"&filedown4&filetype4
	Path = Server.MapPath("../upload")&"\"&filedown4&filetype4 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 50
	Jpeg.Height = 40
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown4&filetype4 
  end if 
  if filename5<>"" then
  filename5=replace(filename5,",","_")
	filedown5 = lrr&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_5"
	filetype5 = strReverse(left(strReverse(filename5),instr(strReverse(filename5),".")))
    uploadFile5.Save2File filepath&"\"&filedown5&filetype5
	Path = Server.MapPath("../upload")&"\"&filedown5&filetype5 '待处理图片路径 
	Jpeg.Open Path '打开图片 
	'高与宽为原图片的1/2 
	Jpeg.Width = 50
	Jpeg.Height = 40
	'保存图片 
	Jpeg.Save Server.MapPath("../upload/images")&"\"&filedown5&filetype5 
  end if 


	
	filename=filenamex&","&filename1&","&filename2&","&filename3&","&filename4&","&filename5&","
	filedown=filedownx&","&filedown1&","&filedown2&","&filedown3&","&filedown4&","&filedown5&","
	filetype=filetypex&","&filetype1&","&filetype2&","&filetype3&","&filetype4&","&filetype5&","
	
	
	a=split(filename,",")
	b=split(filedown,",")
	c=split(filetype,",")
	 
	 piclx=upload.form("piclx")
	for i=LBound(a) to UBound(a)
		if a(i)<>"" then
		xx=xx&piclx&a(i)&","
		end if
	next
	for j=LBound(b) to UBound(b)
		if b(j)<>"" then
		yy=yy&b(j)&","
		end if
	next
	for k=LBound(c) to UBound(c)
		if c(k)<>"" then
		zz=zz&c(k)&","
		end if
	next
	
	
	
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from project where id="&upload.form("id")
	rs.open sql,conn,1,3
		if xx<>"" then
			rs("filename6")=rs("filename6")&xx
			rs("filedown6")=rs("filedown6")&yy
			rs("filetype6")=rs("filetype6")&zz
		end if
	rs.update
	rs.close
	set rs=nothing
	
	
	response.Redirect("success_main.asp?mm=2")

end if

%>