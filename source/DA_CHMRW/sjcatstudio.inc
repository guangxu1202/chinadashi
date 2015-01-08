<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>
'===========================================================================================================
'
'					sjCatSoft ASP 系 列 V2.5
'
'					作者:三角猫@sjCatStudio
'					Email:sjcatsoft@163.com
'						sjcatsoft@yahoo.com.cn
'					MSN: sjcatsoft@hotmail.com
'					版权所有：三角猫
'					本组件可自由传播，但不可用于商业用途，否则一切法律责任由使用者承担
'					感谢：本组件中文件上传的部分采用5x_soft的无组件上传，本人略做修改，转载时请保留此信息
'								
'					声明：你可以随意更改本组件，对其进行优化和修正，但请将修改后的发给我一份参考，谢谢
'============================================================================================================
	dim Data_sjCat

	Class sjCat_Upload
  
	Dim objForm,objFile

	Public function Form(strForm)
		strForm=lcase(strForm)
		if not objForm.exists(strForm) then
			Form=""
		else
			Form=objForm(strForm)
		end if
	End function

	Public function File(strFile)
		strFile=lcase(strFile)
		if not objFile.exists(strFile) then
			set File=new FileInfo
		else
			set File=objFile(strFile)
		end if
	End function


	Private Sub Class_Initialize 
		dim RequestData,sStart,vbEnter,sInfo,iInfoStart,iInfoEnd,tStream,iStart,theFile
		dim iFileSize,sFileType,sFormValue,sFileName
		dim iFindStart,iFindEnd
		dim iFormStart,iFormEnd,sFormName
		set objForm=Server.CreateObject("Scripting.Dictionary")
		set objFile=Server.CreateObject("Scripting.Dictionary")
		if Request.TotalBytes<1 then Exit Sub
		set tStream = Server.CreateObject("adodb.stream")
		set Data_sjCat = Server.CreateObject("adodb.stream")
		Data_sjCat.Type = 1
		Data_sjCat.Mode =3
		Data_sjCat.Open
		Data_sjCat.Write  Request.BinaryRead(Request.TotalBytes)
		Data_sjCat.Position=0
		RequestData =Data_sjCat.Read 

		iFormStart = 1
		iFormEnd = LenB(RequestData)
		vbEnter = chrB(13) & chrB(10)
		sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbEnter)-1)
		iStart = LenB (sStart)
		iFormStart=iFormStart+iStart+1
		while (iFormStart + 10) < iFormEnd 
			iInfoEnd = InStrB(iFormStart,RequestData,vbEnter & vbEnter)+3
			tStream.Type = 1
			tStream.Mode =3
			tStream.Open
			Data_sjCat.Position = iFormStart
			Data_sjCat.CopyTo tStream,iInfoEnd-iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.Charset ="GB2312"
			sInfo = tStream.ReadText
			tStream.Close
	'取得表单项目名称
			iFormStart = InStrB(iInfoEnd,RequestData,sStart)
			iFindStart = InStr(22,sInfo,"name=""",1)+6
			iFindEnd = InStr(iFindStart,sInfo,"""",1)
			sFormName = lcase(Mid (sinfo,iFindStart,iFindEnd-iFindStart))
	'如果是文件
			if InStr (45,sInfo,"filename=""",1) > 0 then
				set theFile=new FileInfo
		'取得文件名
				iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
				iFindEnd = InStr(iFindStart,sInfo,"""",1)
				sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
				theFile.FileName=getFileName(sFileName)
		'取得文件类型
				iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
				iFindEnd = InStr(iFindStart,sInfo,vbCr)
				theFile.FileType =Mid (sinfo,iFindStart,iFindEnd-iFindStart)
				theFile.FileStart =iInfoEnd
				theFile.FileSize = iFormStart -iInfoEnd -3
				theFile.FormName=sFormName
				if not objFile.Exists(sFormName) then
					objFile.add sFormName,theFile
				end if
			else
	'如果是表单项目
				tStream.Type =1
				tStream.Mode =3
				tStream.Open
				Data_sjCat.Position = iInfoEnd 
				Data_sjCat.CopyTo tStream,iFormStart-iInfoEnd-3
				tStream.Position = 0
				tStream.Type = 2
				tStream.Charset ="gb2312"
				sFormValue = tStream.ReadText 
				tStream.Close
				if objForm.Exists(sFormName) then
					objForm(sFormName)=objForm(sFormName)&", "&sFormValue		  
				else
					objForm.Add sFormName,sFormValue
				end if
			end if
			iFormStart=iFormStart+iStart+1
		wend
		RequestData=""
		set tStream =nothing
	End Sub

	Private Sub Class_Terminate  
		If Request.TotalBytes > 0 then
			objForm.RemoveAll
			objFile.RemoveAll
			set objForm = nothing
			set objFile = nothing
			Data_sjCat.Close
			set Data_sjCat = nothing
		End if
	End Sub
   
	Private Sub Kill(fullpath)
		Dim Fso
		Set Fso = Server.CreateObject("Scripting.FileSystemObject")
		If Fso.FileExists(fullpath) Then
			Fso.DeleteFile (fullpath)
		End if
		Set Fso = nothing
	End Sub
	
	Private function GetFilePath(FullPath)
		If FullPath <> "" Then
			GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
		Else
			GetFilePath = ""
		End If
	End  function
 
	Private function GetFileName(FullPath)
		If FullPath <> "" Then
			GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
		Else
			GetFileName = ""
		End If
	End  function

	Public Function Up2DB(ActiveAdoCon,sqlStatement,formFieldnames,dbFieldnames,Filenames,BlobFields,oType,fname,path)
		Dim Rs,adS,i,formArray,FieldArray,fileArray,BlobArray
		Dim theFile
		Dim temp_Fname
		Up2DB = false
		If Trim(formFieldNames) <> "" Then
			formArray = Split(formFieldNames, "|")
			fieldArray = Split(dbFieldNames, "|")
		End If
		fileArray = Split(FileNames, "|")
		BlobArray = Split(BlobFields, "|")
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Set adS = Server.CreateObject("ADODB.Stream")
		Rs.Open sqlStatement,ActiveAdoCon,3,2
		If oType =0 Then
			Rs.AddNew
		End if
		If IsArray(formArray) Then
                    For i = LBound(formArray) To UBound(formArray)
                            Rs.Fields(fieldArray(i)).Value = objForm(formArray(i))
                    Next
		End If
		    rs("fname")=fname
		adS.Mode = 3
		adS.Type = 1
		For i = LBound(fileArray) To UBound(fileArray)
                        adS.Open
                        Set theFile = File(fileArray(i))
                        temp_Fname = "sjCatStudio_" & theFile.FileName
                        Data_sjCat.Position = theFile.FileStart
                        Data_sjCat.CopyTo adS, theFile.FileSize
                        adS.Close
                        adS.Open
                        adS.LoadFromFile path
                        call Kill(path)
                        Rs.Fields(BlobArray(i)).Value = adS.Read
			
                        adS.Close
		Next
		Rs.Update
		Rs.Close
		Set Rs = nothing
		Set adS = nothing
		
		Up2DB = True
		
	End Function
	
       
	
	End Class

'---------------------------------------------------------------------------------------

	Class FileInfo
		dim FormName,FileName,FileSize,FileType,FileStart
		Private Sub Class_Initialize 
			FileName = ""
			FileSize = 0
			FileStart= 0
			FormName = ""
			FileType = ""
		End Sub
  
		Public function Save2File(FullPath)
			dim dr,i
			Save2File = false
			if trim(fullpath)="" or FileStart=0 or FileName="" or right(fullpath,1)="/" then exit function
			set dr = Server.CreateObject("Adodb.Stream")
			dr.Mode=3
			dr.Type=1
			dr.Open
			Data_sjCat.position=FileStart
			Data_sjCat.copyto dr,FileSize
			dr.SaveToFile FullPath,2
			dr.Close
			set dr=nothing 
			Save2File = true
		End function


		Public Function Save2DB(ActiveAdoCon,sqlStatement,BlobFieldName,oType)
			Dim adS,Rs
			Dim temp_Fname
			If  FileStart=0 Then Exit Function
			Save2DB = False
			temp_Fname = "sjcatstudio_" & FileName
			Set Rs = Server.CreateObject("ADODB.Recordset")
			Set adS = Server.CreateObject("ADODB.Stream")
			adS.Mode = 3
			adS.Type = 1
			adS.Open
			Data_sjCat.Position = FileStart
			Data_sjCat.CopyTo adS,FileSize
			adS.SaveToFile temp_Fname,2
			adS.Close
			adS.Open
			adS.LoadFromFile temp_Fname
			Call Kill(temp_Fname)
			Rs.Open sqlStatement,ActiveAdoCon,3,2
			If oType = 0 Then
				Rs.AddNew
			End if
			Rs.Fields(BlobFieldName).Value = adS.Read
			Rs.Update
			Rs.Close
			adS.Close
			Set adS = nothing
			Set Rs = nothing
			Save2DB = True
		End Function
  
	End Class
'-------------------------------------------------------------------------------------
	Class sjcat_DB2Page
		
		Public Sub Show(ActiveDBCon,sqlStatement)
			Dim Rs,FSize
			Set Rs = Server.CreateObject("ADODB.Recordset")
			Rs.Open sqlStatement,ActiveDBCon,1,1,1
			FSize = Rs(0).ActualSize
			Response.Buffer = true
			Response.Clear
			Response.ContentType = "image/*"
			Response.BinaryWrite Rs(0).GetChunk(FSize)
			Rs.Close
			Set Rs = nothing
		End Sub
	End Class
'------------------------------------------------------------------------------------

	Class sjcat_DownLoad
		Public Sub DownLoadFromFile(FilePath,FileName)
			Dim adS
			If Trim(FilePath) = "" or Right(FilePath,1) = "\" then Exit Sub
			Set adS = Server.CreateObject("ADODB.Stream")
			With adS
				.Mode = 3
				.Type = 1
				.Open
				.LoadFromFile FilePath
			End With
			Response.Buffer = true
			Response.Clear
			Response.AddHeader "Content-Disposition","attachment;filename=" & FileName
			Response.AddHeader "Content-Length",adS.Size
                        Response.CharSet = "UTF-8"
			Response.ContentType = "Application/Octet-Stream"
			Response.BinaryWrite adS.Read
			Response.Flush
			adS.Close
			Set adS = nothing
		End Sub

		Public Sub DownLoadFromDB(ActiveDBCon,sqlStatement)
			Dim fSize
			Dim Rs		
			Set Rs = Server.CreateObject("ADODB.Recordset")
			Rs.Open sqlStatement,ActiveDBCon,1,1,1
			fSize = Rs(0).ActualSize
			Response.Buffer = true
			Response.Clear
			Response.AddHeader "Content-Disposition","attachment;filename=" & Rs(1)
			Response.AddHeader "Content-Length",fSize
                        Response.CharSet = "UTF-8"
			Response.ContentType = "Application/Octet-Stream"
			Response.BinaryWrite Rs(0).GetChunk(fSize)
			Response.Flush
			Rs.Close
			Set Rs = nothing
		End Sub

	End Class

'----------------------------------------------------------------------------------------------------

	Class sjcat_File2DB
		Dim Files
		Public Sub Execute(ActiveDBCon,sqlStatement,FileString,FieldString,oType)
			Dim Rs,adS,i
			Dim fileArray,FieldArray
			If Trim(fileString) = "" then Exit Sub
			fileArray = Split(FileString,"|")
			FieldArray = Split(FieldString,"|")
			Set Rs = Server.CreateObject("ADODB.Recordset")
			Rs.Open sqlStatement,ActiveDBCon,3,2
			If oType = 0 then
				Rs.AddNew
			End If
			Set adS = Server.CreateObject("ADODB.Stream")
			adS.Mode = 3
			adS.Type = 1
			Set Files = Server.CreateObject("Scripting.Dictionary")
			For i = Lbound(fileArray) to UBound(FileArray)
				adS.Open
				adS.LoadFromFile fileArray(i)
				Rs.Fields(FieldArray(i)).Value = adS.Read
				Files.Add fieldArray(i),fileArray(i)
				adS.Close
			Next
			Rs.Update
			Rs.Close
			Set adS = nothing
			Set Rs = nothing
		End Sub

		Public Function File(index)
			index = LCase(index)
			IF Files.Exists(index) then
				File = Files(index)
			Else
				File = ""
			End If
		End Function

		Private Sub Class_Terminate
			Set Files = nothing
		End Sub

	End Class
</SCRIPT>