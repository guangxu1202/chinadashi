<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<!-- #include File="inc/Upload_Class.asp" -->
<%
Echo " <link href=""skins/teams/bbs.css"" rel=""stylesheet""> "
Echo " <body topmargin=""0""  rightmargin=""0""  leftmargin=""0"" Style=""background-color:#FFFFFF"">"
Call TestUser()
Dim CheckFolder
CheckFolder = "Images/Upfile/"
Dim fid
fid = Request("fid")
Select Case Request("action")
	Case "ups"
		Call Upfilestar
	Case "upface"
		Call UpFace
	Case Else
		Call Main()
End Select

Sub UpFace
	If CID(team.Forum_setting(70)) = 999 Then 
		team.error2 " 论坛管理了上传功能，请勿上传附件。"
	End If
	If team.TK_UserID = 0 Then
		team.error2 "您没有登陆论坛"
	End if
	If CID(team.Group_Browse(5)) = 0 Then 
		team.error2 " 您没有在本论坛上传头像附件的权限。"
	End If
	If team.TK_UserID>0 Then
		UpUserFace()	'删除旧的头像文件
	End If
	Dim Upload,FilePath,FormName,File,F_FileName
	Dim UserID
	UserID = ""
	If team.TK_UserID>0 Then UserID = Cint(team.TK_UserID)&"_"
	FilePath = "Images/UpFace/"
	Set Upload = New UpFile_Cls
		Upload.UploadType			= CID(team.Forum_setting(70))		'设置上传组件类型
		Upload.UploadPath			= FilePath							'设置上传路径
		Upload.InceptFileType		= "gif,jpg,jpeg,png"				'设置上传文件限制
		Upload.MaxSize				= CID(team.Forum_setting(72))		'单位 KB
		Upload.ChkSessionName		= "uploadcode"						'防止重复提交，SESSION名与提交的表单要一致。
		Upload.RName				= UserID
		Upload.SaveUpFile		'执行上传
		If Upload.ErrCodes<>0 Then
			team.Error2 "错误："& Upload.Description
			Exit Sub
		End If
		If Upload.Count > 0 Then
			For Each FormName In Upload.UploadFiles
				Set File = Upload.UploadFiles(FormName)
					F_FileName = FilePath & File.FileName
					Echo " <script> "
					Echo " parent.document.getElementById('urlavatar').value='" &F_FileName& "'; "
					Echo " parent.document.getElementById('showface').src='"&F_FileName&"'; "
					Echo " parent.document.getElementById('statusid').innerHTML='图片上传成功';"
					Echo " parent.document.getElementById('statusid').style.color='red';"
					Echo "</script> "
					Session("upface")="done"
					Echo  "[ <a href=""#"" onclick=""history.go(-1)"">重新上传?</a> ]"
					TEAm.Execute("Update ["&IsForum&"User] Set UserFace='"&F_FileName&"' Where UserName='"&TK_UserName&"'")	
					Session(CacheName&"_UserLogin")=""
				Set File = Nothing
			Next
		Else
			team.Error2 " 请正确选择要上传的文件。"
			Exit Sub
		End If
	Set Upload = Nothing
End Sub

Sub UpUserFace()
	on Error Resume Next
	Dim objFSO,OldUserFace
	OldUserFace = Server.MapPath("Images/UpFace/"&CID(team.TK_UserID)&"_")&"*.*"
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(OldUserFace) Then
		objFSO.DeleteFile OldUserFace
		If Err<>0 Then Err.Clear
	End If
	Set objFSO = Nothing
End Sub

Sub Main
	Dim PostRanNum
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
	Echo " <script>if(top==self)document.location='default.asp';</script>"
	Echo " <table border=""0""  cellspacing=""0"" cellpadding=""0"" width=""100%""> "
	Echo " <tr><td class=a4>"
	Echo " <form name=""form"" method=""post"" action=""?action=upface"" enctype=""multipart/form-data"">"
	Echo " <INPUT TYPE=""hidden"" name=""uploadcode"" value="""&PostRanNum&""">"
	Echo " <input type=""file"" name=""upfile"" size=""25"">"
	Echo " <input type=""submit"" name=""Submit"" value=""上传"">"
	Echo " </form></body> "
End Sub

Sub Upfilestar
	Dim FilePath,UpNum,Previewpath
	Dim ChildFilePath,DrawInfo,Upload
	Server.ScriptTimeOut=999999		'过期时间
	team.ChkPost()
	If CID(team.Forum_setting(70)) = 999 Then 
		UpErrors " 论坛管理了上传功能，请勿上传附件。"
	End If
	If team.TK_UserID = 0 Then
		UpErrors "您没有登陆论坛"
	End If
	If CID(team.Group_Browse(25)) = 0 Then 
		UpErrors " 您没有在本论坛上传附件的权限。"
	End If
	UpNum=request.cookies("Class")("upNum")
	If UpNum = "" or Not Isnumeric(UpNum) then 
		UpNum=0
	Else
		UpNum=Clng(UpNum)
	End If
	If UpNum > Int(team.Group_Browse(28)) Then
		UpErrors "您已经超过了每日最大上传数"
	End If
	'上传目录
	FilePath = CreatePath(CheckFolder)
	'不带系统上传目录的下级目录路径
	ChildFilePath = Replace(FilePath,CheckFolder,"")
	'预览图片目录路径
	Previewpath = "PreviewImage/"
	Previewpath = CreatePath(Previewpath)
	If team.Forum_setting(75)=1 Then
		DrawInfo = team.Forum_setting(76)
	ElseIf team.Forum_setting(75)=2 Then
		DrawInfo = team.Forum_setting(79)
	Else
		DrawInfo = ""
	End If
	Dim MyCoundUps
	If Int(team.Group_Browse(27))  > Int(team.Forum_setting(71)) Then
		MyCoundUps = Int(team.Forum_setting(71)) 
	Else
		MyCoundUps = Int(team.Group_Browse(27)) 
	End If
	Dim UpTypes
	If Len(team.Group_Browse(29))>2 Or InStr(team.Group_Browse(29),",")>0 Then
		UpTypes = team.Group_Browse(29)
	Else
		UpTypes = ReplaceStr(team.Forum_setting(73),"|",",")
	End if
	Set Upload = New UpFile_Cls
	Upload.UploadType			= CID(team.Forum_setting(70))		'设置上传组件类型
	Upload.UploadPath			= FilePath							'设置上传路径
	Upload.InceptFileType		= UpTypes							'设置上传文件限制
	Upload.MaxSize				= MyCoundUps						'单位 KB
	Upload.InceptMaxFile		= CID(team.Group_Browse(26))		'每次上传文件个数上限
	Upload.ChkSessionName		= "uploadcode"						'防止重复提交，SESSION名与提交的表单要一致。
	Upload.ChkTextName			= "trims"							'表单内容。
	'============预览图片设置==================================================================================
	Upload.PreviewType			= CID(team.Forum_setting(74))		'设置预览图片组件类型
	Upload.DrawImageWidth		= CID(team.Forum_setting(77))		'设置水印图片或文字区域宽度
	Upload.DrawImageHeight		= CID(team.Forum_setting(35))		'设置水印图片或文字区域高度
	Upload.DrawGraph			= team.Forum_setting(84)			'设置水印透明度
	Upload.DrawFontColor		= team.Forum_setting(81)			'设置水印文字颜色
	Upload.DrawFontFamily		= team.Forum_setting(82)			'设置水印文字字体格式
	Upload.DrawFontSize			= team.Forum_setting(80)			'设置水印文字字体大小
	Upload.DrawFontBold			= "1"								'设置水印文字是否粗体 1:粗体 0:无  
	Upload.DrawInfo				= DrawInfo							'设置水印文字信息或图片信息
	Upload.DrawType				= CID(team.Forum_setting(75))		'0=不加载水印 ，1=加载水印文字，2=加载水印图片
	Upload.DrawXYType			= CID(team.Forum_setting(83))		'"0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
	Upload.DrawSizeType			= CID(team.Forum_setting(78))		'"0"=固定缩小，"1"=等比例缩小
	If team.Forum_setting(9)<>"" Then
		Upload.TransitionColor	= team.Forum_setting(9)				'透明度颜色设置
	End If
	'执行上传
	Upload.SaveUpFile
	Call Suc_upload(Upload.Count,UpNum)
	If Upload.ErrCodes<>0 Then
		UpErrors "错误信息："& Upload.Description
	End If
	Dim FormName,File,F_FileName,F_Viewname
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
				F_FileName = FilePath & File.FileName
				'创建预览及水印图片
				If Upload.PreviewType<>999 and File.FileType=1 then
					F_Viewname = Previewpath & "pre" & ReplaceStr(File.FileName,File.FileExt,"") & "jpg"
					'创建预览图片:Call CreateView(原始文件的路径,预览文件名及路径,原文件后缀)
					Upload.CreateView F_FileName,F_Viewname,File.FileExt
				End If
				UploadSave F_FileName,ChildFilePath&File.FileName,File.FileExt,F_Viewname,File.FileSize,File.FileType,Upload.UpTextName
			Set File = Nothing
		Next
	Else
		UpErrors "您的上传附件数为 "& Upload.Count &"，请重新上传。" 
	End If
	Set Upload = Nothing
	UpdateUserpostExc()
End Sub 

Sub UpdateUserpostExc()
	'用户积分部分
	Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u
	Dim UserPostID,My_ExtSort
	If Not team.UserLoginED Then  Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	For U=0 to Ubound(ExtCredits)
		ExtSort=Split(ExtCredits(U),",")
		MustSort=Split(MustOpen(U),",")
		If ExtSort(3)=1 Then
			If U = 0 Then
				UExt = UExt &"Extcredits0=Extcredits0+"&MustSort(3)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"+"&MustSort(3)&""
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where ID = "& team.TK_UserID)
End Sub

'保存上传数据并返回附件ID
Sub UploadSave(FileName,ChildFileName,FileExt,ViewName,FileSize,F_Type,Trims)
	Dim Fileid,Rs,UpFileID
	ChildFileName = team.Checkstr(ChildFileName)
	team.execute("Insert into UpFile (FID,UserName,Filename,Types,Lasttime,FileSize) values ("&fid&",'"&TK_UserName&"','"&ChildFileName&"','"&FileExt&"',"&SqlNowString&","&FileSize&")")
	Set Rs = team.execute("Select Max(Fileid) from ["&IsForum&"UpFile] ")
	If Not Rs.Eof Then
		Fileid = RS(0)
		UpFileID = Fileid & ","
	End if
	Rs.Close:Set Rs=Nothing
	Echo " <script language=""javascript"">"
	Echo "	var oEditor =  window.parent.FCKeditorAPI.GetInstance('message') ; "
	Echo "	if(oEditor.EditMode == window.parent.FCK_EDITMODE_WYSIWYG){"
	Echo "		oEditor.InsertHtml('[UPLOAD="&FileExt&"]ShowFile.asp?ID="&Fileid&"[/UPLOAD]<br>附件尺寸："& Round(FileSize/1024,2) &" KB<br>');	 "
	Echo " }else{"
	Echo "		parent.message.value+='[UPLOAD="&FileExt&"]ShowFile.asp?ID="&Fileid&"[/UPLOAD]\n附件尺寸："& Round(FileSize/1024,2) &" KB\n';	 "
	Echo " };"
	Echo "	parent.document.getElementById('updiv').style.display='';					"
	Echo "	parent.document.getElementById('showupdiv').style.display='block';					"
	Echo "	parent.document.getElementById('showupfileids').style.color='red';				"
	Echo "	parent.document.getElementById('showupfileids').innerHTML='上传附件成功。';"
	Echo "	parent.upfile.upfilesubmit.disabled = true; "
	Echo "	parent.postform.upfileid.value+='"&UpFileID&"'; "
	Echo " </script>	"
End Sub

Sub Suc_upload(UpCount,upNum)
	Dim u
	upNum = upNum + UpCount
	Response.Cookies("class")("upNum") = upNum
	If InStr(team.UserUp,"|")>0 Then
		u = Split(team.UserUp,"|")(0) + Int(UpCount)
	Else
		u = 0 + Int(UpCount)
	End If
	team.Execute("Update [user] Set UserUp='"&u&"|"&Now()&"' Where ID=" & team.TK_UserID)
End Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'按月份自动明名上传文件夹,需要ＦＳＯ组件支持。
Private Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	'以年月创建上传文件夹，格式：2003－8
	uploadpath = year(now) & "-" & month(now)
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
		End If
		If Err.Number = 0 Then
			CreatePath = PathValue & uploadpath & "/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function
Sub UpErrors(a)
	Echo " <script language=""javascript"">"
	'Echo "	parent.document.getElementById('updiv').style.display='none';					"
	Echo "	parent.document.getElementById('showupdiv').style.display='block';					"
	Echo "	parent.document.getElementById('showupfileids').style.color='red';				"
	Echo "	parent.document.getElementById('showupfileids').innerHTML='"&a&"';"
	'Echo "	parent.upfile.upfilesubmit.disabled = true; "
	Echo " </script>"
	Response.End
End Sub
%>