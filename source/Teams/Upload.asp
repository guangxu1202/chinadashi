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
		team.error2 " ��̳�������ϴ����ܣ������ϴ�������"
	End If
	If team.TK_UserID = 0 Then
		team.error2 "��û�е�½��̳"
	End if
	If CID(team.Group_Browse(5)) = 0 Then 
		team.error2 " ��û���ڱ���̳�ϴ�ͷ�񸽼���Ȩ�ޡ�"
	End If
	If team.TK_UserID>0 Then
		UpUserFace()	'ɾ���ɵ�ͷ���ļ�
	End If
	Dim Upload,FilePath,FormName,File,F_FileName
	Dim UserID
	UserID = ""
	If team.TK_UserID>0 Then UserID = Cint(team.TK_UserID)&"_"
	FilePath = "Images/UpFace/"
	Set Upload = New UpFile_Cls
		Upload.UploadType			= CID(team.Forum_setting(70))		'�����ϴ��������
		Upload.UploadPath			= FilePath							'�����ϴ�·��
		Upload.InceptFileType		= "gif,jpg,jpeg,png"				'�����ϴ��ļ�����
		Upload.MaxSize				= CID(team.Forum_setting(72))		'��λ KB
		Upload.ChkSessionName		= "uploadcode"						'��ֹ�ظ��ύ��SESSION�����ύ�ı�Ҫһ�¡�
		Upload.RName				= UserID
		Upload.SaveUpFile		'ִ���ϴ�
		If Upload.ErrCodes<>0 Then
			team.Error2 "����"& Upload.Description
			Exit Sub
		End If
		If Upload.Count > 0 Then
			For Each FormName In Upload.UploadFiles
				Set File = Upload.UploadFiles(FormName)
					F_FileName = FilePath & File.FileName
					Echo " <script> "
					Echo " parent.document.getElementById('urlavatar').value='" &F_FileName& "'; "
					Echo " parent.document.getElementById('showface').src='"&F_FileName&"'; "
					Echo " parent.document.getElementById('statusid').innerHTML='ͼƬ�ϴ��ɹ�';"
					Echo " parent.document.getElementById('statusid').style.color='red';"
					Echo "</script> "
					Session("upface")="done"
					Echo  "[ <a href=""#"" onclick=""history.go(-1)"">�����ϴ�?</a> ]"
					TEAm.Execute("Update ["&IsForum&"User] Set UserFace='"&F_FileName&"' Where UserName='"&TK_UserName&"'")	
					Session(CacheName&"_UserLogin")=""
				Set File = Nothing
			Next
		Else
			team.Error2 " ����ȷѡ��Ҫ�ϴ����ļ���"
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
	Echo " <input type=""submit"" name=""Submit"" value=""�ϴ�"">"
	Echo " </form></body> "
End Sub

Sub Upfilestar
	Dim FilePath,UpNum,Previewpath
	Dim ChildFilePath,DrawInfo,Upload
	Server.ScriptTimeOut=999999		'����ʱ��
	team.ChkPost()
	If CID(team.Forum_setting(70)) = 999 Then 
		UpErrors " ��̳�������ϴ����ܣ������ϴ�������"
	End If
	If team.TK_UserID = 0 Then
		UpErrors "��û�е�½��̳"
	End If
	If CID(team.Group_Browse(25)) = 0 Then 
		UpErrors " ��û���ڱ���̳�ϴ�������Ȩ�ޡ�"
	End If
	UpNum=request.cookies("Class")("upNum")
	If UpNum = "" or Not Isnumeric(UpNum) then 
		UpNum=0
	Else
		UpNum=Clng(UpNum)
	End If
	If UpNum > Int(team.Group_Browse(28)) Then
		UpErrors "���Ѿ�������ÿ������ϴ���"
	End If
	'�ϴ�Ŀ¼
	FilePath = CreatePath(CheckFolder)
	'����ϵͳ�ϴ�Ŀ¼���¼�Ŀ¼·��
	ChildFilePath = Replace(FilePath,CheckFolder,"")
	'Ԥ��ͼƬĿ¼·��
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
	Upload.UploadType			= CID(team.Forum_setting(70))		'�����ϴ��������
	Upload.UploadPath			= FilePath							'�����ϴ�·��
	Upload.InceptFileType		= UpTypes							'�����ϴ��ļ�����
	Upload.MaxSize				= MyCoundUps						'��λ KB
	Upload.InceptMaxFile		= CID(team.Group_Browse(26))		'ÿ���ϴ��ļ���������
	Upload.ChkSessionName		= "uploadcode"						'��ֹ�ظ��ύ��SESSION�����ύ�ı�Ҫһ�¡�
	Upload.ChkTextName			= "trims"							'�����ݡ�
	'============Ԥ��ͼƬ����==================================================================================
	Upload.PreviewType			= CID(team.Forum_setting(74))		'����Ԥ��ͼƬ�������
	Upload.DrawImageWidth		= CID(team.Forum_setting(77))		'����ˮӡͼƬ������������
	Upload.DrawImageHeight		= CID(team.Forum_setting(35))		'����ˮӡͼƬ����������߶�
	Upload.DrawGraph			= team.Forum_setting(84)			'����ˮӡ͸����
	Upload.DrawFontColor		= team.Forum_setting(81)			'����ˮӡ������ɫ
	Upload.DrawFontFamily		= team.Forum_setting(82)			'����ˮӡ���������ʽ
	Upload.DrawFontSize			= team.Forum_setting(80)			'����ˮӡ���������С
	Upload.DrawFontBold			= "1"								'����ˮӡ�����Ƿ���� 1:���� 0:��  
	Upload.DrawInfo				= DrawInfo							'����ˮӡ������Ϣ��ͼƬ��Ϣ
	Upload.DrawType				= CID(team.Forum_setting(75))		'0=������ˮӡ ��1=����ˮӡ���֣�2=����ˮӡͼƬ
	Upload.DrawXYType			= CID(team.Forum_setting(83))		'"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
	Upload.DrawSizeType			= CID(team.Forum_setting(78))		'"0"=�̶���С��"1"=�ȱ�����С
	If team.Forum_setting(9)<>"" Then
		Upload.TransitionColor	= team.Forum_setting(9)				'͸������ɫ����
	End If
	'ִ���ϴ�
	Upload.SaveUpFile
	Call Suc_upload(Upload.Count,UpNum)
	If Upload.ErrCodes<>0 Then
		UpErrors "������Ϣ��"& Upload.Description
	End If
	Dim FormName,File,F_FileName,F_Viewname
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
				F_FileName = FilePath & File.FileName
				'����Ԥ����ˮӡͼƬ
				If Upload.PreviewType<>999 and File.FileType=1 then
					F_Viewname = Previewpath & "pre" & ReplaceStr(File.FileName,File.FileExt,"") & "jpg"
					'����Ԥ��ͼƬ:Call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��,ԭ�ļ���׺)
					Upload.CreateView F_FileName,F_Viewname,File.FileExt
				End If
				UploadSave F_FileName,ChildFilePath&File.FileName,File.FileExt,F_Viewname,File.FileSize,File.FileType,Upload.UpTextName
			Set File = Nothing
		Next
	Else
		UpErrors "�����ϴ�������Ϊ "& Upload.Count &"���������ϴ���" 
	End If
	Set Upload = Nothing
	UpdateUserpostExc()
End Sub 

Sub UpdateUserpostExc()
	'�û����ֲ���
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

'�����ϴ����ݲ����ظ���ID
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
	Echo "		oEditor.InsertHtml('[UPLOAD="&FileExt&"]ShowFile.asp?ID="&Fileid&"[/UPLOAD]<br>�����ߴ磺"& Round(FileSize/1024,2) &" KB<br>');	 "
	Echo " }else{"
	Echo "		parent.message.value+='[UPLOAD="&FileExt&"]ShowFile.asp?ID="&Fileid&"[/UPLOAD]\n�����ߴ磺"& Round(FileSize/1024,2) &" KB\n';	 "
	Echo " };"
	Echo "	parent.document.getElementById('updiv').style.display='';					"
	Echo "	parent.document.getElementById('showupdiv').style.display='block';					"
	Echo "	parent.document.getElementById('showupfileids').style.color='red';				"
	Echo "	parent.document.getElementById('showupfileids').innerHTML='�ϴ������ɹ���';"
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

'���·��Զ������ϴ��ļ���,��Ҫ�ƣӣ����֧�֡�
Private Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	'�����´����ϴ��ļ��У���ʽ��2003��8
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