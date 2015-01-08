<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim ID,Upid,FileName,RS,FileUrl,Types,i,IsNotExts,fid

If Request("action") = "iunm" Then
	Call NumCount()
Else 
	Call main()
End If 

Sub NumCount
	Dim dNum,Rs,s
	dNum = Fixjs(Request("num"))
	Set Rs = team.execute("Select Upcount from [Upfile] Where FILEID="& CID(dNum) )
	If Not Rs.Eof Then
		s = Rs(0)
	Else
		s = 0
	End If
	Rs.Close:Set Rs=Nothing
	Response.Write "document.write('"
	Response.Write s
	Response.Write "');"
	Response.Write vbNewline
End Sub
	
Function Fixjs(Strings)
	Dim Str
	Str = Strings
	str = Replace(str, CHR(39), "\'")
	str = Replace(str, CHR(13), "")
	str = Replace(str, CHR(10), "")
	str = Replace(str, "]]>","]]&gt;")
	Fixjs = str
End Function

Sub main()
	Types = Array("gif","jpg","jpeg","bmp","png","swf","swi")
	ID = CID(Request("ID"))
	Set Rs=team.execute("Select FileName,UpCount,Types from [upfile] where Fileid="&Int(ID) )
	If Rs.BOF and Rs.EOF Then
		team.error "该附件已被删除"
		Response.End
	Else
		IsNotExts = False
		For i = 0 To UBound(Types)
			If InStr(Rs(2),Types(i))>0 Then IsNotExts = True
		Next
		If Not IsNotExts Then UpdateUserpostExc()
		team.execute("Update [upfile] Set UpCount=UpCount+1 where Fileid="&Int(ID) )
		FileUrl="Images/UpFile/"
		If team.Forum_setting(93) = 0 Then
			Response.Redirect FileUrl&rs(0)
		Else
			FileName=ReplaceStr(rs(0),"..","")&""
			If Request.ServerVariables("HTTP_REFERER")="" Or InStr(Request.ServerVariables("HTTP_REFERER"),Request.ServerVariables("SERVER_NAME"))=0 Or FileName="" Then 
				Response.Redirect "Default.asp"
			Else
				Call DownLoadFile(Server.MapPath(FileUrl&FileName))
			End If
		End If
	End If
	Rs.Close:Set Rs=Nothing
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
				UExt = UExt &"Extcredits0=Extcredits0-"&MustSort(4)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"-"&MustSort(4)&""
			End If
			If (team.User_SysTem(14+U)-MustSort(4))-MustSort(8)<0 Then
				team.Error "您的"&ExtSort(0)&" ["& team.User_SysTem(14+U) - MustSort(4) &"] 低于积分策略下限值 ["& MustSort(8)&"] ，所以无法进行此操作。"
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where ID = "& Int(team.TK_UserID))
End Sub

Sub DownLoadFile(strFile)
	On error resume next
	Server.ScriptTimeOut=999999
	Dim S,fso,f,intFilelength,strFilename
	strFilename = strFile
	Response.Clear
	Set s = Server.CreateObject("ADODB.Stream") 
	s.Open
	s.Type = 1 
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	If Not fso.FileExists(strFilename) Then
		Response.Write("<h1>错误: </h1><br>系统找不到指定文件")
		Exit Sub		
	End If
	Set f = fso.GetFile(strFilename)
		intFilelength = f.size
		s.LoadFromFile(strFilename)
		If err Then
		 	Response.Write("<h1>错误: </h1>" & err.Description & "<p>")
			Response.End 
		End If
		Set fso=Nothing
		Dim Data
		Data=s.Read
		s.Close
		Set s=Nothing
		If Response.IsClientConnected Then 
			Response.AddHeader "Content-Disposition", "attachment; filename=" & f.name 
			Response.AddHeader "Content-Length", intFilelength 
 			Response.CharSet = "UTF-8" 
			Response.ContentType = "application/octet-stream"
			Response.BinaryWrite Data
			Response.Flush
		End If
End Sub
%>

