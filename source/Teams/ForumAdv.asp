<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim uID,fID,x1,x2,ComeUrl
Uid = HRF(2,2,"Uid")
team.Headers(Team.Club_Class(1) & "- ��̳�ƹ�ϵͳ")
'ComeUrl = Lcase(trim(request.ServerVariables("HTTP_REFERER")))
'If ComeUrl = "" Then
	'team.Error "������Դ����"
'Else
	Call Main
'End if
Sub Main()
	If UId = 0 Then
		team.Error "�û���������"
	Else
		Dim Rs,UserPostID,My_ExtSort,uName,i
		Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u
		If CID(Request.Cookies("advclass")("forumsadv")) = 1 Then
			Response.Redirect team.Club_Class(2)
			Response.End
		End If
		If CID(Session("advclass")) = 1 Then
			Response.Redirect team.Club_Class(2)
			Response.End
		End if
		Set Rs = team.execute("Select UserName From ["&IsForum&"User] Where ID=" & Int(Uid))
		If Rs.Eof And Rs.Bof Then
			team.Error "ϵͳ�����ڴ��û���"
		Else
			If Trim(tk_UserName) = Trim(Rs(0)) Then 
				team.Error1 "��Ϊ������Ϊ�û����ˣ����Ա��η�����Ч����ȴ�ϵͳ�Զ�������̳��<meta http-equiv=refresh content=3;url=default.asp>"
			End if
			ExtCredits = Split(team.Club_Class(21),"|")
			MustOpen = Split(team.Club_Class(22),"|")
			uName = ""
			For U=0 to Ubound(ExtCredits)
				ExtSort=Split(ExtCredits(U),",")
				MustSort=Split(MustOpen(U),",")
				If ExtSort(3)=1 Then
					If U = 0 Then
						UExt = UExt &"Extcredits0=Extcredits0+"&MustSort(7)&""
					Else
						UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"+"&MustSort(7)&""
					End If
					uName = uName & "����"& ExtSort(0) &"������"& MustSort(7) & ExtSort(1) &"<br />" 
				End if
			Next
			team.execute("Update ["&IsForum&"User] Set "&UExt&",Newmessage=Newmessage+1 Where UserName = '"&Rs(0)&"'")
			team.Execute("insert into ["&Isforum&"Message] (author,incept,content,Sendtime,MsgTopic) values ('��ϵͳ��Ϣ��','"&Rs(0)&"','�������Ա���̳�Ĵ����ƹ㣬���õ�ϵͳ�������Ľ������£�<br> "&uName&" ',"&SqlNowString&",'�ƹ�ϵͳ֪ͨ')")
			'�ж�Cookies����Ŀ¼
			Dim cookies_path_s,cookies_path_d,cookies_path
			cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
			cookies_path_d=ubound(cookies_path_s)
			cookies_path="/"
			For i=1 to cookies_path_d-1
				cookies_path=cookies_path&cookies_path_s(i)&"/"
			Next
			Response.Cookies("advclass").Expires = DateAdd("s", 360, Now())
			Response.Cookies("advclass").Path = cookies_path
			Response.Cookies("advclass")("forumsadv") = "1"
			Session("advclass") = "1"
		End If
		team.Error1 "����������.... <meta http-equiv=refresh content=3;url=default.asp>"
		Rs.Close:Set Rs=Nothing
	End if
End Sub
team.footer
%>
