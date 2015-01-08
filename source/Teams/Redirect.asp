<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim tID,fID
tID = HRF(2,2,"tid")
fID = HRF(2,2,"fid")
If tID = 0 Or fID = 0 Then
	team.Error "ID参数错误"
Else
	Dim Alls,Rs
	Alls = "Locktopic=0 and CloseTopic=0 and Deltopic=0 and "
	If Request("action") = "next" Then
		Set Rs = team.Execute("Select TOP 1 ID From ["&IsForum&"Forum] Where "&Alls&" forumid = "&fID&" and ID>"& tID & " Order By Lasttime")
		If Rs.Eof And Rs.bof Then
			team.Error "没有找到下一篇帖子"
		Else
			response.redirect "Thread.asp?tid="& Rs(0)
		End if
	Else
		Set Rs = team.Execute("Select TOP 1 ID From ["&IsForum&"Forum] Where "&Alls&" forumid = "&fID&" and ID<"& tID & " Order By Lasttime Desc ")
		If Rs.Eof And Rs.bof Then
			team.Error "没有找到上一篇帖子"
		Else
			response.redirect "Thread.asp?tid="& Rs(0)
		End if
	End If
	Rs.Close:Set Rs=Nothing
End if
%>
