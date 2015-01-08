<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Call main()

Sub main()
	Dim tID,BbsName,rs1
	tID = HRF(2,2,"tid")
	Echo "<html> "
	Echo "	<head> "
	Echo "	<style type=""text/css"">"
	Echo "		body,table {font-size: 12px; font-family: Tahoma, Verdana }"
	Echo "	</style>"
	Echo "	</head>"
	Echo "	<body leftmargin=""40"">"
	Dim Rs,RforumName,PTitle
	Set Rs = team.execute("Select ReList,Topic,Username,Posttime,Content,forumid From ["&Isforum &"Forum] Where Locktopic = 0 and CloseTopic = 0 and ID="& tID )
	If Rs.Eof Then
		team.error "不存在此贴内容，或此帖被关闭了。"
	Else
		Set Rs1 = team.Execute("Select ID,followid,Pass,Bbsname,lookperm From ["&IsForum&"bbsconfig] where hide=0 and id="&CID(RS(5))&" ")
		If ( Rs1.Eof and RS1.Bof) Then
			team.error "此板块不存在或您没有查看此板块的权限"
		Else
			Bbsname = RS1(3)
		End If
		If Int(Rs1(1)) = 0 Then 
			Response.Redirect "../Default.asp?rootid="&RS(5)
		End if
		If Not (RS1(4) = ",") Then
			If Instr(RS1(4),",") > 0 Then 
				Response.Redirect "../Thread.asp?tid="& tid
			End If
		End If
		If Rs1(2) <> "" Then
			Response.Redirect "../Thread.asp?tid="& tid
		End if
		RS1.Close:Set RS1=Nothing	
		RforumName = Rs(0)
		PTitle = Rs(1)
		Echo "	<title>"&Team.Club_Class(1)&" - "&RS(1)&" - Power By Team Board</title>"
		Echo "	<b>标题: </b> "&RS(1)&" &nbsp; &nbsp;  [<a href=""###"" onclick=""this.style.visibility='hidden';window.print();this.style.visibility='visible'""><b>打印本页</b>]</a><br> "
		Echo "	<b>来自: </b> <a href="""& Team.Club_Class(2) &"""> "& Team.Club_Class(1) &" </a> - <a href="""& Team.Club_Class(2) &"/Forums.asp?fid="&Rs(5)&"""> "& Sign_Code(BbsName,1) &" </a><br> "
		Echo "	<b>链接: </b> <a href="""& Team.Club_Class(2) &"/Thread.asp?tid="&tID&" ""> "& Team.Club_Class(2) &"/Thread.asp?tid="&tID&" </a><br> "
		Echo "	<hr noshade size=""3"" width=""100%"" color=""#808080"">"
		Echo "	<b>作者: </b> "&Rs(2)&" &nbsp; &nbsp; <b>时间: </b> "&RS(3)&" &nbsp; &nbsp; <b>标题: </b> "&RS(1)&"<br />"
		Echo "	<br /> "&ReadPowers(UBB_Code(RS(4))) &" <br /> "
	End if
	Rs.Close:Set Rs=Nothing
	Set Rs = team.execute("Select Username,Content,Posttime,ReTopic From "& Isforum & RforumName&" Where Topicid = "& tID)
	Do While Not Rs.Eof
		Echo "	<hr noshade size=""1"" width=""100%"" color=""#808080"">"
		Echo "	<b>作者: </b> "&Rs(0)&" &nbsp; &nbsp; <b>时间: </b> "&Rs(2)&" &nbsp; &nbsp; <b>标题: </b> "&IIF(RS(3)<>"",RS(3),"RE: "&PTitle&"")&"<br /> " 
		Echo "	<br /> "&ReadPowers(UBB_Code(RS(1))) &" <br /> "
		Rs.Movenext
	Loop
	Rs.Close:Set Rs=Nothing
	Echo "	<hr noshade size=""3"" width=""100%"" color=""#598687""><p align=""right"">Powered by Team Board</p> "
	team.Htmlend
End sub
%>
