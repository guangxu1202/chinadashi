<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="team board,team论坛系统" name=author />
<meta name="keywords" content="<%=team.Forum_setting(30)%>" />
<meta name="description" content="<%=team.Forum_setting(31)%>" />
<meta name="copyright" content="Copyright 2005-2008 - team5.cn By DayMoon" />
<link href="images/tree/tree.css" rel="stylesheet" type="text/css" id="css" />
<title><%=team.Club_Class(1)%> - 分栏模式 - Powered by team studio</title>
<script language="JavaScript">
function showsubmenu(i) {
	var ot1 = document.getElementById('showsubmenu_' + i + '');
	var ot2 = document.getElementById('menus_' + i + '');
	if (ot1.style.display == 'none') {
		ot1.style.display = '';
		ot2.src = 'images/tree/L2.gif';
	}else{
		ot1.style.display = 'none';
		ot2.src = 'images/tree/L1.gif';
	}
}
</script>
</head>
<body>
<div class="treeclass">
	<img src="images/tree/L6.gif" border="0" align="absmiddle">
<%
	Echo ListShow
	Function ListShow()	
		Dim tmp1,i,Boards
		Boards = team.myBoardJump()
		If IsArray(Boards) Then
			For i = 0 To UBound(Boards,2)
				If Boards(2,i)=0 Then
					tmp1 = tmp1 & "<div id='"&i&"'><a href=""#"" onClick=""showsubmenu('"&i&"'); return false"" class=""bigword""><IMG SRC='images/tree/L2.gif'  BORDER=""0"" align=""absmiddle"" id='menus_"&i&"'>"&Boards(1,i) &"</a></div> <div id='showsubmenu_"&i&"' class='parent'> "& ListShow_Li(Boards(0,i),0) &"</div>"
				End if
			Next
		End if
		ListShow = tmp1 
	End Function

	Function ListShow_Li(a,b)
		Dim tmp1,i,Boards
		Boards = team.myBoardJump()
		If isArray(Boards) Then
			For i=0 To Ubound(Boards,2)
				If Boards(2,i) = a Then
					If b>0 Then 
						tmp1 = tmp1 & "<IMG SRC='images/tree/L4.gif' BORDER=""0"" align=""absmiddle"">"
					End if
					tmp1 = tmp1 & "<a href=""Forums.asp?fid="&Boards(0,i)&""" target='main'><IMG SRC='images/tree/L3.gif' BORDER=""0"" align=""absmiddle"">"& Boards(1,i)&"</a><br> " & ListShow_Li(Boards(0,i),1)
				End If
			Next
		End if
		ListShow_Li = tmp1
	End function
%>
	<div>
	<%If Not team.UserLoginED Then 
			Echo "<a href=""login.asp"" target=""main"">[ 登陆 ]</a> <a href=""reg.asp"" target=""main"">[ 注册 ]</a> "
		Else
			Echo "<img src=""images/tree/L5.gif"" border=""0"" align=""absmiddle""> <A href=""Login.asp?menu=out"" target=""main"">退出论坛</A>" 
		End if
	%>
	</div>
	<div style="text-align:left;margin-top:10px;">
		<a href="Rss.asp" target="main"><img src="images/xml.gif" border="0" class="absmiddle" alt="RSS订阅全部论坛" /></a> 
	</div>
</div>
</body></html>


