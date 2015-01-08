<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="team board,team论坛系统" name=author />
<meta name="keywords" content="<%=team.Forum_setting(30)%>" />
<meta name="description" content="<%=team.Forum_setting(31)%>" />
<meta name="copyright" content="Copyright 2005-2008 - team5.cn By DayMoon" />
<link rel="icon" href="favicon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<title><%=team.Club_Class(1)%> - 分栏模式 - Powered by team studio</title>
<style type="text/css">
	body {margin: 0px;}
	#frameswitch {background: url(images/tree/frame_switch.gif) no-repeat 0;cursor: pointer;}
</style>
<script type="text/javascript">
function framebutton(){
	var obj = document.getElementById('navigation');
	var frameswitch = document.getElementById('frameswitch');
	var switchbar = document.getElementById('switchbar');
	if(obj.style.display == 'none'){
		obj.style.display = '';
		switchbar.style.left = '147px';
		frameswitch.style.backgroundPosition = '0';
	}else{
		obj.style.display = 'none';
		switchbar.style.left = '0px';
		frameswitch.style.backgroundPosition = '-11';
	}
}
if(top != self) {
	top.location = self.location;
}
</script>
</head>
<body scroll="no">
<table border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
	<tr>
		<td align="middle" id="navigation" valign="center" name="frametitle" width="180">
			<iframe name="leftmenu" frameborder="0" src="Forumlist.asp" scrolling="auto" style="height: 100%; visibility: inherit; width: 150px; z-index: 1"></iframe>
		<td style="width: 100%">
			<table id="switchbar" border="0" cellPadding="0" cellSpacing="0" width="11" height="100%" style="position: absolute; left: 148px; background-repeat: repeat-y; background-position: -148px">
				<tr><td onClick="framebutton()"><img id="frameswitch" src="images/tree/none.gif" alt="" border="0" width="11" height="49" /></td></tr>
			</table>
			<iframe frameborder="0" scrolling="yes" name="main" src="default.asp" style="height: 100%; visibility: inherit; width: 100%; z-index: 1;overflow: auto;"></iframe>
		</td>
	</tr>
</table>
</body>












