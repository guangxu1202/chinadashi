<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
'-------------版权说明----------------
'BOKECC视频展区 V3.0 TEAM BOARD ASP版
'本列表为CC视频联盟会员专用的支持文件,可以完整显示您站点上所有的CC视频
'官方网站：http://www.bokecc.com   官方论坛:http://bbs.bokecc.com
'本程序由 team board 修改,官方论坛:http://www.team5.cn 欢迎转载

'*****请修改以下信息以便您能正常使用视频展区功能*****'
Dim tID,fID,x1,x2,vID
tID = HRF(2,2,"tid")
fID = HRF(2,2,"fid")
team.Headers(Team.Club_Class(1) & "- 视频展区")
'=======================功能设置区===================================================================

vID = team.Forum_setting(115)	    '您的展区ID，替换为里面的0.展区ID请登陆到union.bokecc.com登录管理后台---安装视频展区里查看，记住必须点-生成展区'

'=======================功能设置区===================================================================
If CID(team.Forum_setting(116))=0 Then
	Echo "本功能模块已经关闭."
Else
	Call Main()
End If

team.footer


Sub Main()
	x1 = "<a href=""Cclist.asp"">视频展区</a>"
	Echo team.MenuTitle
	Echo " <iframe style='PADDING: 0px; MARGIN: 0px;align:center;overflow-x:hidden;width:100%;' src=""http://show.bokecc.com/showzone/"& vID&""" frameBorder='0' width='100%' scrolling=auto height='1350px' allowTransparency='true'></iframe>"
End Sub 
%>