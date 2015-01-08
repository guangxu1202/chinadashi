<!-- #include file="Conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
If team.Forum_setting(33)=0 Then
	Response.ContentType = "text/xml"
	Echo "<?xml version=""1.0"" encoding=""gbk""?><rss version=""2.0""><item>RSS订阅关闭</item></rss>"
Else
	Call RssMain()
End If
Sub rssMain()
	Dim SQL,Tag,ForumID,TopCount,XmlRs,rssTitle
	TopCount = 30 '取出数据条数

	ForumID=CID(Request.QueryString("fid"))
	Tag=CID(Request.QueryString("Tag"))
	SQL=" ID,Topic,UserName,PostTime,Content"
	Select Case Tag
		Case "1"
			SQL="Select Top "& TopCount & " "& SQL &" From ["&IsForum&"Forum] Where deltopic<>1 and Posttime>"&SqlNowString&"-7 Order By Views Desc"
			rssTitle = "论坛本周人气帖"
		Case "2"
			SQL="Select Top "& TopCount & " "& SQL &" From ["&IsForum&"Forum] Where deltopic<>1 and Posttime>"&SqlNowString&"-7 Order By Replies Desc"
			rssTitle = "论坛本周热门帖"
		Case "3"
			SQL="Select Top "& TopCount & " "& SQL &" From ["&IsForum&"Forum] Where Goodtopic=1 And deltopic<>1 Order By ID Desc"
			rssTitle = "论坛精华帖"
		Case "4"
			SQL="Select "& SQL &" From ["&IsForum&"Forum] Where ForumID="&ForumID&" And deltopic<>1 Order By ID Desc"
			rssTitle = "定阅本版帖子更新"
		Case "5"
			SQL="Select "& SQL &" From ["&IsForum&"Forum] Where ID="&ForumID&" And deltopic<>1 Order By ID Desc"
			rssTitle = "定阅帖子更新"
		Case  Else 
			SQL="Select Top "& TopCount & " "& SQL &" From ["&IsForum&"Forum] Where deltopic<>1 and posttime>"&SqlNowString&"-7 Order By ID Desc"
			rssTitle = "论坛新帖"
	End Select
	Dim i,rs,iGUid
	Response.ContentType = "text/xml"
	With Response
		.Write"<?xml version=""1.0"" encoding=""gbk""?> " 
		.Write"<rss version=""2.0"">" 
		.Write"<channel> "
		.Write"<title>"& rssTitle &"</title> "
		.Write"<link>"& xmlfilter(team.Club_Class(4)) &"/XML.ASP</link>" 
		.Write"<description>TEAM Board - "& xmlfilter(team.Club_Class(3)) &"</description> "
		.Write"<copyright>"& team.Forum_setting(8) &"</copyright>"
		.Write"<generator>TEAM Board by TEAM5.Cn Studio</generator> "
		.Write"<ttl>"&xmlfilter(team.Forum_setting(34))&"</ttl>"
		Set Rs=team.Execute(SQL)
		If (Rs.Eof And Rs.Bof) Then
			Response.Write("<item />")
		Else
			XmlRs=Rs.GetRows(-1)
			Rs.Close:Set Rs=Nothing
		End If
		If IsArray(XmlRs) Then
			For i=0 To Ubound(XmlRs,2)
				.Write("<item>")
				.Write("<link>"& xmlfilter(team.Club_Class(4)) &"/Thread.asp?tid="& XmlRs(0,i) &" </link>")
				.Write("<title>"& xmlfilter(XmlRs(1,i)) &"</title>")
				.Write("<author>"&XmlRs(2,i)&"</author>")
				.Write("<pubDate>"&XmlRs(3,i)&"</pubDate>")
				Set iGUid = team.execute("Select UserGroupID from ["& isforum &"User] Where UserName='"& XmlRs(2,i) &"' ")
				If Not iGUid.Eof Then
					If iGUid(0) = 6 Or iGUid(0) = 7 Then
						.Write("<description>==该用户已被锁定==</description>")
					Else
						.Write("<description><![CDATA["&CleariCode(XmlRs(4,i))&"]]></description>")
					End If
				Else
					.Write("<description><![CDATA["&CleariCode(XmlRs(4,i))&"]]></description>")
				End If
				iGUid.Close:Set iGUid = Nothing
				.Write("</item>")
			Next
		End If
		.Write("</channel></rss>")
	End With
	Conn.Close
	Set Conn=Nothing
End Sub

function xmlfilter(a)
	If a="" or IsNull(a) Then 
		Exit Function
	Else
		If Instr(a,"'")>0 Then 
			a = replace(a, "'","&#39;")
		End If
		a = replace(a, ">", "&gt;")
		a = replace(a, "<", "&lt;")
		a = Replace(a, "&", "&amp;")
		a = Replace(a, "'", "&apos;")
		a = Replace(a, CHR(34), "&quot;")
		xmlfilter = a
	End If
end Function

function CleariCode(s)
	If s="" or IsNull(s) Then 
		Exit Function
	Else
		Dim re
		set re = New RegExp
		re.Global = True
		re.IgnoreCase = True
		re.Pattern="\[REPLAYVIEW\]((.|\n)*)\[\/REPLAYVIEW\]"
		s=re.Replace(s,"回复可见贴的内容需要登陆论坛完整版本才可以查看")
		re.Pattern="\[BUY=*([0-9]+)\]((.|\n)*)\[\/BUY\]"
		s=re.Replace(s,"购买可见贴的内容需要登陆论坛完整版本才可以查看")
		CleariCode = s
	End If
End  Function

%>
