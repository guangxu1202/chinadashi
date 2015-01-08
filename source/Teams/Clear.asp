<link href=manage/images/admin.css rel=stylesheet>
<table  border=0  cellPadding=3 cellSpacing=1 width='100%'  align=center class=a2>

<tr class=a1><td colspan=2>缓存管理页面
</td></tr>

<tr class=a1><td colspan=2><%

Response.Write date() &"<br>"
Response.Write now() &"<br>"
Response.Write time() &"<br>"
%>
</td></tr>
<%
For Each Thing in Application.Contents
			Response.Write "<tr class=a4><td><font color=Gray>" & thing & "</font>&nbsp;</td><td>状态："
			If isObject(Application.Contents(Thing)) Then
				'Application.Contents(Thing).close
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				Response.Write "对象成功关闭"
			ElseIf isArray(Application.Contents(Thing)) Then
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = null
				Response.Write "数组成功释放"
			Else
				Response.Write Application.Contents(Thing)
				Application.Contents(Thing) = null
			End If
			Response.Write "</td></tr>"
Next
%>
<form name="form1" method="POST" action="?action=clear">
<tr class=a3><td colspan=2><input type="submit" name="Submit" value="清空缓存">
</td></tr>
</form>
<%
if Request.QueryString("action")="clear" then
Application.Contents.RemoveAll()
Response.Write"缓存已清除"
End If
%>
</table>

