<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim Message
Message = Request("message")
Echo "<link href=""skins/teams/bbs.css"" rel=""stylesheet"">"
Echo "<BR><BR><table border=""0"" cellspacing=""1"" cellpadding=""8"" width=""80%"" align=""center"" class=""a2"">"
Echo "	<tr class=""a1""><td align=""center"" colspan=""2""> 系统提示信息 </td></tr>"
Echo "	<tr class=""a4""><td align=""center"" width=""40%"">消息内容</td><td>"	
If team.Forum_setting(56)>=1 Then
	Dim nexhour,openclock,i
	nexhour=Hour(Now())
	If team.Forum_setting(56)=1 Then Echo "<li>论坛设置了定时开放，请按下面的时间访问："
	If team.Forum_setting(56)=2 Then Echo "<li>论坛设置了定时只读，请按下面的时间发帖："
	Echo "<TABLE border=0 cellspacing=0 cellpadding=0><tr class=a4>"
	openclock=Split(team.Forum_setting(0),"*")
	For i= 0 to UBound(openclock)
		Echo  "<td>　"&i &"点：</td>"
		Echo  " <td>" 
		If openclock(i)=1 Then 
			Echo "　<font color=red>开放</font>　"
		Else
			Echo "　<font color=blue>关闭</font>　"
		end if
		Echo "　</td>"
		If (i+1) mod 4 = 0 Then Echo  "</tr>"
	Next
	Echo "</TABLE>"
ElseIf team.Forum_setting(2)=1 Then
	Echo  team.Forum_setting(3)
Else
	If Message="" Then Message="系统错误"
	Echo  Message 
End if	
Echo "	</td></tr>"
Echo " </table><br><form action=""Login.asp?menu=add"" method=""post"" name=""mylogin"">"
Echo "<table border=""0"" cellspacing=""1"" cellpadding=""8"" width=""80%"" align=""center"" class=""a2"">"
Echo "	<tr class=""a1""><td align=""center"" colspan=""2""> 管理员登陆 </td></tr>"
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> 用户名称 : </span></td>"
Echo "    <td class=""altbg2""><input size=""25"" name=""username"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""></td>"
Echo "</tr>"
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> 用户密码 : </span> </td>"
Echo "    <td class=""altbg2""><input size=""25"" type=""password"" name=""userpass"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""></td>"
Echo "</tr>"
If team.Forum_setting(48)>=1 Then
	Echo "<tr> "
	Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> 验证码 : </span> </td>"
	Echo "    <td class=""altbg2""><input size=""25"" name=""code"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""> <img src=""inc/code.asp"" alt=""验证码,看不清楚?请点击刷新验证码"" style=""cursor : pointer;"" onclick=""this.src='inc/code.asp'"" /></td>"
	Echo "</tr>"
End If
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> 安全提问: </span> <br>如果您开启了安全提问,就必须填写正确的答案才可以登陆 </td>"
Echo "    <td class=""altbg2""><input size=""25"" name=""questionid"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""> <select onchange=""document.mylogin.questionid.value=this.value"" name=""select""> "
Echo "      <option value="""">无安全提问</option>"
Echo "      <option value=""母亲的名字"">母亲的名字</option>"
Echo "      <option value=""爷爷的名字"">爷爷的名字</option>"
Echo "      <option value=""父亲出生的城市"">父亲出生的城市</option>"
Echo "      <option value=""您其中一位老师的名字"">您其中一位老师的名字</option>"
Echo "      <option value=""您个人计算机的型号"">您个人计算机的型号</option>"
Echo "      <option value=""您最喜欢的餐馆名称"">您最喜欢的餐馆名称</option>"
Echo "      <option value=""驾驶执照的最后四位数字"">驾驶执照的最后四位数字</option>"
Echo "      </select></td> "
Echo "</tr>"
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> 回答 : </span> </td>"
Echo "    <td class=""altbg2""><input size=""25"" name=""answer"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""></td>"
Echo "</tr>"
Echo "</table><br><center><input type=""submit"" value="" 登录 "" name=""Submit""></center> </form>"
%>
