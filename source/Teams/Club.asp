<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim Message
Message = Request("message")
Echo "<link href=""skins/teams/bbs.css"" rel=""stylesheet"">"
Echo "<BR><BR><table border=""0"" cellspacing=""1"" cellpadding=""8"" width=""80%"" align=""center"" class=""a2"">"
Echo "	<tr class=""a1""><td align=""center"" colspan=""2""> ϵͳ��ʾ��Ϣ </td></tr>"
Echo "	<tr class=""a4""><td align=""center"" width=""40%"">��Ϣ����</td><td>"	
If team.Forum_setting(56)>=1 Then
	Dim nexhour,openclock,i
	nexhour=Hour(Now())
	If team.Forum_setting(56)=1 Then Echo "<li>��̳�����˶�ʱ���ţ��밴�����ʱ����ʣ�"
	If team.Forum_setting(56)=2 Then Echo "<li>��̳�����˶�ʱֻ�����밴�����ʱ�䷢����"
	Echo "<TABLE border=0 cellspacing=0 cellpadding=0><tr class=a4>"
	openclock=Split(team.Forum_setting(0),"*")
	For i= 0 to UBound(openclock)
		Echo  "<td>��"&i &"�㣺</td>"
		Echo  " <td>" 
		If openclock(i)=1 Then 
			Echo "��<font color=red>����</font>��"
		Else
			Echo "��<font color=blue>�ر�</font>��"
		end if
		Echo "��</td>"
		If (i+1) mod 4 = 0 Then Echo  "</tr>"
	Next
	Echo "</TABLE>"
ElseIf team.Forum_setting(2)=1 Then
	Echo  team.Forum_setting(3)
Else
	If Message="" Then Message="ϵͳ����"
	Echo  Message 
End if	
Echo "	</td></tr>"
Echo " </table><br><form action=""Login.asp?menu=add"" method=""post"" name=""mylogin"">"
Echo "<table border=""0"" cellspacing=""1"" cellpadding=""8"" width=""80%"" align=""center"" class=""a2"">"
Echo "	<tr class=""a1""><td align=""center"" colspan=""2""> ����Ա��½ </td></tr>"
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> �û����� : </span></td>"
Echo "    <td class=""altbg2""><input size=""25"" name=""username"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""></td>"
Echo "</tr>"
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> �û����� : </span> </td>"
Echo "    <td class=""altbg2""><input size=""25"" type=""password"" name=""userpass"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""></td>"
Echo "</tr>"
If team.Forum_setting(48)>=1 Then
	Echo "<tr> "
	Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> ��֤�� : </span> </td>"
	Echo "    <td class=""altbg2""><input size=""25"" name=""code"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""> <img src=""inc/code.asp"" alt=""��֤��,�������?����ˢ����֤��"" style=""cursor : pointer;"" onclick=""this.src='inc/code.asp'"" /></td>"
	Echo "</tr>"
End If
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> ��ȫ����: </span> <br>����������˰�ȫ����,�ͱ�����д��ȷ�Ĵ𰸲ſ��Ե�½ </td>"
Echo "    <td class=""altbg2""><input size=""25"" name=""questionid"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""> <select onchange=""document.mylogin.questionid.value=this.value"" name=""select""> "
Echo "      <option value="""">�ް�ȫ����</option>"
Echo "      <option value=""ĸ�׵�����"">ĸ�׵�����</option>"
Echo "      <option value=""үү������"">үү������</option>"
Echo "      <option value=""���׳����ĳ���"">���׳����ĳ���</option>"
Echo "      <option value=""������һλ��ʦ������"">������һλ��ʦ������</option>"
Echo "      <option value=""�����˼�������ͺ�"">�����˼�������ͺ�</option>"
Echo "      <option value=""����ϲ���Ĳ͹�����"">����ϲ���Ĳ͹�����</option>"
Echo "      <option value=""��ʻִ�յ������λ����"">��ʻִ�յ������λ����</option>"
Echo "      </select></td> "
Echo "</tr>"
Echo "<tr> "
Echo "    <td width=""40%"" class=""altbg1""> <Span class=""bold""> �ش� : </span> </td>"
Echo "    <td class=""altbg2""><input size=""25"" name=""answer"" value="""" onBlur=""this.className='colorblur';"" onfocus=""this.className='colorfocus';this.value=''"" class=""colorblur""></td>"
Echo "</tr>"
Echo "</table><br><center><input type=""submit"" value="" ��¼ "" name=""Submit""></center> </form>"
%>
