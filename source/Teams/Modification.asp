<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/MD5.asp" -->
<%
team.Headers(Team.Club_Class(1) &" - �����һ�")
Dim X1,X2,Fid,acc
X1="��������"
Select Case Request("action")
	Case "edit"
		Call Edit
	Case Else
		Call Main
End Select

Sub Main
	Echo Replace(Team.ElseHtml (7),"{$weburl}",team.MenuTitle)
End Sub

Sub Edit
	Dim cookies_path,cookies_path_s,cookies_path_d
	Dim i,username,question,answer,UserMail,rs
	UserName = HRF(1,1,"username")
	Question = HRF(1,1,"question")
	Answer = HRF(1,1,"answer")
	UserMail = HRF(1,1,"email")
	If (Cid(Session("Login")) >= Cid(team.Forum_setting(54))) or Request.Cookies(Forum_sn)("OpenLogin")=1 Then
		team.error "���Ѿ����� "&team.Forum_setting(54)&" �����������Ϣ��ϵͳ���������ڴγ��ԡ�"
		cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
		cookies_path_d=ubound(cookies_path_s)
		cookies_path="/"
		For i=1 to cookies_path_d-1
			cookies_path=cookies_path&cookies_path_s(i)&"/"
		Next
		Response.Cookies(Forum_sn)("OpenLogin") = 1
		Response.Cookies(Forum_sn).Expires=Date+1
		Response.Cookies(Forum_sn).path = cookies_path
	Else
		Session("Login") = Session("Login") +1
	End If
	If Trim(UserName) = "" Then
		team.Error "�û�������Ϊ��"
	End if
	If Not IstrueName(UserName) Then 
		team.Error " �����û����д�����ַ��� "
	End If
	If Not IsValidEmail(UserMail) Then
		team.Error  "�ʼ���ʽ���� !"
	End If	
	Set Rs=Team.Execute("Select Answer,Question,UserMail,UserGroupID From ["&IsForum&"User] Where UserName='"&username&"' ")
	If RS.Eof or Rs.Bof Then
		Rs.Close:Set RS=Nothing
		Team.Error "ϵͳ�����ڴ��û� "
	Else
		If Rs(3) = 1 Or Rs(3) = 2  Then 
			Team.Error "���û��ĵȼ��޷�ͨ�������һع���!"
		End if
		If Trim(RS(2)) <> Trim(UserMail) Then 
			Team.Error "����д��ȷ��Email��ַ"
		Else
			Session("Login") = 0
			Dim num1,Title,Mailtopic,Body
			Randomize
			num1= Mid((Rnd*999999),1,6)
			If CID(team.Forum_setting(1))=0 Then
				If Rs(0)&""="" Or Rs(1)&""="" Then
					team.Error " ��Ϊϵͳ��֧���ʼ����͹��ܣ�����ֻ���û����ð�ȫ���ʵ�ǰ���²ſ���ʹ�������һع��ܡ�"
				End if
				If Trim(Rs(1))<>question or Trim(RS(0))<>answer Then 
					Team.Error "����Ĵ�!"
				Else
					Team.Execute("Update ["&IsForum&"User] Set userpass='"&Md5(num1,16)&"' Where UserName='"&username&"'")
					team.Error "�𾴵��û� "&username&" ����������Ѿ����޸�Ϊ[ "&num1&" ] <br> ϵͳ���� 30 ����Զ�ת���½���档<meta http-equiv=refresh content=30;url=Login.asp>"
				End If
			Else
				If Rs(1)<>"" Or Rs(2)<>"" Then 
					If Trim(Rs(1))<>question or Trim(RS(0))<>answer Then 
						Team.Error "����Ĵ�!"
					End If
				End if
				Team.Execute("Update ["&IsForum&"User] Set userpass='"&Md5(num1,16)&"' Where UserName='"&username&"'")
				Mailtopic="�������������! ["&team.Club_Class(1)&"ϵͳ��Ϣ-Power By Team Borad]"
				Body=""&vbCrlf&"�װ���"&username&", ����!"&vbCrlf&""&vbCrlf&"��ϲ! ���Ѿ��ɹ����һ���������,"& vbCrlf &" ����������Ϊ��"&num1&" �����½��̳���޸��������롣 "&vbCrlf&"  �ǳ���л��ʹ��"&team.Club_Class(3)&"�ķ���!"&vbCrlf&""&vbCrlf&"�����, �м���ע�����������μ�"&vbCrlf&"1�������ء��������Ϣ�������������ȫ��������취�����һ�й涨��"&vbCrlf&"2��ʹ�����ɶ������Ļ��⣬�����벻Ҫ�漰���Ρ��ڽ̵����л��⡣"&vbCrlf&"3���е�һ����������Ϊ��ֱ�ӻ��ӵ��µ����»����·������Ρ�"&vbCrlf&""&vbCrlf&""&vbCrlf&"��̳������ "&team.Club_Class(1)&"("&team.Club_Class(2)&") �ṩ����������:TEAM5.CN [By DayMoon]"&vbCrlf&""&vbCrlf&""&vbCrlf&""
				Call Emailto(UserMail,Mailtopic,Body)
				Team.Error " �����Ѿ�������ע������䣬��ע����� ��<meta http-equiv=refresh content=3;url=Default.asp> "
			End If
		End If
	End If
End Sub
Team.footer
%>