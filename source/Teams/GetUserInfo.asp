<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim Rs,uName,uNum,fid
team.Headers(Team.Club_Class(1) & "- �û�����ҳ��")
uName = HRF(0,1,"getname")
uNum = HRF(0,1,"getid")
If uName = "" Or IsNull(uName) Then
	team.Error "�û�������Ϊ�ա�"
End If 
If uNum = "" Or Len(uNum)<16 Then
	team.Error "��֤�벻��Ϊ�ա�"
End If
Session("sLogin") = CID(Session("sLogin")) + 1
If CID(Session("sLogin"))>5 Then 
	team.Error "���Ѿ�����5�������˴������֤���롣"
End If
Set Rs=team.execute("Select RegNum,UserGroupID From ["&IsForum&"User] Where UserName='"&uName&"'")
If Rs.Eof And Rs.Bof Then
	team.Error "ϵͳ�����ڴ��û���"
Else
	If Len(uNum)<>16 Then
		team.Error "������֤�������"
	ElseIf Trim(uNum) <> Trim(RS(0)) Then
		team.Error "������֤�������"
	Else
		If Int(Rs(1))=5 Then
			Session("sLogin") = 0
			team.execute("Update ["&IsForum&"User] Set UserGroupID=27,Levelname='��Сһ�꼶||||||0||0',Members='ע���û�' Where UserName='"&uName&"'")
			team.Error1 "�����ʺ��Ѿ������ȴ�ϵͳ�Զ����ص���½ҳ�档<meta http-equiv=refresh content=3;url=login.asp>"
		Else
			team.Error "�����ʺ��Ѿ���������������д˳���<meta http-equiv=refresh content=3;url=login.asp>"
		End if
	End If 
End if
team.footer
%>
