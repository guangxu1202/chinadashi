<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
'-------------��Ȩ˵��----------------
'BOKECC��Ƶչ�� V3.0 TEAM BOARD ASP��
'���б�ΪCC��Ƶ���˻�Աר�õ�֧���ļ�,����������ʾ��վ�������е�CC��Ƶ
'�ٷ���վ��http://www.bokecc.com   �ٷ���̳:http://bbs.bokecc.com
'�������� team board �޸�,�ٷ���̳:http://www.team5.cn ��ӭת��

'*****���޸�������Ϣ�Ա���������ʹ����Ƶչ������*****'
Dim tID,fID,x1,x2,vID
tID = HRF(2,2,"tid")
fID = HRF(2,2,"fid")
team.Headers(Team.Club_Class(1) & "- ��Ƶչ��")
'=======================����������===================================================================

vID = team.Forum_setting(115)	    '����չ��ID���滻Ϊ�����0.չ��ID���½��union.bokecc.com��¼�����̨---��װ��Ƶչ����鿴����ס�����-����չ��'

'=======================����������===================================================================
If CID(team.Forum_setting(116))=0 Then
	Echo "������ģ���Ѿ��ر�."
Else
	Call Main()
End If

team.footer


Sub Main()
	x1 = "<a href=""Cclist.asp"">��Ƶչ��</a>"
	Echo team.MenuTitle
	Echo " <iframe style='PADDING: 0px; MARGIN: 0px;align:center;overflow-x:hidden;width:100%;' src=""http://show.bokecc.com/showzone/"& vID&""" frameBorder='0' width='100%' scrolling=auto height='1350px' allowTransparency='true'></iframe>"
End Sub 
%>