<!-- #include file="conn.asp" -->
<!-- #include file="INC/Const.asp" -->
<%
Dim fID,tID,x1,x2,rID,Mstemp,GInfo
Dim MasterMenu,Inputinfo,temp
Public Values,Board_Setting
fID = HRF(1,2,"fid")
tID = HRF(1,2,"tid")
Call ManageClass()

Select Case Request("action")
	Case "move"
		MasterMenu="�ƶ�����"
		Inputinfo = "moves"
		call movepage
	Case "moves"
		Call forummove
	Case "deltopic"
		MasterMenu="ɾ������"
		Inputinfo = "deltopics"
		Call movepage
	Case "deltopics"
		Call deltopics
	case "movenew"
		MasterMenu="��ǰ����"
		Inputinfo = "movenews"
		call movepage
	case "movenews"
		call movenews
	Case "islockpage"
		MasterMenu="����/�����������"
		Inputinfo = "islockpages"
		call movepage
	Case "islockpages"
		islockpages
	Case "isclosepage"
		MasterMenu="�ر�/������"
		Inputinfo = "isclosepages"
		call movepage
	Case "isclosepages"
		isclosepages
	case "settops"
		MasterMenu="�ö�/����ö�"
		Inputinfo = "settopss"
		call movepage
	case "settopss"
		call settopss
	Case "getlike"
		MasterMenu="���������������"
		Inputinfo = "getlikes"
		call movepage
	Case "getlikes"
		Call getlikes
	Case "digest"
		MasterMenu="����/�������"
		Inputinfo = "digests"
		call movepage	
	Case "iskillofget"
		MasterMenu="����/�ͷ�"
		Inputinfo = "iskillofget"
		call movepage
	Case "digests"
		Call digests
	Case Else
		team.Error "��������"
End Select	

Sub movepage 
	If request.form("ismanage")="" and request.form("fismanage")="" Then 
		team.Error("��û��ѡ���������Ӧ�Ĺ���������뷵���޸ġ�")
	End If
	Echo "<form method=""post"" action=""?action="&Inputinfo&""">"
	Echo "<input type=""hidden"" value="""&fID&""" name=""fid"">"
	Echo "<input type=""hidden"" value="""&Request("rid")&""" name=""rid"">"
	Echo "<input type=""hidden"" value="""&Request("fismanage")&""" name=""fismanage"">"
	Echo "<table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=""a2"">"
	Echo " <tr>"
	Echo "		<td class=""a1"" colspan=""2"">TEAM's��ʾ: "&MasterMenu&"</td>"
	Echo " </tr>"
	Echo " <tr class=""a4"">"
	Echo "		<td width=""40%""><B>�û���</B>:</td><td>"&TK_UserName&"</td>"
	Echo "</tr>"
	If Request("action")="settops" Then
		Echo "<tr class=""a4"">"
		Echo "	<td><b>����:</b></td>"
		Echo "	<td><input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked> �ö����� &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;  "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> ����ö� "
		Echo "	</td>"
		Echo "</tr>	 "
		Echo "<tr class=""a4"">"
		Echo "	<td>ģʽ:</td>"
		Echo "	<td><input type=""radio"" name=""istohead"" value=""1"" class=""radio"" checked> �����ö� &nbsp; &nbsp; <input type=""radio"" name=""istohead"" value=""2"" class=""radio""> ���ö� "
		Echo "</td>"
		Echo "</tr> "
	End If
	If Request("action")="settops" or Request("action")="digest" Then 
		Echo "<tr class=a4>"
		Echo "	<td><b>��������:</b></td>"
		Echo "	<td>"
		Echo "	<input type=""radio"" name=""icolor"" value=""999"" class=""radio"" checked> Ĭ������ &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""icolor"" value=""1"" class=""radio""> <span style=""width:10px;background:#808080;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""2"" class=""radio""> <span style=""width:10px;background:#808000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""3"" class=""radio""> <span style=""width:10px;background:#008000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""4"" class=""radio""> <span style=""width:10px;background:#0000ff;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""5"" class=""radio""> <span style=""width:10px;background:#800000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""6"" class=""radio""> <span style=""width:10px;background:#ff0000;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""7"" class=""radio""> <span style=""width:10px;background:#cc0066;"">&nbsp;</span> "
		Echo "	<input type=""radio"" name=""icolor"" value=""0"" class=""radio""> ȡ������ &nbsp; &nbsp; "
		Echo "</td>"
		Echo "</tr>"
		Echo "<tr class=a4>"
		Echo "	<td><b>��������:</b></td>"
		Echo "	<td>"
		If Request("action")="digest" Then
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""1"" class=""radio"" checked> ��Ϊ�������� &nbsp; &nbsp; "
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""2"" class=""radio""> �������"
		Else
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""0"" class=""radio"" checked> ������ &nbsp; &nbsp; "
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""1"" class=""radio""> ��Ϊ�������� &nbsp; &nbsp; "
			Echo "	<input type=""radio"" name=""togoodtopic"" value=""2"" class=""radio""> �������"
		End if
		Echo "</td>"
		Echo "</tr>"
	End If
	If Request("action")="isclosepage" Then
		Echo "<tr class=""a4"">"
		Echo "	<td><b>����:</b></td>"
		Echo "	<td><input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked>  �ر����� &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> ������"
		Echo "	</td>"
		Echo "</tr> "
	End If
	If Request("action")="islockpage" Then
		Echo "<tr class=""a4""> "
		Echo "		<td><b>����:</b></td> "
		Echo "	<td><input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked> �������� &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> �����������"
		Echo "	</td>"
		Echo "</tr>"
	End If
	If Request("action")="move" Then
		Echo " <tr class=""a4""> "
		Echo "		<td><b>Ŀ����̳/����:</b></td><td>"
		Echo "		<select name=moveid> "
		Echo "		<option selected value="""">�������ƶ���...</option>"
		Echo			team.BBs_Value_List(0,0)
		Echo "		</Select></td> "
		Echo " </tr>"
	End If
	If Request("action")="getlike" Then
		Dim Special,utmp,u
		Special = ""
		If Int(Board_Setting(15))=1 and Int(Board_Setting(17))=1 Then
			If Instr(Board_Setting(19),Chr(13)&Chr(10))>0 Then
				utmp = Split(Board_Setting(19),Chr(13)&Chr(10))
				For U=0 To Ubound(utmp)
					Special = Special &" <option value="""&U&""">"& utmp(u) &"</option>" 
				Next
			Else
				Special = "<option value=""0"">"& Board_Setting(19) &"</option>"
			End if	
		End If
		Echo " <tr class=""a4""> "
		Echo "	<td>���������������:</td><td>"
		Echo "	<input type=""radio"" name=""isclose"" value=""0"" class=""radio"" checked> ������� &nbsp; &nbsp; "
		Echo "	<input type=""radio"" name=""isclose"" value=""1"" class=""radio""> �Ƴ����� "
		Echo "	</td> "
		Echo " </tr>"
		Echo " <tr class=""a4""> "
		Echo "		<td>���������������:</td><td>"
		Echo "		<select name=posttopic><option value=""999"">��ѡ�����·���ר��</option>"
		Echo		Special
		Echo "		</Select></td> "
		Echo " </tr>"
	End If
	Echo "<tr class=""a4"">"
	Echo "<td><B>����ԭ��:</B><BR> ���������µȼ����й��������д����ԭ�򣬳������ϵȼ�����������д����ԭ��<BR> "
	If Not team.IsMaster Then Echo " <B>��������д�ſ��Խ��й��������</B> "
	Echo " <br><input type=""checkbox"" name=""sendpm"" value=""1"" class=""radio"" "
	If request("action")="movenew" Then 
		Echo "disabled"
	End  If
	If Not team.IsMaster Then
		Echo " checked "
	End if
	Echo "> ������Ϣ֪ͨ����"
	Echo "</td><td>"
	Echo "<select name=""selectreason"" size=""6"" style=""height: 8em; width: 8em"" onchange=""this.form.reason.value=this.value"">"
	Echo"<option value="""">�Զ���</option>"
	Dim KillInfo,i
	If Instr(team.Club_Class(8),Chr(13)&Chr(10))>0 Then
		KillInfo = Split(team.Club_Class(8),Chr(13)&Chr(10))
		For i = 0 To Ubound(KillInfo)
			Echo "<option value="""&KillInfo(i)&""">"&KillInfo(i)&"</option>"
		Next
	Else
		Echo "<option value="""&team.Club_Class(8)&""">"&team.Club_Class(8)&"</option>"
	End if	
	Echo " </select> "
	Echo " <textarea name=""reason"" style=""height: 8em; width: 18em""></textarea></td>"
	Echo " </tr>"
	if Request("action")<>"movenew" Then
		Echo "<tr class=a4>"
		Echo "<td><B>�û�����:</B><BR>������ֵ����Ϊ���������۳����֣��෴����ֵΪ��������Ϊ�ӷ֡� </td>"
		Echo "<td>"
		Dim ExtCredits,m,ExtSort
		ExtCredits = Split(team.Club_Class(21),"|")
		For m = 0 To UBound(ExtCredits)
			ExtSort=Split(ExtCredits(M),",")
			If Split(ExtCredits(M),",")(3)=1 Then
				Echo ExtSort(0) & " <select name=""ExtCredits"&M&"""  size=""1""> "
				Call Kills
				Echo "</select> &nbsp;"
			End if
		Next
		Echo "</td>"
		Echo "</tr>"
		Echo "<tr class=a4>"
		Echo "<td><B>׷�ӿ۷�:</B><BR>��ѿ۷ּ���Ĭ�Ͽ۷ֵ�ǰ���£��ٴ�׷�ӿ۷����������ڶ�ĳ�û��Ļ��ֽ��н����� </td>"
		Echo "<td><input type=""radio"" name=""douser"" value=""0"" class=""radio"" checked>Ĭ�Ͽ۷� &nbsp; <input type=""radio"" name=""douser"" value=""1"" class=""radio"">׷�ӿ۷�"
		Echo "</td></tr>"
	End if
	Echo "	 </table><br /><center><input type=""submit"" value="" ȷ �� ""><br /> "
	Dim Rs2,SQL2,ho,rs1
	If request("rid")="" or (Not isnumeric(Request("rid"))) Then
		Echo "<br /><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=a2><tr class=a1><td>����</td><td>����</td><td>�ظ�</td><td>��󷢱�</td></tr>"
		for each ho in request.form("ismanage")
			Set Rs2=Team.Execute("select id,topic,UserName,Replies,Lasttime from ["&IsForum&"Forum] where id="&CID(ho))
			If Not Rs2.Eof Then
				Echo "<tr class=a4><td><input type=checkbox name=ismanages value="&rs2(0)&" checked><a href=Thread.asp?tid="&rs2(0)&" target=_blank>"&Rs2(1)&"</a></td><td>"&Rs2(2)&"</td><td>"&Rs2(3)&"</td><td>"&Rs2(4)&"</td></tr>"
			End if
			Rs2.Close:Set Rs2=Nothing
		next
		Echo "</table></form>"
	Else
		If Request("menu")="move" and (Request.Form("fismanage")="" or not isnumeric(Request.Form("fismanage"))) Then
			team.error " ��ֻ�ܶ���������ƶ���������ѡ������ID��"
		End If
		if isnumeric(Request.Form("fismanage")) and Request.Form("fismanage")<>"" Then
			Response.Write "<br /><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=a2><tr class=a1><td width=""30%"">����</td><td width=""10%"">����</td><td width=""10%"">�ظ�</td><td width=""20%"">��󷢱�</td><td width=""10%"">�û�IP</td></tr>"
			Set Rs2=Team.Execute("select id,topic,UserName,Replies,Lasttime,postip from ["&IsForum&"Forum] where id="& CID(Request.Form("fismanage")) )
			If Not Rs2.Eof Then
				Response.Write "<tr class=a4><td><input type=checkbox name=ismanages value="&rs2(0)&" checked><a href=Thread.asp?tid="&rs2(0)&" target=_blank>"&Rs2(1)&"</a></td><td>"&Rs2(2)&"</td><td>"&Rs2(3)&"</td><td>"&Rs2(4)&"</td><td>"&Rs2(5)&"</td></tr>"
			End if
			Rs2.Close:Set Rs2=Nothing
			Response.Write "</table>"
		Else
			Set Rs1=Team.Execute("select Relist,id,topic from ["&IsForum&"Forum] where id="&CID(request("rid")))
			If Not Rs1.Eof Then
				Echo "<br /><table width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" class=a2><tr class=a1><td>����</td><td>����ID</td><td>����</td><td>�ظ�ʱ��</td><td>�û�IP</td></tr>"
				Echo "<input type=""hidden"" value="&Rs1(0)&" name=relistname>"
				for each ho in request.form("ismanage")
					Set Rs2=Team.Execute("select id,UserName,posttime,postip from ["&IsForum & Rs1(0)&"] where id="&CID(ho))
					If Not Rs2.Eof Then
						Echo "<tr class=a4><td><input type=checkbox name=ismanages value="&rs2(0)&" checked> <a href=Thread.asp?tid="&rs1(1)&"#"&Rs2(0)&" target=_blank>"&Rs1(2)&"</a></td><td>"&Rs2(0)&"</a></td><td>"&Rs2(1)&"</td><td>"&Rs2(2)&"</td><td>"&Rs2(3)&"</td></tr>"
					End if
					Rs2.Close:Set Rs2=Nothing
				Next
				Echo "</table></form>"
			End If
			Rs1.Close:Set Rs1=Nothing
		End if
	End If
End Sub

Sub getlikes()
	Dim ho,rs,posttopic
	posttopic = HRF(1,2,"posttopic")
	If CID(team.Group_Manage(12)) = 1 then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select username,topic from ["&IsForum&"forum] where id="&CID(ho))
			If Not Rs.BOF then
				If request.form("isclose")=0 Then
					Team.execute("update ["&IsForum&"forum] set PostClass="&Int(posttopic)&" where id="&CID(ho))
					Mstemp = "��������������� "
					Call delpiont(RS(0))
					GInfo = "���������������"
					Call Pmsetto(rs(0),rs(1))
					temp="���������������ɹ�"
				Else
					Team.execute("update ["&IsForum&"forum] set PostClass=999 where id="&CID(ho))
					temp="�����Ƴ��������ɹ�"
					Mstemp =  "�����Ƴ��������"
					Call delpiont(RS(0))
					GInfo = "�����Ƴ��������"
					Call Pmsetto(rs(0),Rs(1))
				End if
			End if
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error("<li>�����ڵ��� "&team.Levelname(0)&" û������������������Ȩ��")
	End if
	Call Serverend(Mstemp)
End Sub

Sub Kills()
	Dim i,t
	For i = -50 To -1
		t = t &"<option value="""&i&""">"&i&"</option>"
	Next
	t = t &"<option value=""0"" selected>0</option>"
	For i = 1 To 50
		t = t &"<option value="""&i&""">"&i&"</option>"
	Next
	Echo t
End Sub

Sub digests
	Dim ho,rs,iSetColor
	If team.Group_Manage(8) = "1" then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select username,id,topic from ["&IsForum&"forum] where id="&CID(ho))
			If Not Rs.BOF Then
				If Request.Form("icolor") <> 999 Then
					iSetColor = ",color=" & CID(Request.Form("icolor"))
				End if
				If request.form("togoodtopic")=1 Then
					team.execute("update ["&IsForum&"forum] set goodtopic=1"& iSetColor &" where id="&CID(ho))
					team.execute("update ["&IsForum&"user] set goodtopic=goodtopic+1 where username='"&RS(0)&"'")
					UpdateUserpostExc RS(0),2
					Mstemp =  "�û�"&tk_UserName&"������[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]���뾫������"
					Temp = "���Ӽ��뾫�����ɹ� !"
					GInfo = "�Ӿ�"
				ElseIf request.form("togoodtopic")=2 Then
					team.execute("update ["&IsForum&"forum] set goodtopic=0"& iSetColor &" where id="&CID(ho))
					team.execute("update ["&IsForum&"user] set goodtopic=goodtopic-1 where username='"&RS(0)&"'")
					KillUpdateUserpostExc RS(0),2
					Mstemp = "�û�"&tk_UserName&"������[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]�Ƴ���������"
					Temp = "�����Ƴ��������ɹ� !"
					GInfo = "ȡ������"
				End If
			End if
			Call delpiont(RS(0))
			Call Pmsetto(RS(0),RS(2))
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error " �����ڵ��� "&team.Levelname(0)&" û�м���/������������Ȩ�ޡ�" 
	End if
	Call Serverend(Mstemp)
End Sub

Sub UpdateUserpostExc(uname,m)
	'�û����ֲ���
	Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u
	Dim UserPostID,My_ExtSort
	If Not team.UserLoginED Then  Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	For U=0 to Ubound(ExtCredits)
		ExtSort=Split(ExtCredits(U),",")
		MustSort=Split(MustOpen(U),",")
		If ExtSort(3)=1 Then
			If U = 0 Then
				UExt = UExt &"Extcredits0=Extcredits0+"&MustSort(m)&""
			Else
				UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"+"&MustSort(m)&""
			End if
		End if
	Next
	team.execute("Update ["&IsForum&"User] Set "&UExt&" Where UserName = '"& HTmlEncode(uname) &"' ")
End Sub

Sub KillUpdateUserpostExc(uname,m)
	'�û����ֲ���
	Dim ExtCredits,MustOpen,ExtSort,MustSort,UExt,u,Rs
	Dim UserPostID,My_ExtSort
	If Not team.UserLoginED Then  Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	UExt = ""
	Set Rs = team.execute("Select Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7 From ["&IsForum&"User] Where UserName = '"& HTmlEncode(uname) &"' ")
	If Rs.Eof And Rs.Bof Then
		Exit Sub
	Else 
		For U=0 to Ubound(ExtCredits)
			ExtSort=Split(ExtCredits(U),",")
			MustSort=Split(MustOpen(U),",")
			If ExtSort(3)=1 Then
				If U = 0 Then
					If (Rs(0+U)-MustSort(4))-MustSort(8)<=0 Then
						UExt = UExt &"Extcredits0=0"
					Else
						UExt = UExt &"Extcredits0=Extcredits0-"&MustSort(m)&""
					End if
				Else
					If (Rs(0+U)-MustSort(4))-MustSort(8)<=0 Then
						UExt = UExt &",Extcredits"&U&"=0"
					Else
						UExt = UExt &",Extcredits"&U&"=Extcredits"&U&"-"&MustSort(m)&""
					End if
				End If
			End If
		Next
		team.execute("Update ["&IsForum&"User] Set "&UExt&" Where UserName = '"& HTmlEncode(uname) &"' ")
	End If 
End Sub

Sub settopss
	Dim ho,rs,iSetColor
	If CID(team.Group_Manage(1)) >= 1 then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select Username,ID,Topic,Goodtopic from ["&IsForum&"forum] Where ID="&CID(ho))
			If Not Rs.BOF Then
				If Request.Form("icolor") <> 999 Then
					iSetColor = ",color=" & CID(Request.Form("icolor"))
				End if
				If request.form("isclose")=0 Then
					If request.form("istohead")=1 and CID(team.Group_Manage(1)) >= 1 Then
						team.execute("update ["&IsForum&"forum] set toptopic=1"& iSetColor &" where id="&CID(ho))
						Mstemp = "�û�"&tk_UserName&"������ [<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]�����˱����ö������� "
						GInfo = "�����ö�"
					Elseif request.form("istohead")=2 and CID(team.Group_Manage(1)) = 2 Then
						team.execute("update ["&IsForum&"forum] set toptopic=2"& iSetColor &" where id="&CID(ho))
						Mstemp = "�û�"&tk_UserName&"������ [<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]���������ö������� "
						GInfo = "���ö�"
					End If
					Call delpiont(RS(0))
					temp="���������ö��ɹ���"
				Else
					If CID(team.Group_Manage(1)) >= 1 Then
						team.execute("update ["&IsForum&"forum] set toptopic=0"& iSetColor &" where id="&CID(ho))
						Mstemp = "�û�"&tk_UserName&"���ö������� [<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]����Ϊ��������״̬��"
						Call delpiont(RS(0))
						temp="��������ö��ɹ���"
						GInfo = "����ö�"
					End if
				End if
				If request.form("togoodtopic")=1 and team.Group_Manage(8) = 1 Then
					If Rs(3) = 0 Then
						team.execute("update ["&IsForum&"forum] set goodtopic=1"& iSetColor &" where id="&CID(ho))
						team.execute("update ["&IsForum&"user] set goodtopic=goodtopic+1 where username='"&RS(0)&"'")
						Mstemp = Mstemp & "<br>�û�"&tk_UserName&"������[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]���뾫������"
					End if
				ElseIf request.form("togoodtopic")=2 and team.Group_Manage(8) = 1 Then
					If Rs(3) = 1 Then
						team.execute("update ["&IsForum&"forum] set goodtopic=0"& iSetColor &" where id="&CID(ho))
						team.execute("update ["&IsForum&"user] set goodtopic=goodtopic-1 where username='"&RS(0)&"'")
						Mstemp = Mstemp & "<br>�û�"&tk_UserName&"������[<a href=Thread.asp?tid="&RS(1)&">"&RS(2)&"</a>]�Ƴ���������"
					End if
				End If
				Call Pmsetto(rs(0),Rs(2))
			End if
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error "�����ڵ��� "&team.Levelname(0)&" û�м���/������������Ȩ��"
	End if
	Call Serverend(Mstemp)
End Sub

Sub isclosepages
	Dim ho,rs
	If team.Group_Manage(7) = 1 then
		for each ho in request.form("ismanages")
			set rs=team.execute("Select username,topic from ["&IsForum&"forum] where id="&CID(ho))
			If Not Rs.BOF then
				If request.form("isclose")=0 Then
					Team.execute("update ["&IsForum&"forum] set CloseTopic=1 where id="&CID(ho))
					Mstemp = "�ر�����  [url=Thread.asp?tid="&CID(ho)&"]"&RS(1)&"[/url] "
					Call delpiont(RS(0))
					GInfo = "�ر�����"
					Call Pmsetto(rs(0),rs(1))
					temp="�ر�����ɹ�"
				Else
					Team.execute("update ["&IsForum&"forum] set CloseTopic=0 where id="&CID(ho))
					temp="�򿪱��رյ�����ɹ�"
					Mstemp =  "�򿪱��ر�����  [url=Thread.asp?tid="&CID(ho)&"]"&RS(1)&"[/url]"
					Call delpiont(RS(0))
					GInfo = "�򿪹ر�����"
					Call Pmsetto(rs(0),Rs(1))
				End if
			End if
			Rs.Close:Set Rs=Nothing
		next
	Else
		Team.Error("<li>�����ڵ��� "&team.Levelname(0)&" û�йر�/�������Ȩ��")
	End if
	Call Serverend(Mstemp)
End Sub

Sub movenews
	Dim ho
	If team.Group_Manage(5) = 1 then
		for each ho in request.form("ismanages")
			team.execute("update ["&IsForum&"forum] set Lasttime="&SqlNowString&" Where id="&CID(ho))
			temp="��ǰ����ɹ�"
			Mstemp =  "��ǰ����"
		next
	Else
		Team.Error "�����ڵ��� "&team.Levelname(0)&" û����ǰ�����Ȩ��"
	End if
	Call Serverend(Mstemp)
End Sub

Sub islockpages
	Dim ho,rs,RName
	If team.Group_Manage(6) = 1 then
		for each ho in request.form("ismanages")
			If request.form("isclose")=0 Then
				if isnumeric(Request("rid")) and Request("rid")<>"" and (Request.Form("fismanage")="" or not isnumeric(Request.Form("fismanage"))) then
					set rs=team.execute("Select ReList,username,Topic from ["&IsForum&"forum] where id="&Request("rid"))
					If Not Rs.BOF then
						team.execute("update ["&IsForum & rs(0)&"] set lock=1 where topicid="&CID(Request("rid"))&" and id="&CID(ho))
						RName = team.execute("select username from ["&IsForum & rs(0)&"] where id="&CID(ho))(0)
						Mstemp =  "��������"
						Call delpiont(RName)
						GInfo = "����"
						Call Pmsetto(RName,RS(2))
						Temp="���������ɹ���"
					end if
					Rs.Close:Set Rs=Nothing
				else
					set rs=team.execute("Select username,topic from ["&IsForum&"forum] where id="&CID(ho))
					If Not Rs.BOF then
						team.execute("update ["&IsForum&"forum] set Locktopic = 1 where id="&CID(ho))
						Mstemp =  "�������� [url=Thread.asp?tid="&CID(ho)&"]"&RS(1)&"[/url]"
						Call delpiont(RS(0))
						GInfo = "����"
						Call Pmsetto(RS(0),RS(1))
					End If
					Rs.Close:Set Rs=Nothing
					Temp="��������ɹ���"
				end if
			Else
				if isnumeric(Request("rid")) and Request("rid")<>"" then
					set rs=team.execute("Select ReList,username,Topic from ["&IsForum&"forum] where id="&Request("rid"))
					If Not Rs.BOF then
						team.execute("update ["&IsForum & rs(0)&"] set lock=0 where topicid="&CID(Request("rid"))&" and id="&CID(ho))
						Mstemp =  "�����������"
						GInfo = "����"
						Call delpiont(team.execute("select username from ["&IsForum & rs(0)&"] where id="&CID(ho))(0))
						Call Pmsetto(team.execute("select username from ["&IsForum & rs(0)&"] where id="&CID(ho))(0),RS(2))
						Temp="������������ɹ���"
					end If
					Rs.Close:Set Rs=Nothing
				else
					set rs=team.execute("Select username,Topic from ["&IsForum&"forum] where id="&CID(ho))
					If Not Rs.BOF then
						team.execute("update ["&IsForum&"forum] set Locktopic =0 where id="&CID(ho))
						Mstemp =  "�����������"
						Call delpiont(RS(0))
						GInfo = "����"
						Call Pmsetto(RS(0),RS(1))
					End If
					Rs.Close:Set Rs=Nothing
					Temp="������������ɹ���"
				end if
			End If
		next
		Application.Contents.RemoveAll()
	Else
		Team.Error "�����ڵ��� "&team.Levelname(0)&" û������/������ӵ�Ȩ��"
	End if
	Call Serverend(Mstemp)
End Sub

Sub forummove()
	Dim Ts,UpID,ho,Rs,SQL
	Dim Board_Setting
	if team.Group_Manage(4) = 1 then
		if Request("moveid")="" then 
			team.Error "��û��ѡ��Ҫ�������ƶ��ĸ���̳!"
		End if
		If Request("moveid")=Request("fid") Then 
			team.Error "��ѡ�����̳��Դ��̳��ͬ!"
		End If
		Board_Setting = team.Execute("Select Board_Setting From ["&IsForum&"bbsconfig] where ID="&CID(Request("moveid")))(0)
		if Split(Board_Setting,"$$$")(2) = 1 Then
			team.Error "Ŀ����̳������˰�飬����ת�롣"
		End if
		for each ho in request.form("ismanages")
			Set Rs = team.execute("Select forumid,topic,Toptopic,Locktopic,Lasttime,UserName,ID from ["&IsForum&"forum] where id="&CID(ho))
			If Not (Rs.BOF and Rs.EOF) Then
				team.execute("Update ["&IsForum&"forum] set forumid="&int(Request("moveid"))&",topic='"&RS(1)&"',Toptopic=0,Locktopic=0,Lasttime="&SqlNowString&"  Where ID="&CID(ho))
				GInfo = "�ƶ�����"
				Call delpiont(RS(5))
				Call Pmsetto(RS(5),Rs(1))
				Mstemp = "�ƶ�����ID : [url=Thread.asp?tid="&RS(6)&"]"&RS(1)&"[/url][BR]"
			End If
			Temp="�ƶ�����ɹ�"
			Rs.Close:Set Rs=Nothing
			UpID = Team.Execute("Select Max(ID) From ["&IsForum&"Forum] Where deltopic=0 and forumid="& Request("fid"))(0)
			set Ts=team.execute("select top 1 topic,Lasttime,username,ID from ["&IsForum&"forum] where ID="& UpID )
			If Not Ts.Eof Then
				team.execute("update ["&IsForum&"bbsconfig] set Board_Last='<A href=Thread.asp?tid="&TS(3)&" target=""_blank"">"&Cutstr(TS(0),200)&"</a> ��$@$"&TS(2)&"$@$"&Now()&"' where id="&Request("fid"))
			End If
			Ts.Close:Set Ts=Nothing
		Next
		Cache.DelCache("BoardLists")
	Else
		team.Error " �����ڵ��� "&team.Levelname(0)&" û���ƶ������Ȩ�� "
	End if
	Call Serverend(Mstemp)
End Sub

Sub deltopics
	if team.Group_Manage(3) = 1 then
		Dim Forum_ID,Max_ID,rs1,ho,rs,Isnames,DayDel
		for each ho in request.form("ismanages")
			if Request("rid")<>"" and isnumeric(Request("rid")) And (Request.Form("fismanage")="" or not isnumeric(Request.Form("fismanage"))) then
				set rs=team.execute("Select forumid,ReList,Topic,Posttime from ["&IsForum&"forum] where id="&Request("rid"))
				If Not Rs.BOF then
					Isnames= team.execute("select username from ["&IsForum & rs(1)&"] where id="&CID(ho))(0)
					Call delpiont(Isnames)
					GInfo = "ɾ������"
					Call Pmsetto(Isnames,RS(2))
					KillUpdateUserpostExc Isnames,1
					team.execute("delete from ["&IsForum & rs(1)&"] where id="&CID(ho))
					team.execute("update ["&IsForum&"forum] set replies=replies-1 where id="&CID(Request("rid")))
					If DateDiff("d",RS(3),Date())=0 Then
						DayDel = "today=today-1,"
					End If
					team.execute("update ["&IsForum&"bbsconfig] set "& DayDel &"tolrestore=tolrestore-1 where id="&rs(0))
					team.execute("update ["&IsForum&"user] set postrevert=postrevert-1 where username='"&Isnames&"'")
					If CID(DateDiff("d",CDate(RS(3)),Now()))=1 Then
						team.LockCache "TodayNum" , Application(CacheName&"_TodayNum")-1
					End If
				End If
				Temp="ɾ�������ɹ�"
				Mstemp = "ɾ������"
				Rs.close:Set Rs=Nothing
			Else
				'If Request.Form("fismanage")<>"" And IsNumeric(Request.Form("fismanage")) Then
					'set rs=team.execute("Select forumid,topic,toptopic,goodtopic,Locktopic,lasttime,UserName,id,ReList from ["&IsForum&"forum] where id="& Int(Request.Form("fismanage")))
				'Else
				set rs=team.execute("Select forumid,topic,toptopic,goodtopic,Locktopic,lasttime,UserName,id,ReList from ["&IsForum&"forum] where id="&CID(ho))
				'End if
				If Not Rs.BOF then
					team.execute("update ["&IsForum&"user] set deltopic=deltopic+1 where username='"&rs(7)&"'")
					team.execute("update ["&IsForum&"forum] set toptopic=0,deltopic=1,lasttime="&SqlNowString&",LastText='"&tk_UserName &"$@$�����ѱ�"&tk_UserName &"ɾ��' where deltopic=0 and id="&CID(ho))
					'����������
					team.execute("delete from ["&IsForum &"FVote] where RootID="&CID(ho))
					team.execute("delete from ["&Isforum&"Activity] where RootID="&CID(ho))
					team.execute("delete from ["&Isforum&"ReActivity] where RootID="&CID(ho))
					team.execute("delete from ["&Isforum&"ActivityUser] where RootID="&CID(ho))
					If DateDiff("d",RS(3),Date())=0 Then
						DayDel = "today=today-1,"
					End If
					Max_ID=Team.Execute("Select Max(ID) from ["&IsForum&"forum] where deltopic=0 and Forumid="&rs(0))(0)
					If Max_ID<>"" Then
						Set Rs1=Team.Execute("Select ID,topic,username,posttime from ["&IsForum&"forum] where deltopic=0 and id="&Max_ID)
						if Not rs1.eof then
							team.execute("update ["&IsForum&"bbsconfig] set "&DayDel&"toltopic=toltopic-1,Board_Last='<A href=Thread.asp?tid="&rs1(0)&" target=""_blank"">"&Cutstr(rs1(1),200)&"</a> ��$@$"&rs1(2)&"$@$"&Now()&"' where id="&rs(0))
						End If
						Rs1.Close:Set Rs1 = Nothing
					Else
						team.execute("update ["&IsForum&"bbsconfig] set "&DayDel&"toltopic=toltopic-1,Board_Last='��������$@$"&TK_UserName&"$@$"&Now()&"' where id="&rs(0))
					End If

					Call delpiont(RS(6))
					GInfo = "ɾ������ "
					Call Pmsetto(RS(6),RS(1))
					KillUpdateUserpostExc RS(6),0
				End If
				Temp = "ɾ������ɹ�"
				Mstemp =  Mstemp & " ɾ������ : ["& Rs(1) &"]"
			End If
		Next
		Cache.DelCache("BoardLists")
	Else
		Team.Error "�����ڵ��� "&team.Levelname(0)&" û��ɾ�����ӵ�Ȩ��"
	End if
	Call Serverend(Mstemp)
End Sub

Sub Serverend(s)
	team.SaveLOG("�û�"&TK_UserName&"����: "&s)
	team.Error1 ("<li>"&temp&"<li><a href=""Forums.asp?fid="&request("fid")&""">������̳</a><li><a href=""Default.asp"">������̳��ҳ</a><meta http-equiv=refresh content=3;url=""Forums.asp?fid="&request("fid")&""">")
End Sub

Sub delpiont(s)
	Dim ExtCredits,m,ExtSort,GetMyExs,ExcName,MustOpen,MustSort,Rs
	If HRF(1,2,"douser") = 0 Then Exit Sub
	ExtCredits = Split(team.Club_Class(21),"|")
	MustOpen = Split(team.Club_Class(22),"|")
	GetMyExs=""
	Set Rs = team.execute("Select Extcredits0,Extcredits1,Extcredits2,Extcredits3,Extcredits4,Extcredits5,Extcredits6,Extcredits7 From ["&IsForum&"User] Where UserName = '"& HTmlEncode(s) &"' ")
	If Rs.Eof And Rs.Bof Then
		team.Error "�����ڴ��û�"
	Else 
		For m = 0 To UBound(ExtCredits)
			ExtSort=Split(ExtCredits(M),",")
			MustSort=Split(MustOpen(m),",")
			If Split(ExtCredits(M),",")(3)=1 Then
				If Request.Form("ExtCredits"&M) <> "0" Then
					If GetMyExs = "" Then
						If (Rs(0+M)-Request.Form("ExtCredits"&M))-MustSort(8)<=0 Then
							GetMyExs = "ExtCredits"& M& "=0"
						Else
							GetMyExs = "ExtCredits"& M& "=ExtCredits"& M& "+"& Request.Form("ExtCredits"&M)
						End if
					Else
						If (Rs(0+M)-Request.Form("ExtCredits"&M))-MustSort(8)<=0 Then
							GetMyExs = GetMyExs & ",ExtCredits"& M& "=0"
						Else
							GetMyExs = GetMyExs & ",ExtCredits"& M& "=" & "ExtCredits"& M& "+ "& Request.Form("ExtCredits"&M)
						End if
					End If
					ExcName = ExcName & ExtSort(0) &" : "& Request.Form("ExtCredits"&M)
				End If
			End If
		Next
	End If 
	If GetMyExs <>"" Then
		if s=TK_UserName then 
			team.Error "�㲻�ܶ��Լ����в���!"
		Else
			team.execute("Update ["&IsForum&"User] Set "& GetMyExs &" Where UserName='"&Htmlencode(s)&"' ")
		End If
		temp = temp & "<br>" & ExcName
	End if
End Sub
Sub Pmsetto(s,m)
	If request("sendpm") = "1" Then 
		Dim Istemp,ho
		If Not team.IsMaster and len(request("reason"))<2 Then 
			team.error2 "��û����д����ԭ��"
		Else
			Istemp = "��������̳ϵͳ�Զ����͵�֪ͨ����Ϣ��[br] "
			If request("rid")<>"" or isnumeric(Request("rid")) Then 
				Istemp = Istemp & " �������⣺[url=Thread.asp?tid="&request("rid")&"]"&m&"[/url] �Ļظ����� [br] "
			Else
				Istemp = Istemp & " ������������⣺ "
				for each ho in request.form("ismanages")
					Istemp = Istemp & "  [url=Thread.asp?tid="&CID(ho)&"] "&m&"[/url] [br]"
				Next
			End if
			Istemp = Istemp & " �� "&tk_UserName&" ִ�� "& GInfo  &" ���� [br] ��������:  "&request("reason")&" ��"
		End If
		Team.Execute("insert into ["&IsForum&"message](author,incept,content,Sendtime,MsgTopic) values ('"&TK_UserName&"','"&HTmlEncode(s)&"','"&HTmlEncode(Istemp)&"',"&SqlNowString&",'[ϵͳ��Ϣ]����������ӱ�ִ�й������!')")
		Team.Execute("update ["&IsForum&"user] set newmessage=newmessage+1 where username='"&HtmlEncode(s)&"'")
	End If
End Sub

Sub ManageClass()
	Dim Rs
	team.ChkPost()
	Set Rs = team.Execute("Select ID,bbsname,Board_Setting From ["&IsForum&"bbsconfig] Where ID="&fID)
	If Rs.Eof Then
		team.Error " ��������"
	Else
		Values = Rs.GetRows(-1)
	End If
	If isarray(Values) Then
		Board_Setting = Split(Values(2,0),"$$$")
	End if
	team.Headers("��̳���ӹ������� - "& Values(1,0))
	x1="<a href=""Forums.asp?fid="&fID&""">"& Values(1,0)  &"</a> "
	x2=" ��̳���ӹ������� "
	Echo team.MenuTitle
	If Not team.UserLoginED Then
		team.Error " ��δ��½��̳��<meta http-equiv=refresh content=3;url=login.asp> "
	End if
	If Not team.ManageUser Then
		team.Error " ����Ȩ�޲��������ܲ�����̳���� ��"
	Else
		If Not team.IsMaster and Not team.SuperMaster Then
			If team.execute("Select ID from ["&Isforum&"Moderators] Where ManageUser='"& tk_username &"' and BoardID = "& fid).eof Then 
				team.Error " �����Ǵ˰�İ���,���ܲ���˰�Ĺ���"
			End If
		End if
	End if
End Sub
Team.footer
%>