<!-- #include File="Conn.Asp" -->
<!-- #include File="Inc/Const.Asp" -->
<%
Dim X1,X2,Fid,Acc
team.Headers(Team.Club_Class(1) &" - ��̳����")
X2=" <A Href=Help.Asp><B>��̳����</B></A> "
Select Case Request("page")
	Case "custom"
		X1=" TEAM's Board ���ر�ʹ�ð��� "
		Echo Team.Menutitle
		Call custom
	Case "usermaint"
		X1=" �û���֪ "
		Echo Team.Menutitle
		Call usermaint
	Case "using"
		X1=" ��̳ʹ�� "
		Echo Team.Menutitle
		Call using	
	Case "messages"
		X1=" ��д���Ӻ��շ�����Ϣ "
		Echo Team.Menutitle
		Call messages	
	Case "mise"
		X1=" �������� "
		Echo Team.Menutitle
		Call mise
	Case Else
		X1="  "
		Echo Team.Menutitle
		Call Main
End Select
Team.footer

Sub Main
	Call Menu01 : Call Menu02 : Call Menu03 : Call Menu04
	If team.UserLoginED Then Call Menu05
End Sub

Sub Menu01 %>
	<div class="a2" id="center">
		<div class="a1"  style="padding: 5px;">TEAM's Board ���ر�ʹ�ð���</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=custom#0">���������ӹ���������涨</a></li>
			<li><a href="Help.asp?page=custom#1">��������Ϣ�������취</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu02 %>
	<div class="a2" id="center">
		<div class="a1"  style="padding: 5px;">�û���֪</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=usermaint#1">�ұ���Ҫע����</a></li>
			<li><a href="Help.asp?page=usermaint#2">TEAM's ��̳ʹ�� Cookies ��</a></li>
			<li><a href="Help.asp?page=usermaint#3">���ʹ��ǩ����</a></li>
			<li><a href="Help.asp?page=usermaint#4">���ʹ�ø��Ի���ͷ��</a></li>
			<li><a href="Help.asp?page=usermaint#5">��������������룬�Ҹ���ô�죿</a></li>
			<li><a href="Help.asp?page=usermaint#6">ʲô�ǡ�����Ϣ����</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu03 %>
	<div class="a2" id="center">
		<div class="a1" style="padding: 5px;">��̳ʹ��</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=using#1">��������Ե�¼��</a></li>
			<li><a href="Help.asp?page=using#2">����������˳���</a></li>
			<li><a href="Help.asp?page=using#3">��Ҫ������̳��Ӧ����ô����</a></li>
			<li><a href="Help.asp?page=using#4">�����������˷��͡�����Ϣ����</a></li>
			<li><a href="Help.asp?page=using#5">��������ȫ���Ļ�Ա��</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu04 %>
	<div class="a2" id="center">
		<div class="a1" style="padding: 5px;">��д���Ӻ��շ�����Ϣ</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=messages#1">��η��������ӣ�</a></li>
			<li><a href="Help.asp?page=messages#2">��λظ����ӣ�</a></li>
			<li><a href="Help.asp?page=messages#3">���ܹ�ɾ��������</a></li>
			<li><a href="Help.asp?page=messages#4">�����༭�Լ���������ӣ�</a></li>
			<li><a href="Help.asp?page=messages#5">�ҿɲ������ϴ�������</a></li>
			<li><a href="Help.asp?page=messages#6">����������һ��ͶƱ��</a></li>
		</div>
	</div>
	<br>
<%
End Sub

Sub Menu05 
	%>
	<div class="a2" id="center">
		<div class="a1" style="padding: 5px;">��������</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li><a href="Help.asp?page=mise#1">UBB Code��ʹ�÷�����</a></li>
			<li><a href="Help.asp?page=mise#2">��ͨ�û���γ�Ϊ������ </a></li>
			<li><a href="Help.asp?page=mise#3">�������TEAM Board ���и����Ȩ�ޣ�</a></li>
			<li><a href="Help.asp?page=mise#4">�鿴�ҵ�Ȩ��</a></li>
		</div>
	</div>
	<br><%
End Sub


Sub mise
	If Not team.UserLoginED Then 
		Call Main()
	Else
		Call Menu05
		%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">UBB Code��ʹ�÷����� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ������ʹ�� TEAM's ����--һ�� HTML ����ļ򻯰汾�����򻯶�������ʾ��ʽ�Ŀ��ơ�<br><br>
<ol type="1">
<li>[b]�������� Abc[/b] &nbsp; Ч��:<b>�������� Abc</b> �������֣�<br><br></li>
<li>[i]б������ Abc[/i] &nbsp; Ч��:<i>б������ Abc</i> ��б���֣�<br><br></li>
<li>[u]�»������� Abc[/u] &nbsp; Ч��:<u>�»������� Abc</u> ���»��ߣ�<br><br></li>
<li>[color=red]����ɫ[/color] &nbsp; Ч��:<font color="red">����ɫ</font> ���ı�������ɫ��<br><br></li>
<li>[size=3]���ִ�СΪ 3[/size] &nbsp; Ч��:<font size="3">���ִ�СΪ 3</font> ���ı����ִ�С��<br><br></li>
<li>[font=����]����Ϊ����[/font] &nbsp; Ч��:<font face"����">����Ϊ����</font> ���ı����壩<br><br></li>
<li>[align=Center]���ݾ���[/align] &nbsp; ����ʽ����λ�ã� Ч��:<br><center>���ݾ���</center><br></li>
<li>[url]http://www.team5.cn[/url] &nbsp; Ч��:<a href="http://www.team5.cn" target="_blank">http://www.team5.cn</a> ���������ӣ�<br><br></li>
<li>[url=http://www.team5.cn]TEAM's ��̳[/url] &nbsp; Ч��:<a href="http://www.TEAM5.cn" target="_blank">TEAM's ��̳</a> ���������ӣ�<br><br></li>
<li>[email]myname@mydomain.com[/email] &nbsp; Ч��:<a href="mailto:myname@mydomain.com">myname@mydomain.com</a> ��E-Mail ���ӣ�<br><br></li>
<li>[email=teamserver@163.com]TEAM's ����֧��[/email] &nbsp; Ч��:<a href="mailto:teamserver@163.com">TEAM's ����֧��</a> ��E-Mail ���ӣ�<br><br></li>
<li>[quote]TEAM Board ����TEAM Studio ��������̳���[/quote] &nbsp; ���������ݣ����ƵĴ��뻹�� [code][/code]��<br><br></li>
<li>[REPLAYVIEW]����ʺ�Ϊ: username/password[/REPLAYVIEW] &nbsp; �����ظ��������ݣ�<br>Ч��:ֻ�е�����߻ظ�����ʱ������ʾ���е����ݣ�������ʾΪ<fieldset class=textquote><legend><strong>�ظ��ɼ���</strong></legend>���������ѱ�����,���½��鿴!</fieldset><br><br></li>
<li>[money=20]����ʺ�Ϊ: username/password[/money] &nbsp; ��������������ݣ�<br>Ч��:ֻ�е�����߽�Ҹ��� 20 ��ʱ������ʾ���е����ݣ�������ʾΪ<fieldset class=textquote><legend><strong>�޽�Ǯ����</strong></legend>�����޽�Ǯ������20�ſ������!</fieldset><br><br></li>
<li>[marquee]This is sample text[/marquee] &nbsp; (����ˮƽ�ƶ���Ч����������HTML&lt;marquee&gt;��ǩ��ע�⣺����IE������¿��á�)<br><br></li>
<li>[qq]688888[/qq] &nbsp; (��ʾQQ����״̬������ͨ�������ͼ��ʹ������졣)<br><br></li>
<br>���� TEAM's ��������̳���� [img] �������ʹ��<hr noshade size="0" width="50%" color="#698CC3" align="left"><br>
<li>[img]http://www.team5.cn/images/default/logo.gif[/img] &nbsp; ������ͼ��<br>Ч��:<br><img src="images/logo.gif"> <br><br></li>
<li>[flash=480,360]http://www.team5.cn/images/banner.swf[/flash]&nbsp; ������ flash �������÷��� [img] ���ƣ�<br><br></li>
</ol>
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">��ͨ�û���γ�Ϊ������  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ��̳�İ�������Ը����ģ�����Ա���ܻ�Ҫ�������Ҫ�ﵽһ�����֣�������̳ע�ᳬ��һ��ʱ��ȡ�����Ӧ���ǳ�ʵ���š��������ˡ�����˽�ı��ʣ�ͬʱ��Ҫ��Ϥרҵ������ḻ�������õĿڱ��������ȷ���Ѿ��ﵽ���漸�㣬��ϣ�����α�վ�İ��������������Ա��ϵ��
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">�������TEAM Board ���и����Ȩ�ޣ�  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ��վ��ʹ�õ� TEAM ��̳�ǰ���ϵͳͷ�κ��û��������ֵģ����ֿ��Բο����ķ��������Լ�����Ա�����֣��������ۺ��������������ִﵽһ���ȼ�Ҫ��ʱ��ϵͳ���Զ�Ϊ����ͨ�µ�Ȩ�ޣ���������Ӧ�ȼ���־����ˣ�ӵ�нϸߵĻ������������������ڱ���̳���������Ծ�̶ȣ�ͬʱҲ��ζ���ܹ�ӵ�б������û�����ĸ߼�Ȩ�ޡ�
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">�鿴�ҵ�Ȩ��  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>�û������ƣ�<%=team.Levelname(0)%> </li>
			<li>�û�����ʾ��ʽ��<span Style='<%=team.Levelname(1)%>'> ��Ա�� </span> </li>
		</div>
		<div class="a1" style="padding: 5px;">����Ȩ��  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>���������̳��<%if Team.Group_Browse(0)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>�Ķ�Ȩ�ޣ�<%=Team.Group_Browse(1)%></li>
			<li>����鿴�û����ϣ�<%if Team.Group_Browse(2)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>�������ת�ˣ�<%if Team.Group_Browse(3)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>����ʹ��������<%if Team.Group_Browse(4)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>����ʹ��ͷ��<%if Team.Group_Browse(5)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>������û����֣�<%if Team.Group_Browse(10)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>����ʹ���ļ���<%if Team.Group_Browse(7)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>������ͶƱ��<%if Team.Group_Browse(8)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>��������<%if Team.Group_Browse(9)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>�����������⣺<%if Team.Group_Browse(20)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>�����Զ���ͷ�Σ�<%if Team.Group_Browse(11)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>����Ϣ�ռ���������<%=Team.Group_Browse(12)%></li>
		</div>
		<div class="a1" style="padding: 5px;">�������ѡ�� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>�����»��⣺<%if Team.Group_Browse(13)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>������ظ���<%if Team.Group_Browse(14)=0 then%>��ֹ<%Else%>����<%End If%></li>
			<li>�������ͶƱ��<%if Team.Group_Browse(15)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>������������<%if Team.Group_Browse(17)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>������������Ȩ�ޣ�<%if Team.Group_Browse(18)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>����ʹ�������ɫ��<%if Team.Group_Browse(19)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>����ǩ����ʹ�� UBB ���룺<%if Team.Group_Browse(21)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>����ǩ����ʹ�� [img] ���룺<%if Team.Group_Browse(22)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>���ǩ�����ȣ�<%=Team.Group_Browse(23)%> </li>
		</div>
		<div class="a1" style="padding: 5px;">�������ѡ�� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			<li>��������/�鿴������<%if Team.Group_Browse(24)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>������������<%if Team.Group_Browse(25)=0 then%>��ֹ<%Else%>����<%End If%> </li>
			<li>ÿ���ϴ�����������<%=Team.Group_Browse(26)%> </li>
			<li>��󸽼��ߴ�(KB)��<%=Team.Group_Browse(27)%> </li>
			<li>ÿ���ϴ���������������<%=Team.Group_Browse(28)%></li>
			<li>���������ͣ�<%
				If Team.Group_Browse(29)&""="" Then
					Echo team.Forum_setting(73)
				Else
					Echo Team.Group_Browse(29)
				End if
				%> </li>
		</div>
	</div><BR>
	<%
	End If
End Sub

Sub messages
	Call Menu04()
	%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">��η��������ӣ� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ����̳����У��㡰�������������ɽ��빦����ȫ�ķ������档��Ȼ��Ҳ����ʹ�ð������ġ����ٷ�������������(�����ѡ���)��ע�⣬һ����̳������Ϊ��Ҫ��¼����ܷ�����
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">��λظ����ӣ� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ���������ʱ���㡰�ظ����ӡ����ɽ��빦����ȫ�Ļظ����档��Ȼ��Ҳ����ʹ�ð������ġ����ٻظ�������ظ�(�����ѡ���)��ע�⣬һ����̳������Ϊ��Ҫ��¼����ܻظ���
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">���ܹ�ɾ�������� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ��̳������ֻ��ӵ�й���ȼ����û��ſ���ɾ�����ӡ�
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">�����༭�Լ���������ӣ� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ��������ʾ�������á��༭���Ϳ��Ա༭�Լ���������ӡ��������Աͨ����̳���ý�����������ε����ٿ��Խ��д˲���
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="5"></a>
		<div class="a1" style="padding: 5px;">�ҿɲ������ϴ������� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ���ԡ����������κ�֧���ϴ������İ���У�ͨ�������������߻ظ��ķ�ʽ�ϴ�������ֻҪ����Ȩ���㹻�����������ܳ���ϵͳ�޶��ߴ磬���ڿ������͵ķ�Χ�ڲ����ϴ���
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="6"></a>
		<div class="a1" style="padding: 5px;">����������һ��ͶƱ�� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ����������һ���ڰ���з���ͶƱ��ÿ������һ�����ܵ�ѡ����10������������ͨ���Ķ����ͶƱ��ѡ���Լ��Ĵ𰸣�ÿ��ֻ��ͶƱһ�Σ�֮�󽫲����ٶ�����ѡ���������ġ�<br><br>&nbsp; &nbsp; ����Աӵ����ʱ�رպ��޸�ͶƱѡ���Ȩ����
		</div>
	</div><BR>
	<%
End Sub


Sub using
	Call Menu03()
	%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">��������Ե�¼�� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; �������δ��¼��������Ͻǵġ���¼���������û��������룬ȷ�����ɡ������Ҫ���ֵ�¼����ѡ����Ӧ�� Cookie ʱ�䣬�ڴ�ʱ�䷶Χ�������Բ�����������������ϴεĵ�¼״̬��
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">����������˳��� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ������Ѿ���¼��������Ͻǵġ��˳�����ϵͳ����� Cookie���˳���¼״̬��
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">��Ҫ������̳��Ӧ����ô���� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; �������� <a href="search.asp">����</a>�����������Ĺؼ��ֲ�ѡ��һ����Χ���Ϳ��Լ���������Ȩ�޷�����̳�е���ص����ӡ�
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">�����������˷��͡�����Ϣ���� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ������ѵ�¼���˵��ϻ���ʾ�� <a href="Message.asp" target="_blank">���ŷ���</a>  ��Լ���������ʾ��������ʾ"����Ϣ"��ͼƬ��, ����󵯳�����Ϣ���ڣ�ͨ�����Ʒ����ʼ�һ������д���㡰���͡�����Ϣ�ͱ������Է��ռ������ˡ�����/��������̳����Ҫҳ��ʱ��ϵͳ������ʾ��/������Ϣ��
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="5"></a>
		<div class="a1" style="padding: 5px;">��������ȫ���Ļ�Ա�� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ������ͨ����� <a href="ShowBBS.asp">���а� </a> �鿴���еĻ�Ա�������ϣ�����ʵ�ֻ�Ա���ϵ����������
		</div>
	</div><BR>
	<%
End Sub



Sub usermaint
	Call Menu02()
	%>
	<div class="a2" id="center">
		<a name="1"></a>
		<div class="a1" style="padding: 5px;">�ұ���Ҫע���� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ��ȡ���ڹ���Ա������� TEAM ��̳���û���Ȩ��ѡ��������п��ܱ�����ע�����ʽ�û�������������ӡ���Ȼ����ͨ������£�������Ӧ������ʽ�û����ܷ������ͻظ��������ӡ��� <a href="Reg.asp">�������</a> ���ע���Ϊ���ǵ����û���<br><br>&nbsp; &nbsp; ǿ�ҽ�����ע�ᣬ������õ��ܶ����ο�����޷�ʵ�ֵĹ��ܡ�
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="2"></a>
		<div class="a1" style="padding: 5px;">TEAM's ��̳ʹ�� Cookies �� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; TEAM ���� Session+Cookie ��˫�ط�ʽ�����û���Ϣ����ȷ���ڸ��ֻ��������� Cookie ��ȫ�޷�ʹ�õ����������������ʹ����̳����ܡ��� Cookies ��ʹ����Ȼ����Ϊ������һϵ�еķ���ͺô����������ǿ�ҽ���������������²�Ҫ��ֹ Cookie ��Ӧ�ã�TEAM's �İ�ȫ��ƽ�ȫ����֤�������ϰ�ȫ��<br><br>&nbsp; &nbsp; �ڵ�¼ҳ���У�������ѡ�� Cookie ��¼ʱ�䣬�ڸ�ʱ�䷶Χ�����������������̳��ʼ�ձ�������һ�η���ʱ�ĵ�¼״̬��������ÿ�ζ��������롣�����ڰ�ȫ���ǣ�������ڹ��������������̳������ѡ����������̡��������뿪���������ǰѡ���˳���(<a href="Login.asp?menu=out">�������</a> �˳���̳)�Զž����ϱ��Ƿ�ʹ�õĿ��ܡ�
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="3"></a>
		<div class="a1" style="padding: 5px;">���ʹ��ǩ���� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ǩ���Ǽ�������������������С�����֣�ע��֮�����Ϳ��������Լ��ĸ���ǩ���ˡ�<br><br>&nbsp; &nbsp; <a href="EditProfile.asp?menu=index">�������</a> ���������� - �����޸ģ���ǩ����������ǩ�����֣���ȷ����Ҫ��������Ա���õ��������(����������ͼ��)������ϵͳ���Զ�ѡ������¼����ҳ�����ʾǩ��ѡ����ĵ�ǩ�������������Զ�����ʾ��
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="4"></a>
		<div class="a1" style="padding: 5px;">���ʹ�ø��Ի���ͷ�� </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ͬ���� <a href="EditProfile.asp?menu=index">�������</a>  - �����޸� �У���һ����ͷ��ѡ�ͷ������ʾ�����û��������Сͼ��ʹ��ͷ�������Ҫһ����Ȩ�ޣ����򽫲�����ʾ�������������ѯ<a href="?page=usergroup">����̳�ļ����趨</a>��
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="5"></a>
		<div class="a1" style="padding: 5px;">��������������룬�Ҹ���ô�죿 </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; TEAM �ṩ����ȡ���������ӵ� Email �ķ��񣬵����¼ҳ���е� <a href="Modification.asp">ȡ������</a> ���ܣ�����Ϊ����ȡ������ķ������͵�ע��ʱ��д�� Email �����С�������� Email ��ʧЧ���޷��յ��ż���������̳����Ա��ϵ��
		</div>
	</div><BR>
	<div class="a2" id="center">
		<a name="6"></a>
		<div class="a1" style="padding: 5px;">ʲô�ǡ�����Ϣ����  </div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			&nbsp; &nbsp; ������Ϣ������̳ע���û��佻���Ĺ��ߣ���Ϣֻ�з������ռ��˿��Կ������յ���Ϣ��ϵͳ�������������Ӧ��ʾ��������ͨ������Ϣ������ͬһ��̳�ϵ������û�����˽����ϵ��<a href="Message.asp" target="_blank">�ռ���</a> �� <a href="EditProfile.asp">�������</a> ���ṩ�˶���Ϣ���շ�����
		</div>
	</div><BR>
	<%
End Sub


Sub custom	
	Call Menu01()
	%>
	<div class="a2" id="center">
		<a name="0"></a>
		<div class="a1" style="padding: 5px;">�������ӹ���������涨</div>
		<div class="a4" style="padding: 5px;">
			<ul style="margin-top: 2px">
			�л����񹲺͹���Ϣ��ҵ���������� <br />
<br />
�����������ӹ���������涨���Ѿ�2000��10��8�յ��Ĵβ������ͨ�������跢�����Է���֮����ʩ�С�<br />
<br />
��Ϣ��ҵ�����������<br />
<br />
��һ�� Ϊ�˼�ǿ�Ի��������ӹ������(���¼�Ƶ��ӹ������)�Ĺ����淶���ӹ�����Ϣ������Ϊ��ά�����Ұ�ȫ������ȶ������Ϲ��񡢷��˺�������֯�ĺϷ�Ȩ�棬���ݡ���������Ϣ�������취���Ĺ涨���ƶ����涨��<br />
�ڶ��� ���л����񹲺͹����ڿ�չ���ӹ����������õ��ӹ��淢����Ϣ�����ñ��涨��<br />
�������涨���Ƶ��ӹ��������ָ�ڻ��������Ե��Ӳ����ơ����Ӱװ塢������̳�����������ҡ����԰�Ƚ�����ʽΪ�����û��ṩ��Ϣ������������Ϊ��<br />
������ ���ӹ�������ṩ�߿�չ������Ӧ�����ط��ɡ����棬��ǿ��ҵ���ɣ�������Ϣ��ҵ����ʡ����������ֱϽ�е��Ź�������������й����ܲ�������ʵʩ�ļල��顣 <br />
������ �����û�ʹ�õ��ӹ������ϵͳ��Ӧ�����ط��ɡ����棬��������������Ϣ����<br />
������ ���»�������Ϣ�����⿪չ���ӹ������ģ�Ӧ������ʡ����������ֱϽ�е��Ź������������Ϣ��ҵ�����뾭Ӫ�Ի�������Ϣ������ɻ��߰���Ǿ�Ӫ�Ի�������Ϣ���񱸰�ʱ�����ר���������ר�����<br />
����ʡ����������ֱϽ�е��Ź������������Ϣ��ҵ���������������ģ�Ӧ���ڹ涨ʱ������ͬ��������Ϣ����һ��������׼���߱��������ھ�Ӫ���֤�򱸰��ļ���ר��ע���������������ģ�������׼���߲��豸��������֪ͨ�����˲�˵�����ɡ� <br />
������ ��չ���ӹ�����񣬳�Ӧ�����ϡ���������Ϣ�������취���涨�������⣬��Ӧ���߱�����������<br />
����(һ)��ȷ���ĵ��ӹ������������Ŀ��<br />
����(��)�����Ƶĵ��ӹ���������<br />
����(��)�е��ӹ������ȫ���ϴ�ʩ�����������û��Ǽǳ��������û���Ϣ��ȫ�����ƶȡ�����������ʩ��<br />
����(��)����Ӧ��רҵ������Ա�ͼ�����Ա���ܹ��Ե��ӹ������ʵʩ��Ч����<br />
������ ��ȡ�þ�Ӫ��ɻ��������б��������Ļ�������Ϣ�����ṩ�ߣ��⿪չ���ӹ������ģ�Ӧ����ԭ��ɻ��߱����������ר���������ר�����<br />
����ʡ����������ֱϽ�е��Ź������������Ϣ��ҵ����Ӧ�����յ�ר���������ר�������֮����60���ڽ��������ϡ��������������ģ�������׼���߱��������ھ�Ӫ���֤�򱸰��ļ���ר��ע���������������ģ�������׼���߲��豸��������֪ͨ�����˲�˵�����ɡ�<br />
�ڰ��� δ��ר����׼����ר����������κε�λ���߸��˲������Կ�չ���ӹ������<br />
�ھ��� �κ��˲����ڵ��ӹ������ϵͳ�з���������������֮һ����Ϣ��<br />
����(һ)�����ܷ���ȷ���Ļ���ԭ��ģ�<br />
����(��)Σ�����Ұ�ȫ��й¶�������ܣ��߸�������Ȩ���ƻ�����ͳһ�ģ�<br />
����(��)�𺦹�������������ģ�<br />
����(��)ɿ�������ޡ��������ӣ��ƻ������Ž�ģ�<br />
����(��)�ƻ������ڽ����ߣ�����а�̺ͷ⽨���ŵģ�<br />
����(��)ɢ��ҥ�ԣ�������������ƻ�����ȶ��ģ�<br />
����(��)ɢ�����ࡢɫ�顢�Ĳ�����������ɱ���ֲ����߽�������ģ�<br />
����(��)������߷̰����ˣ��ֺ����˺Ϸ�Ȩ��ģ�<br />
����(��)���з��ɡ����������ֹ���������ݵģ� <br />
��ʮ�� ���ӹ�������ṩ��Ӧ���ڵ��ӹ������ϵͳ������λ�ÿ��ؾ�Ӫ���֤��Ż��߱�����š����ӹ��������򣬲���ʾ�����û�������Ϣ��Ҫ�е��ķ������Ρ�<br />
��ʮһ�� ���ӹ�������ṩ��Ӧ�����վ���׼���߱�����������Ŀ�ṩ���񣬲��ó���������������Ŀ�ṩ����<br />
��ʮ���� ���ӹ�������ṩ��Ӧ���������û��ĸ�����Ϣ���ܣ�δ�������û�ͬ�ⲻ��������й¶�����������й涨�ĳ��⡣<br />
��ʮ���� ���ӹ�������ṩ�߷�������ӹ������ϵͳ�г����������ڱ��취�ھ������е���Ϣ����֮һ�ģ�Ӧ������ɾ���������йؼ�¼����������йػ��ر��档 <br />
��ʮ���� ���ӹ�������ṩ��Ӧ����¼�ڵ��ӹ������ϵͳ�з�������Ϣ���ݼ��䷢��ʱ�䡢��������ַ������������¼����Ӧ������60�գ����ڹ����йػ���������ѯʱ�������ṩ�� <br />
��ʮ���� ��������������ṩ��Ӧ����¼�����û�������ʱ�䡢�û��ʺš���������ַ�������������е绰�������Ϣ����¼����Ӧ����60�գ����ڹ����йػ���������ѯʱ�������ṩ��<br />
��ʮ���� Υ�����涨�ڰ�������ʮһ���Ĺ涨�����Կ�չ���ӹ��������߳�������׼���߱����������Ŀ�ṩ���ӹ������ģ����ݡ���������Ϣ�������취����ʮ�����Ĺ涨������<br />
��ʮ���� �ڵ��ӹ������ϵͳ�з������涨�ھ����涨����Ϣ����֮һ�ģ����ݡ���������Ϣ�������취���ڶ�ʮ���Ĺ涨������<br />
��ʮ���� Υ�����涨��ʮ���Ĺ涨��δ���ؾ�Ӫ���֤��Ż��߱�����š�δ���ص��ӹ������������δ�������û���������Ϣ��Ҫ�е�����������ʾ�ģ����ݡ���������Ϣ�������취���ڶ�ʮ�����Ĺ涨������ <br />
��ʮ���� Υ�����涨��ʮ�����Ĺ涨��δ�������û�ͬ�⣬�����˷Ƿ�й¶�����û�������Ϣ�ģ���ʡ����������ֱϽ�е��Ź����������������������û�����𺦻�����ʧ�ģ������е��������Ρ� <br />
�ڶ�ʮ�� δ���б��涨��ʮ��������ʮ��������ʮ�����涨������ģ����ݡ���������Ϣ�������취���ڶ�ʮһ�����ڶ�ʮ�����Ĺ涨������<br />
�ڶ�ʮһ�� �ڱ��涨ʩ����ǰ�ѿ�չ���ӹ������ģ�Ӧ���Ա��涨ʩ��֮����60���ڣ����ձ��涨����ר���������ר���������<br />
�ڶ�ʮ���� ���涨�Է���֮����ʩ�С�
</div>
<a name="1"></a>
<div class="a1" style="padding: 5px;">��������Ϣ�������취</div>
<div class="a4" style="padding: 5px;">
<ul style="margin-top: 2px">
�л����񹲺͹�����Ժ���292�ţ�<br />
&nbsp; &nbsp; ��һ�� Ϊ�˹淶��������Ϣ�������ٽ���������Ϣ���񽡿�����չ���ƶ����취�� <br />
&nbsp; &nbsp; �ڶ��� ���л����񹲺͹����ڴ��»�������Ϣ�������������ر��취�� <br />
&nbsp; &nbsp; ���취���ƻ�������Ϣ������ָͨ���������������û��ṩ��Ϣ�ķ����� <br />
&nbsp; &nbsp; ������ ��������Ϣ�����Ϊ��Ӫ�ԺͷǾ�Ӫ�����ࡣ <br />
&nbsp; &nbsp; ��Ӫ�Ի�������Ϣ������ָͨ���������������û��г��ṩ��Ϣ������ҳ�����ȷ����� <br />
&nbsp; &nbsp; �Ǿ�Ӫ�Ի�������Ϣ������ָͨ���������������û��޳��ṩ���й����ԡ���������Ϣ�ķ����� <br />
&nbsp; &nbsp; ������ ���ҶԾ�Ӫ�Ի�������Ϣ����ʵ������ƶȣ��ԷǾ�Ӫ�Ի�������Ϣ����ʵ�б����ƶȡ� <br />
&nbsp; &nbsp; δȡ����ɻ���δ���б��������ģ����ô��»�������Ϣ���� <br />
&nbsp; &nbsp; ������ �������š����桢������ҽ�Ʊ�����ҩƷ��ҽ����е�Ȼ�������Ϣ�������շ��ɡ����������Լ������йع涨�뾭�й����ܲ������ͬ��ģ������뾭Ӫ��ɻ������б�������ǰ��Ӧ���������й����ܲ������ͬ�⡣ <br />
&nbsp; &nbsp; ������ ���¾�Ӫ�Ի�������Ϣ���񣬳�Ӧ�����ϡ��л����񹲺͹������������涨��Ҫ���⣬��Ӧ���߱����������� <br />
&nbsp; &nbsp; ��һ����ҵ��չ�ƻ�����ؼ��������� <br />
&nbsp; &nbsp; �������н�ȫ����������Ϣ��ȫ���ϴ�ʩ��������վ��ȫ���ϴ�ʩ����Ϣ��ȫ���ܹ����ƶȡ��û���Ϣ��ȫ�����ƶȣ� <br />
&nbsp; &nbsp; ������������Ŀ���ڱ��취�������涨��Χ�ģ���ȡ���й����ܲ���ͬ����ļ��� <br />
&nbsp; &nbsp; ������ ���¾�Ӫ�Ի�������Ϣ����Ӧ����ʡ����������ֱϽ�е��Ź���������߹���Ժ��Ϣ��ҵ���ܲ����������������Ϣ������ֵ����ҵ��Ӫ���֤�����¼�ƾ�Ӫ���֤���� <br />
&nbsp; &nbsp; ʡ����������ֱϽ�е��Ź���������߹���Ժ��Ϣ��ҵ���ܲ���Ӧ�����յ�����֮����60���������ϣ�������׼���߲�����׼�ľ�����������׼�ģ��䷢��Ӫ���֤��������׼�ģ�Ӧ������֪ͨ�����˲�˵�����ɡ� <br />
&nbsp; &nbsp; ������ȡ�þ�Ӫ���֤��Ӧ���־�Ӫ���֤����ҵ�Ǽǻ��ذ���Ǽ������� <br />
&nbsp; &nbsp; �ڰ��� ���·Ǿ�Ӫ�Ի�������Ϣ����Ӧ����ʡ����������ֱϽ�е��Ź���������߹���Ժ��Ϣ��ҵ���ܲ��Ű�����������������ʱ��Ӧ���ύ���в��ϣ� <br />
&nbsp; &nbsp; ��һ�����쵥λ����վ�����˵Ļ�������� <br />
&nbsp; &nbsp; ��������վ��ַ�ͷ�����Ŀ�� <br />
&nbsp; &nbsp; ������������Ŀ���ڱ��취�������涨��Χ�ģ���ȡ���й����ܲ��ŵ�ͬ���ļ��� <br />
&nbsp; &nbsp; ʡ����������ֱϽ�е��Ź�������Ա���������ȫ�ģ�Ӧ�����Ա�������š� <br />
&nbsp; &nbsp; �ھ��� ���»�������Ϣ�����⿪����ӹ������ģ�Ӧ�������뾭Ӫ�Ի�������Ϣ������ɻ��߰���Ǿ�Ӫ�Ի�������Ϣ���񱸰�ʱ�����չ����йع涨���ר���������ר����� <br />
&nbsp; &nbsp; ��ʮ�� ʡ����������ֱϽ�е��Ź�������͹���Ժ��Ϣ��ҵ���ܲ���Ӧ������ȡ�þ�Ӫ���֤���������б��������Ļ�������Ϣ�����ṩ�������� <br />
&nbsp; &nbsp; ��ʮһ�� ��������Ϣ�����ṩ��Ӧ�����վ���ɻ��߱�������Ŀ�ṩ���񣬲��ó�������ɻ��߱�������Ŀ�ṩ���� <br />
&nbsp; &nbsp; �Ǿ�Ӫ�Ի�������Ϣ�����ṩ�߲��ô����г����� <br />
&nbsp; &nbsp; ��������Ϣ�����ṩ�߱��������Ŀ����վ��ַ������ģ�Ӧ����ǰ30����ԭ��ˡ���֤���߱������ذ����������� <br />
&nbsp; &nbsp; ��ʮ���� ��������Ϣ�����ṩ��Ӧ��������վ��ҳ������λ�ñ����侭Ӫ���֤��Ż��߱�����š� <br />
&nbsp; &nbsp; ��ʮ���� ��������Ϣ�����ṩ��Ӧ���������û��ṩ���õķ��񣬲���֤���ṩ����Ϣ���ݺϷ��� <br />
&nbsp; &nbsp; ��ʮ���� �������š������Լ����ӹ���ȷ�����Ŀ�Ļ�������Ϣ�����ṩ�ߣ�Ӧ����¼�ṩ����Ϣ���ݼ��䷢��ʱ�䡢��������ַ������������������������ṩ��Ӧ����¼�����û�������ʱ�䡢�û��ʺš���������ַ�������������е绰�������Ϣ�� <br />
&nbsp; &nbsp; ��������Ϣ�����ṩ�ߺͻ�������������ṩ�ߵļ�¼����Ӧ������60�գ����ڹ����йػ���������ѯʱ�������ṩ�� <br />
&nbsp; &nbsp; ��ʮ���� ��������Ϣ�����ṩ�߲������������ơ����������������������ݵ���Ϣ�� <br />
&nbsp; &nbsp; ��һ�������ܷ���ȷ���Ļ���ԭ��ģ� <br />
&nbsp; &nbsp; ������Σ�����Ұ�ȫ��й¶�������ܣ��߸�������Ȩ���ƻ�����ͳһ�ģ� <br />
&nbsp; &nbsp; �������𺦹�������������ģ� <br />
&nbsp; &nbsp; ���ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ� <br />
&nbsp; &nbsp; ���壩�ƻ������ڽ����ߣ�����а�̺ͷ⽨���ŵģ� <br />
&nbsp; &nbsp; ������ɢ��ҥ�ԣ�������������ƻ�����ȶ��ģ� <br />
&nbsp; &nbsp; ���ߣ�ɢ�����ࡢɫ�顢�Ĳ�����������ɱ���ֲ����߽�������ģ� <br />
&nbsp; &nbsp; ���ˣ�������߷̰����ˣ��ֺ����˺Ϸ�Ȩ��ģ� <br />
&nbsp; &nbsp; ���ţ����з��ɡ����������ֹ���������ݵġ� <br />
&nbsp; &nbsp; ��ʮ���� ��������Ϣ�����ṩ�߷�������վ�������Ϣ�������ڱ��취��ʮ������������֮һ�ģ�Ӧ������ֹͣ���䣬�����йؼ�¼����������йػ��ر��档 <br />
&nbsp; &nbsp; ��ʮ���� ��Ӫ�Ի�������Ϣ�����ṩ�������ھ��ھ������л���ͬ���̺��ʡ�������Ӧ�����Ⱦ�����Ժ��Ϣ��ҵ���ܲ������ͬ�⣻���У�����Ͷ�ʵı���Ӧ�������йط��ɡ���������Ĺ涨�� <br />
&nbsp; &nbsp; ��ʮ���� ����Ժ��Ϣ��ҵ���ܲ��ź�ʡ����������ֱϽ�е��Ź�������������Ի�������Ϣ����ʵʩ�ල���� <br />
&nbsp; &nbsp; ���š����桢������������ҩƷ�ල����������������͹��������Ұ�ȫ���й����ܲ��ţ��ڸ���ְ��Χ�������Ի�������Ϣ����ʵʩ�ල���� <br />
&nbsp; &nbsp; ��ʮ���� Υ�����취�Ĺ涨��δȡ�þ�Ӫ���֤�����Դ��¾�Ӫ�Ի�������Ϣ���񣬻��߳�����ɵ���Ŀ�ṩ����ģ���ʡ����������ֱϽ�е��Ź�������������ڸ�������Υ�����õģ�û��Υ�����ã���Υ������3������5�����µķ��û��Υ�����û���Υ�����ò���5��Ԫ�ģ���10��Ԫ����100��Ԫ���µķ��������صģ�����ر���վ�� <br />
&nbsp; &nbsp; Υ�����취�Ĺ涨��δ���б������������Դ��·Ǿ�Ӫ�Ի�������Ϣ���񣬻��߳�����������Ŀ�ṩ����ģ���ʡ����������ֱϽ�е��Ź�������������ڸ������ܲ������ģ�����ر���վ�� <br />
&nbsp; &nbsp; �ڶ�ʮ�� ���������ơ��������������취��ʮ������������֮һ����Ϣ�����ɷ���ģ�����׷���������Σ��в����ɷ���ģ��ɹ������ء����Ұ�ȫ�������ա��л����񹲺͹��ΰ������������������������Ϣ�������������ȫ��������취�����йط��ɡ���������Ĺ涨���Դ������Ծ�Ӫ�Ի�������Ϣ�����ṩ�ߣ����ɷ�֤��������ͣҵ����ֱ��������Ӫ���֤��֪ͨ��ҵ�Ǽǻ��أ��ԷǾ�Ӫ�Ի�������Ϣ�����ṩ�ߣ����ɱ�������������ʱ�ر���վֱ���ر���վ�� <br />
&nbsp; &nbsp; �ڶ�ʮһ�� δ���б��취��ʮ�����涨������ģ���ʡ����������ֱϽ�е��Ź���������������������صģ�����ͣҵ���ٻ�����ʱ�ر���վ�� <br />
&nbsp; &nbsp; �ڶ�ʮ���� Υ�����취�Ĺ涨��δ������վ��ҳ�ϱ����侭Ӫ���֤��Ż��߱�����ŵģ���ʡ����������ֱϽ�е��Ź�����������������5000Ԫ����5��Ԫ���µķ�� <br />
&nbsp; &nbsp; �ڶ�ʮ���� Υ�����취��ʮ�����涨������ģ���ʡ����������ֱϽ�е��Ź���������������������صģ��Ծ�Ӫ�Ի�������Ϣ�����ṩ�ߣ����ɷ�֤���ص�����Ӫ���֤���ԷǾ�Ӫ�Ի�������Ϣ�����ṩ�ߣ����ɱ�����������ر���վ�� <br />
&nbsp; &nbsp; �ڶ�ʮ���� ��������Ϣ�����ṩ������ҵ���У�Υ���������ɡ�����ģ������š����桢������������ҩƷ�ල����͹�������������й����ܲ��������йط��ɡ�����Ĺ涨������ <br />
&nbsp; &nbsp; �ڶ�ʮ���� ���Ź�������������й����ܲ��ż��乤����Ա�����ְ�ء�����ְȨ����˽��ף����ڶԻ�������Ϣ����ļල����������غ�������ɷ���ģ�����׷���������Σ��в����ɷ���ģ���ֱ�Ӹ����������Ա������ֱ��������Ա�������轵������ְֱ���������������֡� <br />
&nbsp; &nbsp; �ڶ�ʮ���� �ڱ��취����ǰ���»�������Ϣ����ģ�Ӧ���Ա��취����֮����60�������ձ��취���йع涨�����й������� <br />
&nbsp; &nbsp; �ڶ�ʮ���� ���취�Թ���֮����ʩ�С�
		</div>
	</div>
	<br>
<%
End Sub

%>
