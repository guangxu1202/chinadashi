<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="��ʵ��������,�����˴�ִ��̤ʵ��������ʵ��ҵ�������޹�˾�ĺ��Ĳ�ҵ�Ƿ��ز�����,�ѿ������ɵ���ɽׯ�����俥�����������⡢���ֺ�ɽ��סլС��������30����ƽ����">
<meta name="keywords" content="��ʵ���ţ���ʵ�����ֺ�ɽ������¥�̣��������ز������ز�" >
<title>��ʵ����-��������-�����Ǽ�</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
<script type="text/javascript" src="dhtml.js"></script>
<%if request("msg")=1 then%>
	<script>alert("��д�ɹ�����ȴ�����������ϵ��")</script>
<%end if%>
<script language="javascript">

//var select1_len = document.frm.s1.options.length;
var select2 = new Array(4);


for (i=0; i<5; i++) 
{
select2[i] = new Array();
}
//�������ѡ��
select2[0][0] = new Option("��ѡ��", "δѡ��");
select2[0][1] = new Option("", "δѡ��");
select2[0][2] = new Option("", "δѡ��");
select2[0][3] = new Option("", "δѡ��");
select2[0][4] = new Option("", "δѡ��");
select2[0][5] = new Option("", "δѡ��");

select2[1][0] = new Option("����ɽׯС��", "����ɽׯС��");
select2[1][1] = new Option("���俥��С��", "���俥��С��");
select2[1][2] = new Option("��������С��", "��������С��");
select2[1][3] = new Option("���ֺ�ɽС��", "���ֺ�ɽС��");
select2[1][4] = new Option("����ӳɽС��", "����ӳɽС��");
select2[1][5] = new Option("�����硤��", "�����硤��");

select2[2][0] = new Option("���Ż�԰", "���Ż�԰");
select2[2][1] = new Option("", "δѡ��");
select2[2][2] = new Option("", "δѡ��");
select2[2][3] = new Option("", "δѡ��");
select2[2][4] = new Option("", "δѡ��");
select2[2][5] = new Option("", "δѡ��");

select2[3][0] = new Option("�󻪡�ˮ����ۡ", "�󻪡�ˮ����ۡ");
select2[3][1] = new Option("", "δѡ��");
select2[3][2] = new Option("", "δѡ��");
select2[3][3] = new Option("", "δѡ��");
select2[3][4] = new Option("", "δѡ��");
select2[3][5] = new Option("", "δѡ��");


function redirec(x)
{
var temp = document.form1.zone; 

for (i=0;i<select2[x].length;i++)
{
    temp.options[i]=new Option(select2[x][i].text,select2[x][i].value);
}
temp.options[0].selected=true;

}

function checkForm(){
	if(form1.Uname.value==''){
		alert("��������Ϊ��")
		form1.Uname.focus();
		return false;
	}
	if(form1.Utel.value==''){
		alert("��ϵ�绰����Ϊ��")
		form1.Utel.focus();
		return false;
	}
	if(form1.Umail.value==''){
		alert("�������䲻��Ϊ��")
		form1.Umail.focus();
		return false;
	}
	if(form1.Utime.value==''){
		alert("�ƻ�����ʱ�䲻��Ϊ��")
		form1.Utime.focus();
		return false;
	}
	document.form1.submit();
}
</script>
</head>

<body  onload=init();>
<div id="DsTop">
	<img src="images/logo.gif" />
    <ul>
    	<li><a href="default.asp" id="menu1"><img src="images/top1.jpg" /></a></li>
        <li><a href="news.asp" id="menu2"><img src="images/top2.jpg" /></a></li>
        <li><a href="about.asp" id="menu3"><img src="images/top3.jpg" /></a></li>
        <li><a href="brand_list.asp" id="menu4"><img src="images/top4.jpg" /></a></li>
        <li><a href="estate.asp" id="menu5"><img src="images/top5.jpg" /></a></li>
        <li><a href="teams/default.asp" id="menu8"><img src="images/top8.jpg" /></a></li>
        <li><a href="join.asp" id="menu6"><img src="images/top6.jpg" /></a></li>
        <li><a href="contact.asp" id="menu7"><img src="images/top7.jpg" /></a></li>
    </ul>
</div>
<script type="text/javascript" src="dropdown_initialize.js"></script>
<div id="Main">
	<div id="Main_left" class="fastlink">
    	<em><img src="images/fastlink.jpg" /></em>
        <ul>
        	<li><a href="link_scrx.asp">�۳�����</a></li>
            <li><a href="link_gfdj.asp">�����Ǽ�</a></li>
            <li><a href="link_vip.asp">��ʵVIP</a></li>
            <li><a href="link_cpsc.asp">��Ʒ�Ӵ�</a></li>
          <li><a href="link_kfrx.asp">�ͷ�����</a></li>
            <li><a href="link_sqfw.asp">��������</a></li>
        </ul>
  </div>
    <div id="Main_right">
    	<img src="images/fast_gfdj.jpg" class="guide" />
        <div><img src="images/fast1.jpg" /></div>
        <h4  class="fast_h4">����������������Ϣ��ʵ���ؿͻ��������Ľ������մ����������ʱ��������������ϵ��</h4>
      <div class="Main_right_content">
        	<table width="90%" border="0" cellspacing="0" cellpadding="0">
             <form id="form1" name="form1" method="post" action="link_gfdj_sql.asp">
              <tr>
                <td width="100" height="30" align="right">����</td>
                <td width="20">&nbsp;</td>
                <td><input type="text" name="Uname" id="Uname" />
                  *</td>
              </tr>
              <tr>
                <td height="30" align="right">�Ա�</td>
                <td>&nbsp;</td>
                <td><input name="Uxb" type="radio" id="radio" value="��" checked="checked" />
                  ��
                  <input type="radio" name="Uxb" id="radio2" value="Ů" />
                  Ů</td>
              </tr>
              <tr>
                <td height="30" align="right">��ϵ�绰</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Utel" id="Utel" />
                *</td>
              </tr>
              <tr>
                <td height="30" align="right">��ַ</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Udz" id="Udz" /></td>
              </tr>
              <tr>
                <td height="30" align="right">�ʱ�</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Uyb" id="Uyb" /></td>
              </tr>
              <tr>
                <td height="30" align="right">��������</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Umail" id="Umail" />
                *</td>
              </tr>
              <tr>
                <td height="30" align="right">�ƻ�����ҵ̬</td>
                <td>&nbsp;</td>
                <td>
                <select name="Place" onChange="redirec(this.options.selectedIndex)">
                  <option selected value="0">��ѡ��</option>
                  <option value="1">����</option>
                  <option value="2">����</option>
                  <option value="3">�Ĵ�</option>
                </select>
                  <select name="zone" >
                    <option value="δѡ��" selected>��ѡ��</option>
                  </select>

                </td>
              </tr>
              <tr>
                <td height="30" align="right">�ƻ�����ʱ��</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Utime" id="Utime" />
                *</td>
              </tr>
              <tr>
                <td height="30" align="right">��ʵҵ��</td>
                <td>&nbsp;</td>
                <td><input name="Udsyz" type="radio" id="radio3" value="��" checked="checked" />
����
  <input type="radio" name="Udsyz" id="radio4" value="��" /> 
  ��
</td>
              </tr>
              <tr>
                <td height="30" align="right">��������</td>
                <td>&nbsp;</td>
                <td><input name="Ugfjl" type="radio" id="radio5" value="��һ�ι���" checked="checked" />
                  ��һ�ι���
                    <input type="radio" name="Ugfjl" id="radio6" value="��ι���" />
��ι���</td>
              </tr>
              <tr>
                <td height="30" align="right">��ע</td>
                <td>&nbsp;</td>
                <td><textarea name="Ubz" id="Ubz" cols="45" rows="5"></textarea></td>
              </tr>
              <tr>
                <td height="30" align="right">&nbsp;</td>
                <td>&nbsp;</td>
                <td>
                <img src="images/fast_tj.gif" style="cursor:pointer;" onclick="checkForm()"  />          &nbsp;&nbsp;&nbsp;&nbsp;<input type="image" src="images/fast_cz.gif" onclick="javascript:form1.reset()"/>
                <input name="act" type="hidden" id="act" value="addj" /></td>
              </tr>
              </form>
        </table>

           
          
        </div>
  </div>
</div>
<div id="bottom">
	<div id="bottom_content">
    <select>
    	<option selected="selected">----------��������----------</option>
    </select>
    <em>
    	<a href="#">��ϵ���� |</a>
        <a href="#">����ͳ�� |</a>
        <a href="#">��վ��ͼ |</a>
        <a href="#">��������</a>
    </em>
    </div>
</div>
