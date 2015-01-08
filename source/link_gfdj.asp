<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="大实集团理念,有容乃大、执诚踏实。大连大实企业集团有限公司的核心产业是房地产开发,已开发建成叠翠山庄、叠翠骏景、泊林阳光、泊林和山等住宅小区，共计30多万平方米">
<meta name="keywords" content="大实集团，大实，泊林和山，大连楼盘，大连房地产，房地产" >
<title>大实集团-快速链接-购房登记</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/banner.js"></script>
<script type="text/javascript" src="dhtml.js"></script>
<%if request("msg")=1 then%>
	<script>alert("填写成功，请等待我们与您联系！")</script>
<%end if%>
<script language="javascript">

//var select1_len = document.frm.s1.options.length;
var select2 = new Array(4);


for (i=0; i<5; i++) 
{
select2[i] = new Array();
}
//定义基本选项
select2[0][0] = new Option("请选择", "未选择");
select2[0][1] = new Option("", "未选择");
select2[0][2] = new Option("", "未选择");
select2[0][3] = new Option("", "未选择");
select2[0][4] = new Option("", "未选择");
select2[0][5] = new Option("", "未选择");

select2[1][0] = new Option("叠翠山庄小区", "叠翠山庄小区");
select2[1][1] = new Option("叠翠骏景小区", "叠翠骏景小区");
select2[1][2] = new Option("泊林阳光小区", "泊林阳光小区");
select2[1][3] = new Option("泊林和山小区", "泊林和山小区");
select2[1][4] = new Option("泊林映山小区", "泊林映山小区");
select2[1][5] = new Option("海·风·景", "海·风·景");

select2[2][0] = new Option("名雅花园", "名雅花园");
select2[2][1] = new Option("", "未选择");
select2[2][2] = new Option("", "未选择");
select2[2][3] = new Option("", "未选择");
select2[2][4] = new Option("", "未选择");
select2[2][5] = new Option("", "未选择");

select2[3][0] = new Option("大华·水岸福邸", "大华·水岸福邸");
select2[3][1] = new Option("", "未选择");
select2[3][2] = new Option("", "未选择");
select2[3][3] = new Option("", "未选择");
select2[3][4] = new Option("", "未选择");
select2[3][5] = new Option("", "未选择");


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
		alert("姓名不能为空")
		form1.Uname.focus();
		return false;
	}
	if(form1.Utel.value==''){
		alert("联系电话不能为空")
		form1.Utel.focus();
		return false;
	}
	if(form1.Umail.value==''){
		alert("电子邮箱不能为空")
		form1.Umail.focus();
		return false;
	}
	if(form1.Utime.value==''){
		alert("计划购买时间不能为空")
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
        	<li><a href="link_scrx.asp">售场热线</a></li>
            <li><a href="link_gfdj.asp">购房登记</a></li>
            <li><a href="link_vip.asp">大实VIP</a></li>
            <li><a href="link_cpsc.asp">产品视窗</a></li>
          <li><a href="link_kfrx.asp">客服热线</a></li>
            <li><a href="link_sqfw.asp">社区服务</a></li>
        </ul>
  </div>
    <div id="Main_right">
    	<img src="images/fast_gfdj.jpg" class="guide" />
        <div><img src="images/fast1.jpg" /></div>
        <h4  class="fast_h4">购房意向表（本表格信息大实各地客户服务中心将认真收存管理，待合适时机会主动与您联系）</h4>
      <div class="Main_right_content">
        	<table width="90%" border="0" cellspacing="0" cellpadding="0">
             <form id="form1" name="form1" method="post" action="link_gfdj_sql.asp">
              <tr>
                <td width="100" height="30" align="right">姓名</td>
                <td width="20">&nbsp;</td>
                <td><input type="text" name="Uname" id="Uname" />
                  *</td>
              </tr>
              <tr>
                <td height="30" align="right">性别</td>
                <td>&nbsp;</td>
                <td><input name="Uxb" type="radio" id="radio" value="男" checked="checked" />
                  男
                  <input type="radio" name="Uxb" id="radio2" value="女" />
                  女</td>
              </tr>
              <tr>
                <td height="30" align="right">联系电话</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Utel" id="Utel" />
                *</td>
              </tr>
              <tr>
                <td height="30" align="right">地址</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Udz" id="Udz" /></td>
              </tr>
              <tr>
                <td height="30" align="right">邮编</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Uyb" id="Uyb" /></td>
              </tr>
              <tr>
                <td height="30" align="right">电子邮箱</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Umail" id="Umail" />
                *</td>
              </tr>
              <tr>
                <td height="30" align="right">计划购买业态</td>
                <td>&nbsp;</td>
                <td>
                <select name="Place" onChange="redirec(this.options.selectedIndex)">
                  <option selected value="0">请选择</option>
                  <option value="1">大连</option>
                  <option value="2">沈阳</option>
                  <option value="3">四川</option>
                </select>
                  <select name="zone" >
                    <option value="未选择" selected>请选择</option>
                  </select>

                </td>
              </tr>
              <tr>
                <td height="30" align="right">计划购买时间</td>
                <td>&nbsp;</td>
                <td><input type="text" name="Utime" id="Utime" />
                *</td>
              </tr>
              <tr>
                <td height="30" align="right">大实业主</td>
                <td>&nbsp;</td>
                <td><input name="Udsyz" type="radio" id="radio3" value="否" checked="checked" />
不是
  <input type="radio" name="Udsyz" id="radio4" value="是" /> 
  是
</td>
              </tr>
              <tr>
                <td height="30" align="right">购房经历</td>
                <td>&nbsp;</td>
                <td><input name="Ugfjl" type="radio" id="radio5" value="第一次购房" checked="checked" />
                  第一次购房
                    <input type="radio" name="Ugfjl" id="radio6" value="多次购房" />
多次购房</td>
              </tr>
              <tr>
                <td height="30" align="right">备注</td>
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
    	<option selected="selected">----------友情链接----------</option>
    </select>
    <em>
    	<a href="#">联系我们 |</a>
        <a href="#">在线统计 |</a>
        <a href="#">网站地图 |</a>
        <a href="#">法律声明</a>
    </em>
    </div>
</div>
