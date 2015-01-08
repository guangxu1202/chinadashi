 <!--#include file="error.asp"-->
<style type="text/css">
#lefttab { background-image:url(../images/bg_000.jpg);
           background-repeat:repeat-y; color: #FFFFFF;}
#lefttab A { text-decoration:none; color: #FFFF79;}
.write { color:#FFFFFF;}
.list_title SPAN {
	FONT-WEIGHT: bold;
	LEFT: 8px;
	POSITION: relative;
	TOP: 2px;
	visibility: visible;
}
body { margin:0; padding:0px;}
.STYLE3 {color: #000000}
.STYLE4 {color: #003300}
</style>

<table width="100%" height="100%" border="0" align="right" cellpadding="0" cellspacing="0"  id="lefttab">
  <TR vAlign=top> 
    <TD height=450 align="center" ><table cellspacing="0" cellpadding="0" width="201" align="center" class="left">
      <tbody>
        <tr style="CURSOR: hand">
          <td width="201" height="60" align="center" valign="middle" background="images/nav_head.jpg"><span class="write">欢迎您：<%=session("love_uname")%></span></td>
        </tr>
      </tbody>
    </table>
      <table cellspacing="0" cellpadding="0" width="158" align="center"  class="left">
        <tbody>
          <tr style="CURSOR: hand">
            <td 
            height="25" align="center" 
			onclick="javascript:window.location.href='back_up.asp?mm=1'"
         background="images/title_show.JPG" class="list_title" id="list1"><span class="STYLE4">数据库备份与恢复</span> </td>
          </tr>
        </tbody>
      </table>
      <table cellspacing="0" cellpadding="0" width="158" align="center"  class="left">
        <tbody>
          <tr style="CURSOR: hand">
            <td 
            height="25" align="center" 
			onclick="javascript:window.location.href='../ART_BACK_ADMIN/user_main.asp?mm=1'"
         background="images/title_show.JPG" class="list_title" id="list1"><span class="STYLE3">网站管理员</span> </td>
          </tr>
        </tbody>
      </table>
      <table cellspacing="0" cellpadding="0" width="158" align="center"  class="left">
        <tbody>
          <tr style="CURSOR: hand">
            <td 
            height="25" align="center" 
			onclick="javascript:window.location.href='../ART_BACK_ADMIN/pwd_edit.asp?mm=1'"
         background="images/title_show.JPG" class="list_title" id="list1"><span class="STYLE3">密码修改</span> </td>
          </tr>
        </tbody>
      </table>
      <table cellspacing="0" cellpadding="0" width="158" align="center"  class="left">
        <tbody>
          <tr style="CURSOR: hand">
            <td 
            height="25" align="center" 
			onclick="javascript:window.location.href='../ART_BACK_ADMIN/login.asp?ex=1'"
         background="images/title_show.JPG" class="list_title" id="list1"><span class="STYLE3">安全退出</span> </td>
          </tr>
        </tbody>
      </table></TD>
  </TR>
  
</table>

