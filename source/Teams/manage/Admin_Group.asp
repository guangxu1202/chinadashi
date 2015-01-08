<!--#include file="../conn.asp"-->
<!--#include file="const.asp"-->
<%
Dim ii,ID
Dim Admin_Class
Call Master_Us()
Header()
ii=0:
ID=CID(Request("ID"))
Admin_Class=",5,"
Call Master_Se()
team.SaveLog ("分组与级别 [包括：管理组 ，用户组 ] ")
Select Case Request("Action")
	Case "Editmodel_1"
		Call Editmodel_1
	Case "Editmodel_2"
		Call Editmodel_2
	Case "IsuserGroup"
		Call IsuserGroup
	Case "EditUserGroup"
		Call EditUserGroup
	Case "EditUserManages"
		Call EditUserManages
	Case "ManagesMember"
		Call ManagesMember
	Case Else
		Call Main()
End Select

Sub Main()
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
  <tr class="a1">
    <td>TEAM's提示</td>
  </tr>
  <tr class="a3">
    <td><br>
      <ul>
        <li>TEAM's! 管理组包括管理员、超级版主、版主以及关联了管理权限的特殊组，除管理员组以外，其他管理组均可详细设置管理权限。
      </ul>
      <ul>
        <li>增加自定义管理组方法为：
          <ul>
            <li>进入<a href="Admin_Group.asp?Action=IsuserGroup"><b>用户组设置</b></a>，增加一个新的特殊组；
            <li>编辑该特殊组，将该特殊组关联某种管理权限(管理员、超级版主或版主)，同时编辑该组的其他项目设置；
            <li>进入<a href="Admin_Group.asp"><b>管理组设置</b></a>，编辑该组的管理权限。
          </ul>
      </ul>
      <ul>
        <li>删除自定义管理组的方法有如下二种：
          <ul>
            <li>编辑该组基本设置，取消管理权限关联；
            <li>进入<a href="Admin_Group.asp?Action=IsuserGroup"><b>用户组设置</b></a>，编辑并取消管理权限关联或者直接删除该特殊组。
          </ul>
      </ul></td>
  </tr>
</table>
<br />
<table cellspacing="1" cellpadding="4" width="90%" align="center" class=a2>
<tr class="a1" align="center">
  <td>名称</td>
  <td>类型</td>
  <td>管理级别</td>
  <td>基本设置</td>
  <td>管理权限</td>
</tr>
<%
Dim Rs
Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,UserColor,UserImg,Rank from "&IsForum&"UserGroup Where GroupTips = 3 and MemberRank = -1 or (GroupRank = 1 or GroupRank = 2 or GroupRank = 3)")
Do While Not Rs.Eof
%>
<tr align="center">
  <td class="a3"><%=Rs(2)%></td>
  <td class="a4">内置</td>
  <td class="a3"><%=Rs(1)%></td>
  <td class="a4"><a href="Admin_Group.asp?Action=Editmodel_1&ID=<%=Rs(0)%>">[编辑]</a></td>
  <td class="a3"><a href="Admin_Group.asp?Action=Editmodel_2&ID=<%=Rs(0)%>">[编辑]</a></td>
</tr>
<%
	Rs.MoveNext
Loop
Rs.Close:Set Rs=Nothing
Response.Write "</table>"
End Sub


Sub Editmodel_1 
	Dim Rs,Group_Set_Class
	Set Rs=team.Execute("Select ID,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,UserColor,UserImg,Rank from "&IsForum&"UserGroup Where ID="&ID)
	If Rs.Eof Then 
		SuccessMsg "参数错误! "
	Else
		Group_Set_Class = Split(Rs(4),"|")
%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<form method="post" action="?Action=EditUserGroup">
 <input type="hidden" name="myid" value="<%=ID%>">
  <a name="编辑用户组"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">编辑用户组</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>用户组头衔:</b></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="GroupName" value="<%=Replace(RS(1),"&nbsp;","")%>"></td>
    </tr>
  </table>
  <br>
  <%If Request("OnGroup")="yes" Then%>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">用户所属组类别</td>
    </tr>
    <tr class="a4">
      <td><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
          <tr class="a4">
		  <input type="hidden" name="oldgroup" value="<%=Rs(3)%>">
            <%
			Dim Gs,u
			Set Gs = team.execute("Select Max(ID) as ID,Members From ["&IsForum&"UserGroup] Where GroupTips=3 or GroupTips=2 or GroupTips=1 Group By Members")
			If Gs.Eof or Gs.Bof Then
				SuccessMsg " 用户权限表数据损坏,请手动导入新表! "
			Else
				u=0
				Do While Not Gs.Eof 
					u = u+1
					Echo "<td> "
					Response.write "<input type=""radio"" name=""mygroups"" value="""&Gs(0)&""" "
					If int(Rs(3)) =  int(Gs(0)) Then Response.write " checked "
					Response.write "> "&Gs(1)&" </td>"
					If U= 5 Then 
						Echo "</tr><tr class=""a4"">"
						U=0
					End If
					Gs.MoveNext
				Loop
			End If
			Gs.Close:Set Gs=Nothing%>
        </table></td>
    </tr>
  </table>
  <%End if%>
  <br>
  <a name="基本权限"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">基本权限</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许访问论坛:</b><br>
        <span class="a4">选择“否”将彻底禁止用户访问论坛的任何页面</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(0)" value="1" <%if CID(Group_Set_Class(0))=1 Then%>checked<%end if%>>
		是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(0)" value="0" <%if CID(Group_Set_Class(0))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>阅读权限:</b><br>
        <span class="a4">设置用户浏览帖子或附件的权限级别，范围 0～255，0 为禁止用户浏览任何帖子或附件。当用户的阅读权限小于帖子的阅读权限许可(默认时为 1)时，用户将不能阅读该帖子或下载该附件</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(1)" value="<%=CID(Group_Set_Class(1))%>"></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许查看用户资料:</b><br>
        <span class="a4">设置是否允许查看其他用户的资料信息</span></td>
      <td bgcolor="#FFFFFF">
	  <input type="radio" name="Group_Set_Class(2)" value="1" <%if CID(Group_Set_Class(2))=1 Then%>checked<%end if%>>是 &nbsp; &nbsp;
      <input type="radio" name="Group_Set_Class(2)" value="0" <%if CID(Group_Set_Class(2))=0 Then%>checked<%end if%>>否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许积分转账:</b><br>
        <span class="a4">设置是否允许用户在银行中将自己的交易积分转让给其他用户。注意: 本功能需在选项中启用交易积分及个人开设银行帐号后才可使用</span></td>
      <td bgcolor="#FFFFFF">
	    <input type="radio" name="Group_Set_Class(3)" value="1" <%if CID(Group_Set_Class(3))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(3)" value="0" <%if CID(Group_Set_Class(3))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许使用搜索:</b><br>
        <span class="a4">设置是否允许通过数据库进行帖子搜索和短消息搜索。注意: 当数据量大时，全文搜索将非常耗费服务器资源，请慎用</span></td>
      <td bgcolor="#FFFFFF">
	    <input type="radio" name="Group_Set_Class(4)" value="0" <%if CID(Group_Set_Class(4))=0 Then%>checked<%end if%>>
        禁用搜索<br>
        <input type="radio" name="Group_Set_Class(4)" value="1" <%if CID(Group_Set_Class(4))=1 Then%>checked<%end if%>>
        只允许搜索标题<br>
        <input type="radio" name="Group_Set_Class(4)" value="2" <%if CID(Group_Set_Class(4))=2 Then%>checked<%end if%>>
        允许全文搜索</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许使用头像:</b><br>
        <span class="a4">设置是否允许使用自定义头像功能</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(5)" value="0" <%if CID(Group_Set_Class(5))=0 Then%>checked<%end if%>>
        禁止使用头像<br>
        <input type="radio" name="Group_Set_Class(5)" value="1" <%if CID(Group_Set_Class(5))=1 Then%>checked<%end if%>>
        允许使用论坛头像<br>
        <input type="radio" name="Group_Set_Class(5)" value="2" <%if CID(Group_Set_Class(5))=2 Then%>checked<%end if%>>
        允许使用论坛头像或上传头像<br>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许对用户评分:</b><br>
        <span class="a4">设置是否允许对其他用户的帖子进行评分操作</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(10)" value="1" <%if CID(Group_Set_Class(10))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(10)" value="0" <%if CID(Group_Set_Class(10))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>操作理由短消息通知作者:</b><br>
        <span class="a4">设置用户在对他人评分或管理操作时是否强制输入理由和通知作者</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(6)" value="0" <%if CID(Group_Set_Class(6))=0 Then%>checked<%end if%>>
        不强制<br>
        <input type="radio" name="Group_Set_Class(6)" value="1" <%if CID(Group_Set_Class(6))=1 Then%>checked<%end if%>>
        强制输入理由<br>
        <input type="radio" name="Group_Set_Class(6)" value="2" <%if CID(Group_Set_Class(6))=2 Then%>checked<%end if%>>
        强制通知作者<br>
        <input type="radio" name="Group_Set_Class(6)" value="3" <%if CID(Group_Set_Class(6))=3 Then%>checked<%end if%>>
        强制输入理由和通知作者</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许使用 文集:</b><br>
        <span class="a4">设置是否允许把文章加入个人的 文集 中，从而供他人浏览</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(7)" value="1" <%if CID(Group_Set_Class(7))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(7)" value="0" <%if CID(Group_Set_Class(7))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许发起投票:</b><br>
        <span class="a4">设置是否允许用户发表投票帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(8)" value="1" <%if CID(Group_Set_Class(8))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(8)" value="0" <%if CID(Group_Set_Class(8))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许发起活动:</b><br>
        <span class="a4">设置是否允许用户发表组织活动帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(9)" value="1" <%if CID(Group_Set_Class(9))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(9)" value="0" <%if CID(Group_Set_Class(9))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许悬赏问题:</b><br>
        <span class="a4">设置是否允许用户发表悬赏问题的帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(20)" value="1" <%if CID(Group_Set_Class(20))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(20)" value="0" <%if CID(Group_Set_Class(20))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许自定义头衔:</b><br>
        <span class="a4">设置是否允许用户设置自己的头衔名字并在帖子中显示</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(11)" value="1" <%if CID(Group_Set_Class(11))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(11)" value="0" <%if CID(Group_Set_Class(11))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>短消息收件箱容量:</b><br>
        <span class="a4">设置用户短消息最大可保存的消息数目，0 为禁止使用短消息</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(12)" value="<%=Group_Set_Class(12)%>"></td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="提 交">
  </center>
  <br>
  <a name="帖子相关"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">帖子相关</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许发新话题:</b><br>
        <span class="a4">设置是否允许发新话题。注意: 只有当用户组阅读权限高于 0 时，才能发新话题</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(13)" value="1" <%if CID(Group_Set_Class(13))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(13)" value="0" <%if CID(Group_Set_Class(13))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许发表回复:</b><br>
        <span class="a4">设置是否允许发表回复。注意: 只有当用户组阅读权限高于 0 时，才能发表回复</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(14)" value="1" <%if CID(Group_Set_Class(14))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(14)" value="0" <%if CID(Group_Set_Class(14))=0 Then%>checked<%end if%>>
        否</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许参与投票:</b><br>
        <span class="a4">设置是否允许参与论坛的投票</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(15)" value="1" <%if CID(Group_Set_Class(15))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(15)" value="0" <%if CID(Group_Set_Class(15))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
    <td width="60%" bgcolor="#F8F8F8" ><b>允许直接发帖:</b><br>
        <span class="a4">本选项只在论坛被设置为需要发帖审核时才起作用，用以选择“是”将允许该组用户不经审核而直接发布新帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(16)" value="0" <%if CID(Group_Set_Class(16))=0 Then%>checked<%end if%>>
        全部需要审核<br>
        <input type="radio" name="Group_Set_Class(16)" value="1" <%if CID(Group_Set_Class(16))=1 Then%>checked<%end if%>>
        发新回复不需要审核<br>
        <input type="radio" name="Group_Set_Class(16)" value="2" <%if CID(Group_Set_Class(16))=2 Then%>checked<%end if%>>
        发新主题不需要审核<br>
        <input type="radio" name="Group_Set_Class(16)" value="3" <%if CID(Group_Set_Class(16))=3 Then%>checked<%end if%>>
        全部不需要审核</td>
    </tr><tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许发匿名贴:</b><br>
        <span class="a4">是否允许用户匿名发表主题和回复，只要用户组或本论坛允许，用户均可使用匿名发帖功能。匿名发帖不同于游客发帖，用户需要登录后才可使用，版主和管理员可以查看真实作者</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(17)" value="1" <%if CID(Group_Set_Class(17))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(17)" value="0" <%if CID(Group_Set_Class(17))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许设置帖子权限:</b><br>
        <span class="a4">设置是否允许设置帖子需要指定阅读权限才可浏览</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(18)" value="1" <%if CID(Group_Set_Class(18))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(18)" value="0" <%if CID(Group_Set_Class(18))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许使用主题标色:</b><br>
        <span class="a4">设置是否允许发表主题时，可以选择帖子的标题颜色。</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(19)" value="1" <%if CID(Group_Set_Class(19))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(19)" value="0" <%if CID(Group_Set_Class(19))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许签名中使用 UBB 代码:</b><br>
        <span class="a4">设置是否解析用户签名中的 UBB 代码</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(21)" value="1" <%if CID(Group_Set_Class(21))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(21)" value="0" <%if CID(Group_Set_Class(21))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许签名中使用 [img] 代码:</b><br>
        <span class="a4">设置是否解析用户签名中的 [img] 代码</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(22)" value="1" <%if CID(Group_Set_Class(22))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(22)" value="0" <%if CID(Group_Set_Class(22))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>最大签名长度:</b> <br>
        <span class="a4">设置用户签名最大字节数，0 为不允许用户使用签名</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(23)" value="<%=Group_Set_Class(23)%>">
      </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="提 交">
  </center>
  <br>
  <a name="附件相关"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">附件相关</td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许下载/查看附件:</b><br>
        <span class="a4">设置是否允许在没有设置特殊权限的论坛中下载或查看附件</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(24)" value="1" <%if CID(Group_Set_Class(24))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(24)" value="0" <%if CID(Group_Set_Class(24))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许发布附件:</b><br>
        <span class="a4">设置是否允许上传附件到论坛中。需要基本选项允许上传附件才有效，请参考 <B>基本选项</B> - <B>附件设置</B> </span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Group_Set_Class(25)" value="1" <%if CID(Group_Set_Class(25))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Group_Set_Class(25)" value="0" <%if CID(Group_Set_Class(25))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>每次上传附件个数:</b><br>
        <span class="a4">设置用户每次上传附件时，可同时上传的附件个数。</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(26)" value="<%=Group_Set_Class(26)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>最大附件尺寸(KB):</b><br>
        <span class="a4">设置附件最大字节数，需要 基本选项 允许才有效，此设置数只能小于或等于 基本选项 的设置数，超过则按照 基本选项 的设置数执行。<BR> 目前的基本设置参数为：<%=team.Forum_setting(71)%></span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(27)" value="<%=Group_Set_Class(27)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>每天上传附件的最大个数:</b><br>
        <span class="a4">设置用户每 24 小时可以上传的附件总个数，0 为不限制。</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(28)" value="<%=Group_Set_Class(28)%>">
      </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许附件类型:</b><br>
        <span class="a4">设置允许上传的附件扩展名，多个扩展名之间用半角逗号 "," 分割，留空为不限制</span></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30" name="Group_Set_Class(29)" value="<%=Group_Set_Class(29)%>">
      </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="detailsubmit" value="提 交">
  </center>
</form>
<%
	End If
	Rs.Close:Set Rs=Nothing
End Sub
	
Sub EditUserGroup
	Dim myid,Saves,u
	myid = Request.form("myid")
	If myid = "" or (Not IsNumeric(myid)) Then  
		SuccessMsg "参数错误! "
	Else
		for u=0 to 29
			if Saves ="" Then
				Saves = Request.form("Group_Set_Class(0)")
			else
				Saves = Saves & "|"& Request.form("Group_Set_Class("&u&")")
			End if
		next
		team.execute("Update ["&IsForum&"UserGroup] set IsBrowse = '"&Saves&"',GroupName='"&request.form("GroupName")&"' Where ID="& MyID)
		If Cid(Request.Form("oldgroup")) <> Cid(Request.Form("mygroups")) Then
			team.execute("Update ["&IsForum&"UserGroup] set GroupRank="&Cid(Request.Form("mygroups"))&"  Where ID="& MyID)
		End If
		Application.Contents.RemoveAll()
		SuccessMsg " 用户浏览权限更新成功!"
	End If
End Sub

Sub Editmodel_2
	Dim Rs,Manage_Set_Class
	Set Rs=team.Execute("Select Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,UserColor,UserImg,Rank from "&IsForum&"UserGroup Where ID="&ID)
	If Rs.Eof Then 
		SuccessMsg "参数错误! "
	Else
		Manage_Set_Class = Split(Rs(5),"|")
	%><br><br>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<form method="post" action="?Action=EditUserManages">
  <input type="hidden" name="myid" value="<%=ID%>">
  <a name="编辑管理成员组 - <%=RS(1)%>"></a>
  <table cellspacing="1" cellpadding="4" width="90%" align="center" class="a2">
    <tr class="a1">
      <td colspan="2">编辑管理成员组 - <%=RS(1)%></td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许编辑帖子:</b><br>
        <span class="a4">设置是否允许编辑管理范围内的帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(0)" value="1" <%if CID(Manage_Set_Class(0))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(0)" value="0" <%if CID(Manage_Set_Class(0))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许置顶帖子:</b><br>
        <span class="a4">设置允许置顶主题的级别。总置顶将在全论坛置顶，本版置顶 仅在版块置顶。</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(1)" value="0" <%if CID(Manage_Set_Class(1))=0 Then%>checked<%end if%>>
        不允许置顶<br>
        <input type="radio" name="Manage_Set_Class(1)" value="1" <%if CID(Manage_Set_Class(1))=1 Then%>checked<%end if%>>
        允许本版置顶<br>
        <input type="radio" name="Manage_Set_Class(1)" value="2" <%if CID(Manage_Set_Class(1))=2 Then%>checked<%end if%>>
        允许本版置顶/总置顶<br>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许审核帖子:</b><br>
        <span class="a4">设置是否允许审核用户发表的帖子，只在论坛设置需要审核时有效</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(2)" value="1" <%if CID(Manage_Set_Class(2))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(2)" value="0" <%if CID(Manage_Set_Class(2))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许删除帖子:</b><br>
        <span class="a4">设置是否允许删除管理范围内的帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(3)" value="1" <%if CID(Manage_Set_Class(3))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(3)" value="0" <%if CID(Manage_Set_Class(3))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许移动主题:</b><br>
        <span class="a4">设置是否允许移动管理范围内的帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(4)" value="1" <%if CID(Manage_Set_Class(4))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(4)" value="0" <%if CID(Manage_Set_Class(4))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许拉前主题:</b><br>
        <span class="a4">设置是否允许拉前管理范围内的帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(5)" value="1" <%if CID(Manage_Set_Class(5))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(5)" value="0" <%if CID(Manage_Set_Class(5))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许锁定/解除锁定帖子:</b><br>
        <span class="a4">设置是否允许锁定/解除锁定管理范围内的帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(6)" value="1" <%if CID(Manage_Set_Class(6))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(6)" value="0" <%if CID(Manage_Set_Class(6))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许关闭/打开主题:</b><br>
        <span class="a4">设置是否允许锁定/解除锁定管理范围内的帖子</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(7)" value="1" <%if CID(Manage_Set_Class(7))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(7)" value="0" <%if CID(Manage_Set_Class(7))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许将主题加入/移出精华:</b><br>
        <span class="a4">设置是否允许将管理范围内的帖子加入/移出精华区</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(8)" value="1" <%if CID(Manage_Set_Class(8))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(8)" value="0" <%if CID(Manage_Set_Class(8))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许将主题批量加入/移出专题:</b><br>
        <span class="a4">设置是否允许将管理范围内的帖子加入/移出专题</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(12)" value="1" <%if CID(Manage_Set_Class(12))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(12)" value="0" <%if CID(Manage_Set_Class(12))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>

    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许查看 IP:</b><br>
        <span class="a4">设置是否允许查看用户 IP</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(10)" value="1" <%if CID(Manage_Set_Class(10))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(10)" value="0" <%if CID(Manage_Set_Class(10))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许禁止 IP:</b><br>
        <span class="a4">设置是否允许添加或修改禁止 IP 设置</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(11)" value="1" <%if CID(Manage_Set_Class(11))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(11)" value="0" <%if CID(Manage_Set_Class(11))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许修改版块介绍:</b><br>
        <span class="a4">只有管理级别的用户才可以有此权限</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(9)" value="1" <%if CID(Manage_Set_Class(9))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(9)" value="0" <%if CID(Manage_Set_Class(9))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许禁止用户:</b><br>
        <span class="a4">设置是否允许禁止用户发帖或访问</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(13)" value="1" <%if CID(Manage_Set_Class(13))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(13)" value="0" <%if CID(Manage_Set_Class(13))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
    <tr>
      <td width="60%" bgcolor="#F8F8F8" ><b>允许审核用户:</b><br>
        <span class="a4">设置是否允许审核新注册用户，只在论坛设置需要人工审核新用户时有效</span></td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="Manage_Set_Class(14)" value="1" <%if CID(Manage_Set_Class(14))=1 Then%>checked<%end if%>>
        是 &nbsp; &nbsp;
        <input type="radio" name="Manage_Set_Class(14)" value="0" <%if CID(Manage_Set_Class(14))=0 Then%>checked<%end if%>>
        否 </td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="groupsubmit" value="提 交">
  <center>
</form>
<%
	End If
End Sub
	
Sub EditUserManages
	Dim myid,Saves,u
	myid = Request.form("myid")
	If myid = "" or (Not IsNumeric(myid)) Then  
		SuccessMsg "参数错误! "
	Else
		for u=0 to 14
			if Saves ="" Then
				Saves = Request.form("Manage_Set_Class(0)")
			else
				Saves = Saves & "|"& Request.form("Manage_Set_Class("&u&")")
			End if
		next
		team.execute("Update UserGroup set IsManage = '"&Saves&"' Where ID="& MyID)
		Application.Contents.RemoveAll()
		SuccessMsg " 管理组权限更新成功!"
	End If
End Sub
	
Sub  IsuserGroup	
	%>
<body Style="background-color:#8C8C8C" text="#000000" leftmargin="10" topmargin="10">
<br>
<table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
  <tr class="a1">
    <td>技巧提示</td>
  </tr>
  <tr bgcolor="#F8F8F8">
    <td><br>
      <ul>
        <li>TEAM论坛用户组分为系统组、特殊组和会员组，会员组以积分确定组别和权限，而系统组和特殊组是人为设定，不会由论坛系统自行改变。
      </ul>
      <ul>
        <li>系统组和特殊组的设定不需要指定积分，TEAM 预留了从论坛管理员到游客等的 8 个系统头衔，特殊组的用户需要在编辑会员时将其加入。
      </ul>
      <ul>
        <li>如果用户被指定了用户组，那么删除该用户组别将导致用户无法访问论坛，需要手动进行设置该用户所在的用户组。
      </ul>
      <ul>
        <li>如果修改了用户组的名称，那么将导致所有的论坛用户重新赋值。此操作将消耗大量的系统资源。
      </ul>	  
	  </td>
  </tr>
</table>
<br>
<form method="post" action="?Action=ManagesMember&master=0">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="7">会员用户组</td>
    </tr>
	<a name="会员用户组"></a>
    <tr class="a4" align="center">
      <td width="48"><input type="checkbox" name="chkall" class="radio" onClick="checkall(this.form)">
        删?</td>
      <td>组头衔</td>
      <td>积分大于</td>
      <td>星星数</td>
      <td>名称颜色</td>
      <td>组头像</td>
      <td>编辑</td>
    </tr>
	<%
	Dim Rs
	Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,Rank,UserColor,Userimg from "&IsForum&"UserGroup Where GroupTips=1 Order By Memberrank Desc")
	Do While Not Rs.Eof
	%><input type="hidden" name="upid" value="<%=RS(0)%>">
    <tr align="center">
      <td bgcolor="#F8F8F8">
		<input type="checkbox" name="myid" value="<%=RS(0)%>"  class="radio">
	  </td>
      <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName" value="<%=Replace(RS(2)," ","")%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="6" name="Memberrank" value="<%=RS(3)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank" value="<%=RS(7)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor" value="<%=RS(8)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg" value="<%=RS(9)%>"></td>
      <td bgcolor="#FFFFFF" nowrap><a href="?Action=Editmodel_1&ID=<%=RS(0)%>">[详情]</a></td>
    </tr>
	<%	Rs.Movenext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
    <tr>
      <td colspan="7" class="a4" height="2"></td>
    </tr>
    <tr align="center" bgcolor="#F8F8F8">
      <td>新增:</td>
      <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName1" value=""></td>
      <td bgcolor="#F8F8F8"><input type="text" size="6" name="Memberrank1" value=""></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank1" value=""></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor1" value=""></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg1" value=""></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <br>
  <center>
  <input type="submit" name="groupsubmit" value="提 交">
  &nbsp;
</form>
<br>
<br>
<form method="post" action="?Action=ManagesMember&master=1">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="6">特殊用户组</td>
    </tr>
	<a name="特殊用户组"></a>
    <tr class="a4" align="center">
      <td width="48"><input type="checkbox" name="chkall" class="a4" onClick="checkall(this.form)">
        删?</td>
      <td nowrap>组头衔</td>
      <td nowrap>星星数</td>
      <td nowrap>名称颜色</td>
      <td nowrap>组头像</td>
      <td nowrap>编辑</td>
    </tr>
	<%
	Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,Rank,UserColor,Userimg from "&IsForum&"UserGroup Where GroupTips=2 Order By ID Desc")
	Do While Not Rs.Eof
	%><input type="hidden" name="gupid" value="<%=RS(0)%>">
    <tr align="center">
      <td bgcolor="#F8F8F8">
		<input type="checkbox" name="myid" value="<%=RS(0)%>"  class="radio">
	  </td>
	  <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName" value="<%=Replace(RS(2)," ","")%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank" value="<%=RS(7)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor" value="<%=RS(8)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg" value="<%=RS(9)%>"></td>
      <td bgcolor="#FFFFFF" nowrap><a href="?Action=Editmodel_1&ID=<%=RS(0)%>&OnGroup=yes">[详情]</a></td>
    </tr>
	<%  Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
	<tr>
      <td colspan="6" class="a4" height="2"></td>
    </tr>
    <tr align="center" bgcolor="#F8F8F8">
      <td>新增:</td>
      <td><input type="text" size="20" name="GroupName1"></td>
      <td><input type="text" size="2" name="Rank1"></td>
      <td><input type="text" size="30" name="UserColor1"></td>
      <td><input type="text" size="20" name="Userimg1"></td>
	  <td>&nbsp;</td>
    </tr>
  </table>
  <br>
  <center>
    <input type="submit" name="groupsubmit" value="提 交">
  </center>
</form>
<br>
<form method="post" action="?Action=ManagesMember&master=2">
  <table cellspacing="1" cellpadding="4" width="95%" align="center" class="a2">
    <tr class="a1">
      <td colspan="48">系统用户组</td>
    </tr>
	<a name="系统用户组"></a>
    <tr class="a4" align="center">
	  <td width="2"></td>
	  <td>系统头衔</td>
      <td>组头衔</td>
      <td>星星数</td>
      <td>名称颜色</td>
      <td>组头像</td>
      <td>编辑</td>
    </tr>
	<%
	Set Rs=team.Execute("Select ID,Members,GroupName,Memberrank,GroupRank,IsBrowse,IsManage,Rank,UserColor,Userimg from "&IsForum&"UserGroup Where MemberRank = -1 Order By GroupRank Desc")
	Do While Not Rs.Eof
	%>
    <tr align="center">
	  <td bgcolor="#FFFFFF"><input type="hidden" name="myid" value="<%=RS(0)%>"></td>
	  <td bgcolor="#F8F8F8"><%=RS(1)%></td>
      <td bgcolor="#FFFFFF"><input type="text" size="20" name="GroupName" value="<%=Replace(RS(2),"&nbsp;","")%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="2"name="Rank" value="<%=RS(7)%>"></td>
      <td bgcolor="#FFFFFF"><input type="text" size="30"name="UserColor" value="<%=RS(8)%>"></td>
      <td bgcolor="#F8F8F8"><input type="text" size="20" name="Userimg" value="<%=RS(9)%>"></td>
      <td bgcolor="#FFFFFF" nowrap><a href="?Action=Editmodel_1&ID=<%=RS(0)%>">[详情]</a></td>
    </tr>
	<%  Rs.MoveNext
	Loop
	Rs.Close:Set Rs=Nothing
	%>
  </table>
  <br>
  <center>
    <input type="submit" name="groupsubmit" value="提 交">
  </center>
</form>
<br>
<br>
<%
End Sub
Sub ManagesMember
	Dim Myid,GroupName,Rank,Memberrank,UserColor,Userimg,tmp,u
	Dim ho,Rs
	tmp = ""
	GroupName = ""
	If request.form("myid") = "" Then
		Myid=Split(Request.Form("upid"),",")
		GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
		Rank = Split(Replace(Request.Form("Rank")," ",""),",")
		UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
		Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
		Memberrank = Split(Replace(Request.Form("Memberrank")," ",""),",")
		For U=0 To Ubound(Myid)
			Set Rs = team.Execute("Select GroupName,Members From ["&IsForum&"UserGroup] Where ID="&Myid(U))
			If Not (Rs.Eof And Rs.Bof) Then
				If Trim(Rs(0)) <> Trim(GroupName(U)) Then
					team.Execute("Update ["&IsForum&"User] set Members='"&Rs(1)&"',Levelname='"&GroupName(U)&"||"&UserColor(U)&"||"&Userimg(U)&"||"&Cid(Rank(U))&"||0' where UserGroupID="&Myid(U))
				End If
			End If
		Next
	End If

	Select Case Request("master")
		Case "0"
			for each ho in request.form("myid")
				Team.execute("Delete from "&isforum&"UserGroup Where ID="&ho)
			next
			If request.form("myid") = "" Then
				If Request.Form("upid") <> "" Then
					If Instr(Request.Form("upid"),",")>0 Then
						
						Myid=Split(Request.Form("upid"),",")
						GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
						Rank = Split(Replace(Request.Form("Rank")," ",""),",")
						UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
						Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
						Memberrank = Split(Replace(Request.Form("Memberrank")," ",""),",")
						For U=0 To Ubound(Myid)
							team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&GroupName(U)&"',Memberrank="&Cid(Memberrank(U))&",Rank="&Cid(Rank(U))&",UserColor='"&UserColor(U)&"',Userimg='"&Userimg(U)&"' where ID="&Myid(U))
						Next
					Else
						team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&HtmlEncode(Request.Form("GroupName"))&"',Rank="&Cid(Request.Form("Rank"))&",UserColor='"&Request.Form("UserColor")&"',Userimg='"&Request.Form("Userimg")&"' where ID="&Request.Form("upid"))
					End If
				End if
				if Not (Request.Form("GroupName1")="" and Request.Form("Rank1")="") Then
					team.Execute("insert into "&IsForum&"UserGroup (GroupName,GroupTips,Rank,UserColor,Userimg,GroupRank,MemberRank,Members,IsBrowse,IsManage) values ('"&HtmlEncode(Request.Form("GroupName1"))&"',1,"&Cid(Request.Form("Rank1"))&",'"&Request.Form("UserColor1")&"','"&Request.Form("Userimg1")&"',0,"&CID(Request.Form("Memberrank1"))&",'注册用户','1|1|1|0|0|0|1|0|0|0|0|0|20|1|1|1|0|0|0|0|0|0|0|0|0|0|1|100|50|','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
				End If
			End If
			tmp = tmp & " 会员用户组 修改设置成功! "			
		Case "1"
			for each ho in request.form("myid")
				Team.execute("Delete from "&isforum&"UserGroup Where ID="&ho)
			next
			If request.form("myid") = "" Then
				If Request.Form("gupid") <> "" Then
					If Instr(Request.Form("gupid"),",")>0 Then
						Myid=Split(Request.Form("gupid"),",")
						GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
						Rank = Split(Replace(Request.Form("Rank")," ",""),",")
						UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
						Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
						For U=0 To Ubound(Myid)
							team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&GroupName(U)&"',Rank="&Cid(Rank(U))&",UserColor='"&UserColor(U)&"',Userimg='"&Userimg(U)&"' where ID="&Myid(U))
						Next
					Else
						team.Execute("Update "&IsForum&"UserGroup set 		GroupName='"&HtmlEncode(Request.Form("GroupName"))&"',Rank="&Cid(Request.Form("Rank"))&",UserColor='"&Request.Form("UserColor")&"',Userimg='"&Request.Form("Userimg")&"' where ID="&Request.Form("gupid"))
					End If
				End if
				if Not (Request.Form("GroupName1")="" and Request.Form("Rank1")="") Then
					team.Execute("insert into "&IsForum&"UserGroup (GroupName,GroupTips,Rank,UserColor,Userimg,GroupRank,MemberRank,Members,IsBrowse,IsManage) values ('"&HtmlEncode(Request.Form("GroupName1"))&"',2,"&Cid(Request.Form("Rank1"))&",'"&Request.Form("UserColor1")&"','"&Request.Form("Userimg1")&"',4,-1,'"&HtmlEncode(Request.Form("GroupName1"))&"','1|255|1|1|1|2|0|1|1|1|1|1|50|1|1|1|0|1|0|0|1|1|1|50|1|1|10|1024|200|','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
				End If
			End If
			tmp = tmp & " 特殊用户组 修改设置成功! "
		Case "2"
			Myid=Split(Request.Form("myid"),",")
			GroupName=Split(Replace(Request.Form("GroupName")," ",""),",")
			Rank = Split(Replace(Request.Form("Rank")," ",""),",")
			UserColor = Split(Replace(Request.Form("UserColor")," ",""),",")
			Userimg = Split(Replace(Request.Form("Userimg")," ",""),",")
			For U=0 To Ubound(Myid)
				team.Execute("Update "&IsForum&"UserGroup set GroupName='"&GroupName(U)&"',Rank="&Rank(U)&",UserColor='"&UserColor(U)&"',Userimg='"&Userimg(U)&"' where ID="&Myid(U))
			Next
			tmp = tmp & " 系统用户组 修改设置成功! "
		Case Else
			tmp = tmp & "参数错误! "
	End Select
	Application.Contents.RemoveAll()
	SuccessMsg tmp &" <br />感谢使用TEAM论坛系统,稍后系统将自动返回设置页面! <meta http-equiv=refresh content=3;url=Admin_Group.asp?Action=IsuserGroup#系统用户组>"
End Sub
%>
