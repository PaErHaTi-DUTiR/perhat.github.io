<!--#include file="bsconfig.asp"-->
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<%
id=request("id")
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin where Id=" & id
rs.open sqltext,conn,1,1
%>

<!-- #include file="Inc/Head.asp" -->
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--


function form1_onsubmit()
{
if (document.FORM1.pwd1.value!=document.FORM1.pwd2.value)
{
alert ("请确认您的密码。");
document.FORM1.pwd1.value='';
document.FORM1.pwd2.value='';
document.FORM1.pwd1.focus();
return false;
}
}


//-->
</SCRIPT>
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>管理员密码修改</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<FORM language=javascript name=FORM1  onsubmit="return form1_onsubmit()" action=Bs_AdminPass_Save.asp?uid=<%=rs("Id")%> method=post>
			<INPUT type=hidden value=<%=rs("UserName")%> name=UserName>
				<table width="50%" border="0" align="center" cellpadding="0" cellspacing="2">
					<tr> 
						<td height="20" colspan="2">&nbsp;</td>
					</tr>
					<tr> 
						<td width="29%" height="22"> <div align="right">管理员帐号：</div></td>
						<td width="71%"><%=rs("UserName")%></td>
					</tr>
					<tr> 
						<td height="22"> <div align="right">新密码：</div></td>
						<td><input name="pwd1" type="text" size="16" maxlength="20"></td>
					</tr>
					<tr> 
						<td height="22"> <div align="right">密码确认：</div></td>
						<td><input name="pwd2" type="text" size="16" maxlength="20"></td>
					</tr>
					<tr> 
						<td height="22" colspan="2"><div align="center"><INPUT  type=submit value='确认修改'name=Submit2></div></td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
