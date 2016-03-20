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
If Session("ProductList")="" or Session("ProductList")="'undefined'" Then
Response.Redirect"Bs_Loginsb.asp?msg=您没有选购商品，不能够结帐！"
End If
If Session("UserName")<>"" Then
Response.Redirect"Bs_Ment1.asp"
Else
%>
<HTML><HEAD>
<TITLE>会员登录</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("请输入用户名！");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("请输入密码！");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></HEAD>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br>
<TABLE width="600" height="95%"border=1 align="center" cellPadding=0 cellSpacing=0 bordercolor="#999999" bgcolor="#666666">
	<TR align=middle>
	  <TD height="22" background="images/b1.gif" bgcolor="#333333" 2>　※温馨小提示：</TD>
  </TR>
	<TR align=middle> 
		<TD height="22" bgcolor="#999999" 2><div align="left"
><font color="#333333"><font color="#FFFFFF"
>　　对不起，如要购买商品，必须先进行登录，如果您还没有注册为我们的会员，请<a href="Bs_UserReg.asp" target="_blank"
><font color="#FF0000">注册</font></a>后再购买商品！&nbsp;</font></font><br
>
		    <font color="#FFFFFF">如果您已经是会员，请先登录：</font></div></TD>
	</TR>
	<TR vAlign=top align=middle> 
		<TD height="160" valign="middle" bgcolor="#CCCCCC">
		<form action='Bs_Mentlogin.asp'method='post'name='UserLogin'onSubmit='return CheckForm();'>
			<TABLE cellSpacing=0 cellPadding=0 width=100% height="47">
				<TR> 
					<TD align=middle height="47" width="100%">
						<TABLE cellSpacing=1 width=100% height="1">
							<TR> 
								<TD  width=328 height=18><div align="right">会员登录名称：</div></TD>
								<TD  width=419 height=18><input name='UserName'type='text'id='UserName'size='13'maxlength='16'> </TD>
							</TR>
							<tr> 
								<TD width=328 height=1><div align="right">会员登录密码：</div></TD>
								<TD width=419 height=1><input name='Password'type='password'id='Password'size='13'maxlength='16'> </TD>
							</tr>
							<TR> 
								<TD colSpan=2 height="1"><DIV align=center
><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"
><input name='Login'type='submit'id='Login'value='登录 '>&nbsp; </DIV></TD>
							</TR>
						</TABLE>					</TD>
				</TR>
			</TABLE>
		</FORM>	  </TD>
	</TR>
</TABLE>
</BODY></HTML>
<%End If%>
