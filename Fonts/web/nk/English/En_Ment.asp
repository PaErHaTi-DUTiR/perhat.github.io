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
Response.Redirect"En_Loginsb.asp?msg=You have not chosen the goods , can not check out！"
End If
If Session("UserName")<>"" Then
Response.Redirect"En_Ment1.asp"
Else
%>
<HTML><HEAD>
<TITLE>Member login</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("Please input the user name！");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("Please input the password！");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
</HEAD>
<BODY bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br>
<TABLE width="600" height="95%"border=0 align="center" cellPadding=1 cellSpacing=1 bgcolor="#666666">
	<TR align=middle> 
		<TD height="22" bgcolor="#999999" 2><div align="left"
><font color="#333333"><font color="#FFFFFF"
>I am sorry, if will buy the goods , must carry on login first , if you have not registered for our member yet, have been invited <a href="En_UserReg.asp" target="_blank"
><font color="#FF0000">Register</font></a>And then buy the goods again！&nbsp;</font></font><br
><font color="#FFFFFF">If you are a member, please login first：</font></div></TD>
	</TR>
	<TR vAlign=top align=middle> 
		<TD height="160" valign="middle" bgcolor="#CCCCCC">
		<form action='En_MentLogin.asp'method='post'name='UserLogin'onSubmit='return CheckForm();'>
			<TABLE cellSpacing=0 cellPadding=0 width=100% height="47">
				<TR> 
					<TD align=middle height="47" width="100%">
						<TABLE cellSpacing=1 width=100% height="1">
							<TR> 
								<TD  width=328 height=18><div align="right">UserName：</div></TD>
								<TD  width=419 height=18><input name='UserName'type='text'id='UserName'size='13'maxlength='16'> </TD>
							</TR>
							<tr> 
								<TD width=328 height=1><div align="right">UserPass：</div></TD>
								<TD width=419 height=1><input name='Password'type='password'id='Password'size='13'maxlength='16'> </TD>
							</tr>
							<TR> 
								<TD colSpan=2 height="1"><DIV align=center
><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"
><input name='Login'type='submit'id='Login'value='Login '>&nbsp; </DIV></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	  </TD>
	</TR>
	<TR vAlign=top align=middle> 
		<TD height="22" bgcolor="#999999" 2> <p>&nbsp;</p></TD>
	</TR>
</TABLE>
</BODY></HTML>
<%End If%>
