<!--#include file="Include/En_System.asp"-->
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<%If Session("UserName")="" Then
response.redirect"En_Server.asp"
end if

name=request.querystring("name")
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * from Bs_Book where id="&id, conn,3,3

if rs("name")<>name then
response.redirect "En_NetBook.asp"
end if
%>
<html>
<head>
<title>∷<%=BsEnCompanyName%>∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Author" content="haoxiao">
<meta name="Contact" content="haoxiao@21cn.com">
<meta name="Copyright" content="BOSSI">
<meta name="Keywords" content="Bossi Net Maintain,BOSSI company(enterprises) Website administrative system">
<meta name="Description" content="Bossi Net Maintain,BOSSI company(enterprises) Website administrative system">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<body bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br>
<TABLE width="600" height="380" align="center" cellPadding=4 cellSpacing=1 bgColor=#666666>

	<TR vAlign=top bgColor=#eeeeee> 
		<TD  width="100%" height="1" bgcolor="#999999"> 
<%if rs("name")="管理员" then%>
<font color="#FFFFFF">『announcement title』：</font> 
<%else%>
<font color="#FFFFFF">『Your message title』：</font> 
<%end if%><font color="#006699"><%=rs("title")%></font>
		</TD>
	</TR>
	<TR vAlign=top bgColor=#eeeeee> 
		<TD  width="100%" height="366"> 
<%if rs("name")="管理员" then%>
Website announcement： 
<%else%>
Your message content: 
<%end if%>
			<div align="center"> 
			<center>
			<table border="0" width="98%" cellpadding="2">
				<tr> 
					<td width="100%"> <%=rs("content")%> </td>
				</tr>
			</table>
			</center>
<p align="right">（Time：<%=rs("time")%>） </div>
			<div align="center"> 
			<center>
			<table border="0" width="98%" cellpadding="2">
				<tr> 
					<td width="100%">&nbsp; 
<%If rs("rebook")<>"" then%>
<hr size="1"> <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b>Reply given to you</b></p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">　</p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FF0000"><b><%=Session("UserName")%> howdy:</b></font></p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FF0000">&nbsp;&nbsp;&nbsp; 
<%=rs("rebook")%> 
<%else if rs("name")<>"管理员" Then%>
Our staff members have not replied your message yet, has asked to wait patiently ！ 
<%end if%>
<%End If%>
&nbsp;</font></p></td>
				</tr>
			</table>
			</center>
			</div>
		</TD>
	</TR>
	<TR bgColor=#eeeeee> 
		<TD  width="530" height="26" bgcolor="#999999"><div align="center"
><input type="button" value="Close" name="B5" onClick="window.close();" style="font-size: 9pt"></div></TD>
	</TR>

</TABLE>
</body>
</html>
<%
rs.close
set rs=nothing
call CloseConn()
%>
