<!--#include file="Include/Bs_System.asp"-->
<%
'������������������������������������������������������������
'�������������� COPYRIGHT ������������ ��
'���������ƣ���ҵ��վ����ϵͳMac3.0  (ASBDcorpweb Mac3.0)  �� 
'����Ȩ���У�WWW.ASBD.CN  BBS.ASBD.CN                      ��
'������������amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  �� 
'��Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ��
'������������������������������������������������������������
%>
<%If Session("UserName")="" Then
response.redirect"Bs_Server.asp"
end if

name=request.querystring("name")
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * from Bs_Book where id="&id, conn,3,3

if rs("name")<>name then
response.redirect "Bs_NetBook.asp"
end if
%>
<html>
<head>
<title>��<%=BsCompanyName%>��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Author" content="haoxiao">
<meta name="Contact" content="haoxiao@21cn.com">
<meta name="Copyright" content="BOSSI">
<meta name="Keywords" content="���R����,BOSSI ��˾(��ҵ)��վ����ϵͳ">
<meta name="Description" content="���R����,BOSSI ��˾(��ҵ)��վ����ϵͳ">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<body bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br>
<TABLE width=90% height="400" align="center" cellPadding=4 cellSpacing=1 bgColor=#666666>

	<TR vAlign=top bgColor=#eeeeee> 
		<TD  width="100%" height="1" bgcolor="#999999"> 
<%if rs("name")="����Ա" then%>
<font color="#FFFFFF">��������⡻��</font> 
<%else%>
<font color="#FFFFFF">���������Ա��⡻��</font> 
<%end if%><font color="#006699"><%=rs("title")%></font>
		</TD>
	</TR>
	<TR vAlign=top bgColor=#eeeeee> 
		<TD  width="100%" height="366"> 
<%if rs("name")="����Ա" then%>
��վ���棺 
<%else%>
������������: 
<%end if%>
			<div align="center"> 
			<center>
			<table border="0" width="98%" cellpadding="2">
				<tr> 
					<td width="100%"> <%=rs("content")%> </td>
				</tr>
			</table>
			</center>
<p align="right">��ʱ�䣺<%=rs("time")%>�� </div>
			<div align="center"> 
			<center>
			<table border="0" width="98%" cellpadding="2">
				<tr> 
					<td width="100%">&nbsp; 
<%If rs("rebook")<>"" then%>
<hr size="1"> <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b>�����Ļظ�</b></p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">��</p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FF0000"><b><%=Session("UserName")%> ����:</b></font></p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FF0000">&nbsp;&nbsp;&nbsp; 
<%=rs("rebook")%> 
<%else if rs("name")<>"����Ա" Then%>
���ǵĹ�����Ա��δ�ظ��������ԣ������ĵȴ��� 
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
><input type="button" value="�ر�" name="B5" onClick="window.close();" style="font-size: 9pt"></div></TD>
	</TR>

</TABLE>
</body>
</html>
<%
rs.close
set rs=nothing
call CloseConn()
%>
