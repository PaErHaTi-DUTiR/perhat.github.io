<!--#include file="../Inc/Util.asp" -->
<%
'������������������������������������������������������������
'�������������� COPYRIGHT ������������ ��
'���������ƣ���ҵ��վ����ϵͳMac3.0  (ASBDcorpweb Mac3.0)  �� 
'����Ȩ���У�WWW.ASBD.CN  BBS.ASBD.CN                      ��
'������������amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  �� 
'��Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ��
'������������������������������������������������������������
%>
<%
Head="Articles that you put the shopping cart have been returned ��" '�����빺�ﳵ����Ʒ��ȫ���˻�

Session("ProductList") = ""
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<title>Clear empty shopping cart</title>
</head>
<body topmargin="5" bgcolor="#eeeeee">
<div align="center"><center>
<table width="100%" border="0" class="table1" bordercolor="#62ACFF" cellspacing="0" >
<tr>
<td width="80%" valign="top">��<p align="center" ><%=Head%></p>
<p align="center">��<br><input type="button" value="Close" name="B2" onClick="window.close();" style="font-size: 9pt"></td>
</tr>
</table>
</center></div>
</body>
</html>
