<!--#include file="../Inc/Util.asp" -->
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
Head="Articles that you put the shopping cart have been returned ！" '您放入购物车的物品已全数退回

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
<td width="80%" valign="top">　<p align="center" ><%=Head%></p>
<p align="center">　<br><input type="button" value="Close" name="B2" onClick="window.close();" style="font-size: 9pt"></td>
</tr>
</table>
</center></div>
</body>
</html>
