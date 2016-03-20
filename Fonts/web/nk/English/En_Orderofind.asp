<!--#include file="../Inc/Conn.asp" -->
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
Form_ID = Request.form("Form_Id")
IF Session("UserName")="" Then
'response.redirect "En_orderlogin.asp"
Else
set Rs3 = Server.CreateObject("ADODB.recordset")
sql3="select * from Bs_OrderList where Form_Id="&Form_Id&""
rs3.open sql3,conn,1,1
IF  rs3.RecordCount >=1  then
IF Session("UserName")=rs3("UserName") Then
%>
<html>
<head>
<title>Detailed information of customer indent </title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
</head>
<%
id=Form_Id
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_OrderList where Form_Id=" & id
rs.open sqltext,conn,1,1
%>

<body>
<br>
<div align="center"><center>
<TABLE cellSpacing=1 cellPadding=4 width=562 bgColor=#006699 height="159">
	<TR> 
		<TD width="548" height="10"  colSpan=2 valign="top" bgcolor="#DBDBDB"></TD>
	</TR>
	<TR bgColor=#eeeeee> 
		<TD width="548" height="32"  colSpan=2><font color="#000000">Detailed materials of customer indent </font></TD>
	</TR>
	<TR> 
		<TD  width=126 bgColor=#DBDBDB height=12 align="right"><font color="#000000">OrderID：</font></TD>
		<TD  width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Form_Id")%></TD>
	</TR>
	<TR> 
		<TD bgColor=#DBDBDB height=12 align="right">CompanyName<font color="#000000">：</font></TD>
		<TD width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Comane")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">Consignee：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Somane")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">ReceiveAddress：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Add")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">Zipcode：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Zip")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=12 align="right"><font color="#000000">Telephone：</font></TD>
		<TD width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Phone")%></TD>
	</TR>
	<TR> 
		<TD bgColor=#DBDBDB height=12 align="right">Fax<font color="#000000">：</font></TD>
		<TD width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Fox")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">Email：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Email")%></TD>
	</TR>
<!-- 	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">所选汇款账号：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Pays")%></TD>
	</TR>
	<TR> 
		<TD width=126 height=25 align="right" bgColor=#DBDBDB><font color="#000000">备注：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Remark")%></TD>
	</TR> -->
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=24 align="right"><font color="#000000">OrderDate：</font></TD>
		<TD width=410 height=24 bgcolor="#eeeeee">&nbsp; <%=rs("RegTime")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">Whether the order has already been dealt with：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee"
>&nbsp; <%If rs("Flag")="尚未处理" Then%>
Have not dealt with yet  
<%else%>
Have already delivered 
<%End If%></TD>
	</TR>
	<TR> 
		<TD width="548" height="31"  colSpan=2 valign="top" bgcolor="#eeeeee"
><p align="center">Order<font color="#000000">Product</font>Minutia</p></TD>
	</TR>
<%
set rs2=server.createobject("adodb.recordset")
sqltext2="select * from Bs_ShopList where Form_Id=" & id
rs2.open sqltext2,conn,1,1
%>
	<TR> 
		<TD width="548" height="15"  colSpan=2 valign="top" bgcolor="#eeeeee"><div align="center"> 
			<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#006699" bordercolordark="#eeeeee"  height="67">
				<TR> 
					<td width="17%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">ProductID</font></td>
					<td width="41%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">ProductName</font></td>
					<td width="28%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">ProductClass</font></td>
					<td width="14%" height="21" align="center" bgcolor="#DBDBDB"><font color="#000000">ProductQuantity</font></td>
				</TR>
<%Sum=0
While Not rs2.EOF%>
				<TR> 
					<td width="17%" align="center" height="22"><%=rs2("Product_Id")%></td>
					<td width="41%" align="center" height="22"><%=rs2("Product_Name")%></td>
					<td width="28%" align="center" height="22"><%=rs2("BigClassName")%> 
=> <%=rs2("SmallClassName")%></td>
					<td height="22" align="center"><%=rs2("Number")%></td>
<%Sum=Sum+rs2("P_NewPrice")*rs2("Number")%>
				</TR>
<%
rs2.MoveNext
Wend
%>
				<TR> 
					<td colspan="4" height="22"> <p align="right">&nbsp;</p></td>
				</TR>
			</table>
		</div></TD>
	</TR>
	<center>
	<TR> 
		<TD width="548" height="25"  colSpan=2 bgcolor="#eeeeee"> <p align="center"
><input type="button" value="Return" name="B4" onClick="javascript:window.history.go(-1)"></TD>
	</TR>
	<TR bgColor=#DBDBDB> 
		<TD width="548" height="3"  colSpan=2></TD>
	</TR>
</TABLE>
</div>

</form>

<p>
<%
Else
response.redirect "En_Loginsb.asp?msg=You can look over that does not belong to your order , please import your one own order symbol again! "
End If
Else
response.redirect "En_Loginsb.asp?msg=The order number that you input does not exist or the form is incorrect, please input again! "
End IF

End if
rs3.close
set rs3=nothing
call CloseConn()
%>

</body>
</html>
