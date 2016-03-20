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
'response.redirect "orderlogin.asp"
Else
set Rs3 = Server.CreateObject("ADODB.recordset")
sql3="select * from Bs_OrderList where Form_Id="&Form_Id&""
rs3.open sql3,conn,1,1
IF  rs3.RecordCount >=1  then
IF Session("UserName")=rs3("UserName") Then
%>
<html>
<head>
<title>客户订货单详细信息</title>
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
		<TD width="548" height="32"  colSpan=2><font color="#000000">客户订货单详细资料</font></TD>
	</TR>
	<TR> 
		<TD  width=126 bgColor=#DBDBDB height=12 align="right"><font color="#000000">订货单号：</font></TD>
		<TD  width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Form_Id")%></TD>
	</TR>
	<TR> 
		<TD bgColor=#DBDBDB height=12 align="right">公司名称<font color="#000000">：</font></TD>
		<TD width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Comane")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">收货人姓名：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Somane")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">收货地址：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Add")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">邮政编码：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Zip")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=12 align="right"><font color="#000000">联系电话：</font></TD>
		<TD width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Phone")%></TD>
	</TR>
	<TR> 
		<TD bgColor=#DBDBDB height=12 align="right">联系传真<font color="#000000">：</font></TD>
		<TD width=410 height=12 bgcolor="#eeeeee">&nbsp; <%=rs("Fox")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">电子信箱：</font></TD>
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
		<TD width=126 bgColor=#DBDBDB height=24 align="right"><font color="#000000">订货日期：</font></TD>
		<TD width=410 height=24 bgcolor="#eeeeee">&nbsp; <%=rs("RegTime")%></TD>
	</TR>
	<TR> 
		<TD width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">订单是否已经处理：</font></TD>
		<TD width=410 height=25 bgcolor="#eeeeee"
>&nbsp; <%If rs("Flag")="尚未处理" Then%>
尚未处理 
<%else%>
已经发货 
<%End If%></TD>
	</TR>
	<TR> 
		<TD width="548" height="31"  colSpan=2 valign="top" bgcolor="#eeeeee"
><p align="center">订货<font color="#000000">产品</font>细目</p></TD>
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
					<td width="17%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">产品编号</font></td>
					<td width="41%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">产品名称</font></td>
					<td width="28%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">产品类别</font></td>
					<td width="14%" height="21" align="center" bgcolor="#DBDBDB"><font color="#000000">产品数量</font></td>
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
><input type="button" value="返回" name="B4" onClick="javascript:window.history.go(-1)"></TD>
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
response.redirect "Bs_Loginsb.asp?msg=您不能查看不属于您的订单，请重新输入您自己的订单号！"
End If
Else
response.redirect "Bs_Loginsb.asp?msg=您输入的订单号不存在或格式不正确，请重新输入！"
End IF

End if
rs3.close
set rs3=nothing
call CloseConn()
%>

</body>
</html>
