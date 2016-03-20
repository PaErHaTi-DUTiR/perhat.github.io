<!--#include file="Include/En_System.asp"-->
<!--#include file="../Inc/Eshopcode.asp"-->
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
Sub PutToShopBag( cpbm, ProductList )
If Len(ProductList) = 0 Then
ProductList = "'" & cpbm & "'"
ElseIf InStr( ProductList, cpbm ) <= 0 Then
ProductList = ProductList & ", '" & cpbm & "'"
End If
End Sub


sql="select * from home"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1

ProductList = Session("ProductList")
flags="尚未处理"
If Len(ProductList) = 0 Then
Response.Redirect "nothing.asp"
end if

'保存购买人信息
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_OrderList"
rs.open sqltext,conn,3,3,1

'添加一个用户到数据库
rs.addnew
rs("username")=request.form("username")
rs("Comane")=request.form("Comane")
rs("Somane")=request.form("Somane")
rs("Add")=request.form("Add")
rs("Zip")=request.form("Zip")
rs("Phone")=request.form("Phone")
rs("Fox")=request.form("Fox")
rs("Email")=request.form("Email")
'rs("Pays")=request.form("Pays")
rs("Flag")=flags
'rs("Remark")=request.form("Remark")
'If rs("Remark")="" then
'rs("Remark")="无"
'End If
rs.update
rs.close
sql="select Form_Id from Bs_OrderList order by regtime desc"
rs.open sql,conn,3,3
a=rs(0)
rs.close
%>
<%
products=split(request("cpbm"),",")
for i=0 to UBound(Products)
'response.write cpbm
next
set rs=server.createobject("adodb.recordset")
session("productlist")=productlist
sql = "Select * from Bs_Product"
sql = sql & " Where Product_Id In (" & ProductList & ")"
sql = sql & " Order By ArticleID"
Set rs = conn.Execute( sql )
%>
<%
if session(rs("Product_Id"))=0 then
session(rs("Product_Id"))=1
end if

while not rs.eof
set rs2=server.createobject("adodb.recordset")
sqltext2="select * from Bs_ShopList"
rs2.open sqltext2,conn,3,3,1
rs2.addnew
rs2("Product_Id")=rs("Product_Id")
rs2("Form_Id")=a
rs2("Product_Name")=rs("EnTitle")
rs2("BigClassName")=rs("EnBigClassName")
rs2("SmallClassName")=rs("EnSmallClassName")
rs2("Number")=cint(session(rs("Product_Id")))
'rs2("P_NewPrice")=ccur(rs("P_NewPrice"))
rs2.update
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<title>It is successful to refer </title>
</head>
<BODY bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br>
<div align="center">
<TABLE width="500" height="300" align="center" cellPadding=0 cellSpacing=0>
	<TR>
		<TD align=middle width="100%">
			<TABLE cellSpacing=0 cellPadding=0 width=100% bgColor=#999999>
				<TR>
					<TD align="center" height=22 width="100%" bgcolor="#999999"><b><font color="#FFFFFF">It is successful to buy </font></b> </TD>
				</TR>
			</TABLE>
			<div align="center">
			<TABLE cellSpacing=5 width=100% bgColor=#999999 height="123" >
				<TR vAlign=top bgColor=#eeeeee>
					<TD  width="100%" height="50" bgcolor="#CCCCCC"
					>The order is referred successfully, your order number is ：<b><font size="4"><%=a%></font></b><br
					>
				  Please keep your order number firmly in mind, in order to inquire about. </TD>
				</TR>
				<TR>
					<td height="10"> </TD>
				</TR>
				<TR bgColor=#eeeeee>
					<TD  width="100%" height="27" bgcolor="#CCCCCC"> 
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td width="100%"><%=ubbcode(rs_home("content"))%></td>
							</tr>
						</table>
						<p align="center"> 
							<input type="button" name="close" value="CloseWindow "  onClick="window.close()"><br>
							<br>
						</p>
				  </TD>
				</TR>
			</TABLE>
			</div>
		</TD>
	</TR>
</TABLE><br>
</div>

</BODY></HTML>
<%
rs2.close
rs.movenext
wend
rs.close
rs_home.close
conn.close
set conn=nothing
%>
