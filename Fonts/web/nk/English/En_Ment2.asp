<!--#include file="../Inc/Conn.asp" -->
<%
'=========================================================
'
'产品名称： 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有：www.web300.cn
'程序制作：QQ812256@hotmail.com
'Copyright 2006-2008 www.web300.cn - All Rights Reserved.
'
'========================================================
%>
<%
Id=Session("username")
ProductList = Session("ProductList")
Products = Split(Request("cpbm"), ", ")
For I=0 To UBound(Products)
PutToShopBag Products(I), ProductList
Next
Session("ProductList") = ProductList

ProductList = Session("ProductList")
If Len(ProductList) =0 Then
Response.Redirect "En_Product.asp"
response.end
end if

set rs=server.createobject("adodb.recordset")
sql = "Select * from Bs_Product"
sql = sql & " Where Product_Id In (" & ProductList & ")"
rs.open sql,conn,3,3

%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<TITLE>Settlement of the goods</TITLE>
</HEAD>
<BODY bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br>
<FORM name="FORM1" action=En_ment3.asp method=post>
<input type=hidden Name="username" Value="<%=Request.form("username")%>" >
<input type=hidden Name="Comane" Value="<%=Request.form("Comane")%>" >
<input type=hidden Name="Somane" Value="<%=Request.form("Somane")%>" >
<input type=hidden Name="Add" Value="<%=Request.form("Add")%>" >
<input type=hidden Name="Zip" Value="<%=Request.form("Zip")%>" >
<input type=hidden Name="Phone" Value="<%=Request.form("Phone")%>" >
<input type=hidden Name="Fox" Value="<%=Request.form("Fox")%>" >
<input type=hidden Name="Email" Value="<%=Request.form("Email")%>" >

<TABLE cellSpacing=0 cellPadding=0 width=100%>
	<TR>
		<TD align=middle width="100%">
			<div align="center">
			<center>
			<TABLE width=680 align="center" cellSpacing=1 bgColor=#666666 >
				<TR bgColor=#999999> 
				  <TD height="22"  colSpan=2 background="../Chinese/images/b1.gif"><div align="center"><font color="#FFFFFF">Settle account in shopping --(the second step) information and term of payment that do shopping</font></div></TD>
				</TR>
				<tr bgcolor="#CCCCCC"> 
					<TD height=22 colspan="2" align="right">
						<div align="center"><br>
						<table border="0" cellspacing="1" width="580"  height="61" bgcolor="#666666" bordercolorlight="#000000" bordercolordark="#FFFFFF">
							<tr bgcolor="#006699"> 
								<td align="center" width="117"  height="22" bgcolor="#999999"><font color="#FFFFFF">ProductID</font></td>
								<td align="center" width="335"  height="22" bgcolor="#999999"><font color="#FFFFFF">ProductName</font></td>
								<td align="center" width="118" height="22" bgcolor="#999999"><font color="#FFFFFF">ProductQuantity</font></td>
							</tr>
<%
Sum = 0
While Not rs.EOF
Quatity = CInt( Request( "Q_" & rs("Product_Id")) )
If Quatity <= 0 Then
Quatity = CInt( Session(rs("Product_Id")) )
If Quatity <= 0 Then Quatity = 1
End If
Session(rs("Product_Id")) = Quatity
Sum = Sum + ccur(rs("P_NewPrice")) * Quatity
%>
							<tr> 
								<td align="center" width="117" height="23" bgcolor="#EEEEEE"><%=rs("Product_ID")%></td>
								<td align="center" width="335" height="23" bgcolor="#EEEEEE"><%=rs("EnTitle")%></td>
								<td align="center" width="118" height="23" bgcolor="#EEEEEE"><%=Quatity%></td>
							</tr>
<%
rs.MoveNext
Wend
%>
							<tr> 
							<td Align="Right" ColSpan="3" height="22" bgcolor="#EEEEEE">&nbsp;</td>
							</tr>
						</table>  
						</div>
				  </TD>  
				</tr>  
<!-- 				<tr>  
					<TD  width=230 bgcolor=#F0FCFF height=7>我们提供三家银行的账号供您汇款,请您选择:</TD>  
					<TD  width=450 height=22 bgcolor="#F0FCFF">
<select size="1" name="Pays" style="font-size: 14px">  
<option>招商银行一卡通</option> 
<option>交通银行太平洋卡</option> 
<option>建设银行龙卡储蓄卡</option> 
</select>
					</TD> 
				</tr>  -->
<!-- 				<tr> 
					<TD width=230 bgcolor=#F0FCFF height=7>订单备注:<p>　</p></TD>
					<TD width=450 height=22 bgcolor="#F0FCFF"><textarea rows="6" name="Remark" cols="45" style="font-size: 10pt"></textarea></TD>
				</tr> -->
				<TR bgColor=#CCCCCC> 
					<TD height="22"  colSpan=2><DIV align=center>&nbsp; 
<input type="button" value="Previous" name="B4" onClick="javascript:window.history.go(-1)">
<INPUT type="submit" size=3 value="Next" name="Submit2">  
				  </DIV></TD>
				</TR>
			</TABLE>  
			</center>  
			</div>  
		</TD>
	</TR>
</TABLE>
<BR></FORM>  
</BODY></HTML>  
<%
rs.close
set rs=nothing
call CloseConn()
%>
