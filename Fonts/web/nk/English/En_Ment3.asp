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
<TITLE>Settlement of the goods</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<FORM name="FORM1" action=En_ment4.asp method=post>
<input type=hidden Name="username" Value="<%=Request.form("username")%>" >
<input type=hidden Name="Comane" Value="<%=Request.form("Comane")%>" >
<input type=hidden Name="Somane" Value="<%=Request.form("Somane")%>" >
<input type=hidden Name="Add" Value="<%=Request.form("Add")%>" >
<input type=hidden Name="Zip" Value="<%=Request.form("Zip")%>" >
<input type=hidden Name="Phone" Value="<%=Request.form("Phone")%>" >
<input type=hidden Name="Fox" Value="<%=Request.form("Fox")%>" >
<input type=hidden Name="Email" Value="<%=Request.form("Email")%>" >
<!-- <input type=hidden Name="Pays" Value="<%=Request.form("Pays")%>" >
<input type=hidden Name="Remark" Value="<%=Request.form("Remark")%>" > -->
  <TABLE width=680 align="center" cellPadding=4 cellSpacing=1 bgColor=#666666 1>
    
      <TR vAlign=top bgColor=#999999> 
        <TD height="22" colSpan=4 background="../Chinese/images/b1.gif"> <div align="center"><font color="#FFFFFF">Settle account in shopping --(the third step) information is confirmed</font></div></TD>
      </TR>
      <tr bgcolor="#F0FCFF"> 
        <TD height=22 colspan="4" align="right" bgcolor="#CCCCCC"> <div align="center"> 
            <center>
              <table border="0" cellpadding="0" cellspacing="1" width="580" height="61" bgcolor="#666666">
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
                  <td align="center" width="117" height="23" bgcolor="#EEEEEE"><%=rs("Product_ID")%> 
                  </td>
                  <td align="center" width="338" height="23" bgcolor="#EEEEEE"><%=rs("EnTitle")%> 
                  </td>
                  <td align="center" width="118" height="23" bgcolor="#EEEEEE"><%=Quatity%></td>
                </tr>
                <%
rs.MoveNext
Wend
%>
                <tr bgcolor="#F0FCFF"> 
                  <td height="22" ColSpan="3" Align="Right" bgcolor="#EEEEEE">&nbsp;</td>
                </tr>
              </table>
            </center>
          </div>
        </TD>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <TD width=160 height=22 align="right">Consignee：</TD>
        <TD height=22 colspan="3"><%=Request.form("Somane")%></TD>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <TD width=160 height=22 align="right">CompanyName：</TD>
        <TD height=22 colspan="3"><%=Request.form("Comane")%></TD>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <TD width=160 height=22 align="right">ReceiveAddress：</TD>
        <TD height=22 colspan="3"><%=Request.form("Add")%></TD>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <TD width=160 height=22 align="right">Zipcode：</TD>
        <TD width=180 height=22><%=Request.form("Zip")%></TD>
        <TD width=160 height=22> <p align="center">Telephone：</TD>
        <TD width=180 height=22><%=Request.form("Phone")%></TD>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <TD width=160 height=22 align="right">EMail：</TD>
        <TD height=22><%=Request.form("Email")%></TD>
        <TD height=22><div align="center">Fax：</div></TD>
        <TD height=22><%=Request.form("Fox")%></TD>
      </tr>
<!--       <tr> 
        <TD width=160 bgcolor="#CCCCCC" height=22 align="right">所选汇款账号：</TD>
        <TD height=22 bgcolor="#F0FCFF" colspan="3"><%=Request.form("Pays")%></TD>
      </tr>
      <tr> 
        <TD width=160 bgcolor="#CCCCCC" height=22 align="right">订单备注：</TD>
        <TD height=22 bgcolor="#F0FCFF" colspan="3"><%=Request.form("Remark")%></TD>
      </tr> -->
      <TR bgColor=#CCCCCC> 
        <TD colSpan=4> <DIV align=center> 
            <input type="button" value="Previous" name="B4" onClick="javascript:window.history.go(-1)">
            <INPUT  type="submit" value="ReferOrder" name="Submit2">
        </DIV></TD>
      </TR>
    
  </TABLE>
<BR></FORM>
</BODY></HTML>
<%
rs.close
set rs=nothing
call CloseConn()
%>
