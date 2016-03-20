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
Sub PutToShopBag( cpbm, ProductList )
If Len(ProductList) = 0 Then
ProductList = "'" & cpbm & "'"
ElseIf InStr( ProductList, cpbm ) <= 0 Then
ProductList = ProductList & ", '" & cpbm & "'"
End If
End Sub



ProductList = Session("ProductList")
Products = Split(Request("cpbm"), ", ")
For I=0 To UBound(Products)
PutToShopBag Products(I), ProductList
Next
Session("ProductList") = ProductList


Head="以下是您所选购的产品清单"
ProductList = Session("ProductList")
If Len(ProductList) =0 Then
Response.Redirect "Bs_Search.asp"
response.end
end if

If Request("MySelf") = "Yes" Then
ProductList = ""
Products = Split(Request("cpbm"), ", ")
For I=0 To UBound(Products)
PutToShopBag Products(I), ProductList
Next
Session("ProductList") = ProductList
End If
If Len(ProductList) = 0 Then
Response.Redirect "Bs_Product.asp"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql = "Select * from Bs_Product"
sql = sql & " Where Product_Id In (" & ProductList & ")"
rs.open sql,conn,3,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>以下是您所选购的物品清单</title>
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
</head>
<script language="Javascript">
//函数名：fucCheckNUM
//功能介绍：检查是否为数字
//参数说明：要检查的数字
//返回值：1为是数字，0为不是数字
function fucCheckNUM(NUM)
{
var i,j,strTemp;
strTemp="0123456789";
if ( NUM.length== 0)
return 0
for (i=0;i<NUM.length;i++)
{
j=strTemp.indexOf(NUM.charAt(i));	
if (j==-1)
{
//说明有字符不是数字
return 0;
}
}
//说明是数字
return 1;
}

function clean()
{
window.location.href="Bs_Clear.asp"
}
</script>
<body bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="center"> <br>
  <table width="80%" border="0" align="center" cellspacing="0">
    <tr>
<td width="80%" valign="top"><p align="center">
</p>
<p align="center">
<font color="#0000FF" ><%=Head%></font></p>
<script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  var checkOK = "0123456789-";
  var checkStr = theForm.<%="Q_" & rs("Product_Id")%>.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  if (!allValid)
  {
    alert("在 请输入正确的产品数量！ 域中，只能输入 数字 个字符。");
    theForm.<%="Q_" & rs("Product_Id")%>.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
				<form Action="Bs_Eshop.asp" Method="POST" onSubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
          <input type="hidden" name="MySelf" value="Yes">
          <div align="center"><center>
              <table border="0" cellspacing="1" width="100%"  cellpadding="0" bgcolor="#666666">
                <tr> 
                  <td align="center" width="132"  height="22" bgcolor="#999999"><font color="#FFFFFF">产品编号</font></td>
                  <td align="center" width="421"  height="22" bgcolor="#999999"><font color="#FFFFFF">产品名称</font></td>
                  <td align="center" width="128" height="22" bgcolor="#999999"><font color="#FFFFFF">产品数量</font></td>
                  <td align="center" width="119"  height="22" bgcolor="#999999"><font color="#FFFFFF">购买</font></td>
                </tr>
                <%
Sum = 0
While Not rs.EOF
Quatity = CLng( Request( "Q_" & rs("Product_Id")) )
If Quatity <= 0 Then
Quatity = CLng( Session(rs("Product_Id")) )
If Quatity <= 0 Then Quatity = 1
End If
Session(rs("Product_Id")) = Quatity
Sum = Sum + ccur(rs("P_NewPrice")) * Quatity
%>
                <tr> 
                  <td align="center" width="132" bgcolor="#F0FCFF"><%=rs("Product_Id")%> 
                  </td>
                  <td align="center" width="421" bgcolor="#F0FCFF"><%=rs("Title")%> 
                  </td>
                  <td align="center" width="128" bgcolor="#F0FCFF"> 
                    <!--webbot bot="Validation"
S-Display-Name="请输入正确的产品数量！" S-Data-Type="Integer"
S-Number-Separators="x" -->
                    <input Name="<%="Q_" & rs("Product_Id")%>" Value="<%=Quatity%>"  size="8" maxlength="12"> 
                  </td>
                  <td align="center" width="119" bgcolor="#F0FCFF"><input Type="CheckBox" Name="cpbm" Value="<%=rs("Product_Id")%>" Checked> 
                  </td>
                </tr>
                <%
rs.MoveNext
Wend
%>
                <tr bgcolor="#F0FCFF"> 
                  <td height="22" ColSpan="4" Align="Right">&nbsp;</td>
                </tr>
              </table>
</center></div><blockquote>
<p align="center">
<input Type="submit" Value="更改数量" name="B1" style="font-size: 9pt">&nbsp;&nbsp;&nbsp;
<input type="button" value="继续购物" name="B2" onClick="window.close();" style="font-size: 9pt">&nbsp;&nbsp;&nbsp;
<input type="button" value="订单取消" name="B3" OnClick="clean()" style="font-size: 9pt">&nbsp;&nbsp;&nbsp;
<input type="button" value="去收银台" name="B4" onClick="location.href='Bs_Ment.asp'" style="font-size: 9pt" >&nbsp;&nbsp;&nbsp;
<input type="button" value="关闭" name="B5" onClick="window.close();" style="font-size: 9pt">     <p align="center"><font color=FF0000;> 注意：改变“产品数量”与"购买",需点击“更改数量”按钮！   </font>   </blockquote>
</form>
</td></tr></table>
</div>
<%
rs.close
set rs=nothing
call CloseConn()
%>
</body>
</html>
