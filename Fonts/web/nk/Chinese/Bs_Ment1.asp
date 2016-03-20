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
set Rs = Server.CreateObject("ADODB.recordset")
sql="select * from Bs_User where username='"&Id&"'"
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>商品结算</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function FORM1_onsubmit()
{
if (document.FORM1.Comane.value.length==0)
{
alert("请输入正确的公司名称。");
document.FORM1.Comane.focus();
return false;
}
if (document.FORM1.Somane.value.length==0)
{
alert("请输入正确的收货人名字。");
document.FORM1.Somane.focus();
return false;
}
if (document.FORM1.Add.value.length==0)
{
alert("请输入正确的收货地址。");
document.FORM1.Add.focus();
return false;
}
if (document.FORM1.Zip.value.length==0)
{
alert("请输入正确的邮政编码。");
document.FORM1.Zip.focus();
return false;
}
if (document.FORM1.Phone.value.length==0)
{
alert("请输入正确的联系电话。");
document.FORM1.Phone.focus();
return false;
}
if (document.FORM1.Email.value.length==0)
{
alert("请输入正确的联系E-mail。");
document.FORM1.Email.focus();
return false;
}
}
</SCRIPT>
</HEAD>
<BODY bgcolor="#D9D9D9" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<DIV align=center>
<br>
<FORM language=javascript name=FORM1 onSubmit="return FORM1_onsubmit()"
action=Bs_Ment2.asp method=post>
<INPUT type=hidden value=<%=rs("username")%> name=username>
	<TABLE width=600 align="center" cellPadding=4 cellSpacing=1 bgColor=#666666>
		<TR vAlign=top bgColor=#999999> 
		  <TD height="22"  colSpan=2 background="images/b1.gif"><div align="center"><font color="#FFFFFF">购物结算--(第一步)收货人信息</font></div></TD>
		</TR>
		<TR bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">帐  号：</TD>
			<TD width=500 height=25> <%=rs("username")%> </TD>
		</TR>
		<TR bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">公司名称：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=13 value="<%=rs("Comane")%>" name=Comane style="font-size: 14px"></TD>
		</TR>
		<tr bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">收货人：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=36 value="<%=rs("Somane")%>" name=Somane style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">收货地址：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=13 value="<%=rs("Add")%>" name=Add style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">邮政编码：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=23 value="<%=rs("Zip")%>" name=Zip style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">联系电话：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=32 value="<%=rs("Phone")%>" name=Phone style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">联系传真：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=32 value="<%=rs("Fox")%>" name=Fox style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#C6C6C6"> 
			<TD width=200 height=25 align="right">Email：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=32 value="<%=rs("Email")%>" name=Email style="font-size: 14px"></TD>
		</tr>
		<TR vAlign=top bgColor=#F0FCFF> 
		  <TD colSpan=2 bgcolor="#C6C6C6" ><div align="center">您可以修改以上内容</div></TD>
		</TR>
		<TR bgColor=#999999> 
			<TD height="22"  colSpan=2> <DIV align=center> 
			<INPUT type="button" value="上一步" name="B4" onClick="javascript:window.history.go(-1)">
			&nbsp;&nbsp; 
			<INPUT type="submit" size=3 value="下一步" name="Submit2">
			</DIV></TD>
		</TR>
	</TABLE>
</FORM>
</DIV>

</BODY></HTML>
<%
rs.close
set rs=nothing
call CloseConn()
%>
