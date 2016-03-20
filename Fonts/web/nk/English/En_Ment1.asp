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
<HTML><HEAD><TITLE>Settlement of the goods</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Css.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function FORM1_onsubmit()
{
if (document.FORM1.Comane.value.length==0)
{
alert("Please input the correct company name. ");
document.FORM1.Comane.focus();
return false;
}
if (document.FORM1.Somane.value.length==0)
{
alert("Please input correct consignee name. ");
document.FORM1.Somane.focus();
return false;
}
if (document.FORM1.Add.value.length==0)
{
alert("Please input the correct address of receiving. ");
document.FORM1.Add.focus();
return false;
}
if (document.FORM1.Zip.value.length==0)
{
alert("Please input the correct zipcode. ");
document.FORM1.Zip.focus();
return false;
}
if (document.FORM1.Phone.value.length==0)
{
alert("Please input the correct telephone number. ");
document.FORM1.Phone.focus();
return false;
}
if (document.FORM1.Email.value.length==0)
{
alert("Please input the correct connection Email. ");
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
action=En_ment2.asp method=post>
<INPUT type=hidden value=<%=rs("username")%> name=username>
	<TABLE width=700 align="center" cellPadding=4 cellSpacing=1 bgColor=#666666>
		<TR vAlign=top bgColor=#999999> 
		  <TD height="22"  colSpan=2 background="../Chinese/images/b1.gif"><div align="center"><font color="#FFFFFF">Settle account in shopping --(the first step) consignee information</font></div></TD>
		</TR>
		<TR bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">UserName：</TD>
			<TD width=500 height=25> <%=rs("username")%> </TD>
		</TR>
		<TR bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">CompanyName：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=13 value="<%=rs("Comane")%>" name=Comane style="font-size: 14px"></TD>
		</TR>
		<tr bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">Consignee：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=36 value="<%=rs("Somane")%>" name=Somane style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">ReceiveAddress：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=13 value="<%=rs("Add")%>" name=Add style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">Zipcode：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=23 value="<%=rs("Zip")%>" name=Zip style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">Telephone：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=32 value="<%=rs("Phone")%>" name=Phone style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">Fax：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=32 value="<%=rs("Fox")%>" name=Fox style="font-size: 14px"></TD>
		</tr>
		<tr bgcolor="#CCCCCC"> 
			<TD width=200 height=25 align="right">Email：</TD>
		  <TD width=500 height=25> <INPUT maxLength=16 size=32 value="<%=rs("Email")%>" name=Email style="font-size: 14px"></TD>
		</tr>
		<TR vAlign=top bgColor=#CCCCCC> 
		  <TD colSpan=2 ><div align="center">You can revise the above content </div></TD>
		</TR>
		<TR bgColor=#999999> 
			<TD height="22"  colSpan=2> <DIV align=center> 
			<INPUT type="button" value="Previous" name="B4" onClick="javascript:window.history.go(-1)">
			&nbsp;&nbsp; 
			<INPUT type="submit" size=3 value="Next" name="Submit2">
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
