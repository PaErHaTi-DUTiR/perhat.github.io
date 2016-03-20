<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/En_System.asp"-->
<!--#include file="Include/En_StrLeft.asp"-->
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<%If Session("UserName")="" Then
Response.Redirect"En_Server.asp"
end if%>

<SCRIPT language=javascript>
function chkitem(str) {	
	var strSource ="0123456789-";
	var ch;
	var i;
	var temp;
	for (i=0;i<=(str.length-1);i++)	{
		ch = str.charAt(i);
		temp = strSource.indexOf(ch);
		if (temp==-1)
		{
		return 0;
		}
	}
	if (strSource.indexOf(ch)==-1) {
		return 0;
		}
		else
		{
		return 1;
	}
}

function FORM1_onsubmit() {
	if (chkitem(document.FORM1.Form_Id.value)==0) {
		alert("Please input the correct order number. ");
		document.FORM1.Form_Id.focus();
		return false;
	}
}
</SCRIPT>

<!--#include file="Include/En_Head.asp" -->
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR> 
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Member() %>
			</td>
		</tr>
	</table> 
</TD>

<%if strSkins <= 200 and strSkins >= 1 then%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<%end if%>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 订单查询 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align="center"  class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
Order inquiry 
&nbsp;</SPAN></TD>
						<TD class='MC_Pt_TD4'><div align="right">
&nbsp;</div></TD>
						<TD class='MC_Pt_TD5'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_ServeNr_Title'>
					<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
					<TR vAlign=top > 
						<TD  width="100%" height="18">
						<form action="En_Orderofind.asp" method="post" name=FORM1 onSubmit="return FORM1_onsubmit()" language=javascript>
<p align="center"> <font size="4" face="Arial"><%=Session("UserName")%></font><font color="#FF0000">&nbsp;&nbsp;howdy</font></p>
<p style="word-spacing: 0; margin-left: 10; margin-right: 10; margin-top: 0; margin-bottom: 0" align="center">The following is the order that you have already refer been in our stationed.  
<%
dim Id,sql,rs
Id=Session("UserName")
set Rs = Server.CreateObject("ADODB.recordset")
sql="select * from Bs_OrderList where UserName='"&Id&"'"
rs.open sql,conn,1,1
%>
</p>
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"
>Order detailed materials inquire about , <BR>please input you want order number that inquire about <input type="text" name="Form_Id" size="11">
<input type="submit" value="Inquiry" name="B1">
						</form>
							<div align="center"> 
							<TABLE cellSpacing=0 cellPadding=0 align="center"  width=100%>
								<TR> 
									<TD align=middle width="100%"> <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">The following is that all orders which you refer in our station are tabulated </p>
										<table border="0" cellpadding="2" cellspacing="1" width="100%" bgcolor="#000000">
											<tr > 
												<td width="26%" align="center" height="20" bgcolor="#CCCCCC"><font color="#000000">OrderID</font></td>
												<td width="21%" align="center" height="20" bgcolor="#CCCCCC"><font color="#000000">CustomerName </font></td>
												<td width="53%" align="center" height="20" bgcolor="#CCCCCC"><font color="#000000">Situation of treatment to your order</font></td>
											</tr>
<% While Not rs.EOF %>
											<tr > 
												<td width="26%" align="center" height="25" bgcolor="#FFFFFF"><%=rs("Form_Id")%></td>
												<td width="21%" align="center" height="25" bgcolor="#FFFFFF"><%=rs("Somane")%></td>
												<td width="53%" align="center" height="25" bgcolor="#FFFFFF">
<%If rs("Flag")="尚未处理" Then%>
Your fund has not reached our account yet, this order our station has not dealt with yet 
<%else%> <font color="#FF0000">This order has already delivered to you, please pay attention to checking and accept! </font> 
<%End If%>
												</td>
											</tr>
<%
rs.MoveNext
Wend
%>
										</table>
									</TD>
								</TR>
							</TABLE>
							</div>
						</TD>
					</TR>
					<TR vAlign=top > 
						<TD  width="100%" height="15"> &nbsp;&nbsp; </TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<%
rs.close
set rs=nothing
%>
<!-- 订单查询 End -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_bottom_title'>
					<tr> 
						<td>
						</td>
					</tr>
					<tr><td class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	</TABLE>
</TD>
<%if strSkins <= 200 and strSkins >= 1 then%>

</TR>
</TABLE>
<!--#include file="Include/En_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot4.asp" -->
<%end if%>