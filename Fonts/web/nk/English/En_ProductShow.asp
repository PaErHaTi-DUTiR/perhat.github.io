<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/En_System.asp"-->
<!--#include file="../Inc/ubbcode.asp"-->
<!--#include file="Include/En_SysProduct.asp"-->
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
<%
ShowSmallClassType=ShowSmallClassType_Article
dim ArticleID
ArticleID=trim(request("ArticleID"))
if ArticleId="" then
	response.Redirect("En_Product.asp")
end if

sql="select * from Bs_Product where ArticleID=" & ArticleID & ""
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	response.write"<SCRIPT language=JavaScript>alert('Can find this product！');" 'Can find this product 找不到此产品
  response.write"javascript:history.go(-1)</SCRIPT>"
else	
	rs("Hits")=rs("Hits")+1
	rs.update
	if rs("hits")>=HitsOfHot then
		rs("Hot")=True
		rs.update
	end if
	EnBigClassName=rs("EnBigClassName")
	EnSmallClassName=rs("EnSmallClassName")
%>
<!--#include file="Include/En_Head.asp" -->
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR>
<%if strSkins <= 200 and strSkins >= 1 then%>
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<%
Call Style_Product() '产品列表
%>
			</td>
		</tr>
	</table> 
</TD>
<%end if%>  
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 产品查看 -->
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
<%
response.write "<a href='En_Product.asp'>Product show</a>"
response.write "&nbsp;&gt;&gt;&nbsp;<a href='En_Product.asp?EnBigClassName=" & rs("EnBigClassName") & "'>" & rs("EnBigClassName") & "</a>"
if rs("SEnmallClassName") & ""<>"" then
response.write "&nbsp;&gt;&gt;&nbsp;<a href='En_Product.asp?EnBigClassName=" & rs("EnBigClassName")&"&EnSmallClassName=" & rs("EnSmallClassName") & "'>" & rs("EnSmallClassName") & "</a>"
end if
'response.write "&nbsp;&gt;&gt;&nbsp;"& rs("Title")
%>
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
				<TABLE cellSpacing=1 cellPadding=0  align='center' class='MC_PrShow_Title1'>
					<TR class='MC_PrShow_t1Tr1'>
					<td class='MC_PrShow_t1r1Td1'>Product Name</TD>
					<td>Spec</TD>
					<td class='MC_PrShow_t1r1Td3'>No.& Bar code</TD>
					</TR>
					<TR class='MC_PrShow_t1Tr2'>
					<td>&nbsp;<%=rs("EnTitle")%></TD>
					<td>&nbsp;<%=rs("EnSpec")%></TD>
					<td>&nbsp;<%=rs("Size")%></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_PrShow_Title2'>
					<TR class='MC_PrShow_t2Tr1'>
						<td>::Explanation::</TD>
						<TD width="98" align="right">
						<a href='javascript:eshop(<%=rs("Product_Id")%>)'><img border=0 src='../Img/En_addtocart.gif' width=95 height=17></a>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align='center' class='MC_PrShow_Title3'>
					<tr> 
						<td style="line-height: 150%">　　<%call ShowArticleContent()%></td>
					</tr>
				</table>
			</TD>
		</TR>
<!-- 产品查看 END -->
		<TR> 
			<TD>
				<table cellpadding=0 cellspacing=0  align='center' class='MC_bottom_title'>
					<tr> 
						<td><div align="right">
Click count：<%=rs("Hits")%>&nbsp; Input time：<%= FormatDateTime(rs("UpdateTime"),2) %>&nbsp;【<a href='javascript:window.print()'>Print Page</a
>】【<a href='javascript:history.back()'>Return</a
>】【<A href="javascript:window.scroll(0,-360)">top</A>】
&nbsp;</div></td>
					</tr>
					<tr><td class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	</TABLE>
</TD>
<%
end if
rs.close
set rs=nothing
%>

<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<%
Call Style_Product() '产品列表
%>
			</td>
		</tr>
	</table> 
</TD>

</TR>
</TABLE>
<!--#include file="Include/En_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot4.asp" -->
<%end if%>