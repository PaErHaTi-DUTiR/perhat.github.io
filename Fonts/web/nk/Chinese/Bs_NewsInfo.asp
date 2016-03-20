<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
<!--#include file="../Inc/ubbcode.asp"-->
<!--#include file="Include/Bs_SysNews.asp"-->
<!--#include file="Include/Bs_StrLeft.asp"-->
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<!--#include file="Include/Bs_Head.asp" -->
<SCRIPT language=JavaScript>
var currentpos,timer; 
function initializeScroll() { timer=setInterval("scrollwindow()",80);} 
function scrollclear(){clearInterval(timer);}
function scrollwindow() {currentpos=document.body.scrollTop;window.scroll(0,++currentpos);if (currentpos != document.body.scrollTop) sc();} 
document.onmousedown=scrollclear
document.ondblclick=initializeScroll
</SCRIPT>
<SCRIPT language=javascript>
function ContentSize(size)
{
	var obj=document.all.BodyLabel;
	obj.style.fontSize=size+"px";
}
</SCRIPT>

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
<% Call Left_News() %>
			</td>
		</tr>
	</table> 
</TD> 
<%end if%>
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<%end if%>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 新闻 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
<%
if Action="Co" then
	Response.write "公司新闻"
elseif Action="Ye" then
	Response.write "业内资讯"
elseif Action="Pr" then
	Response.write "产品动态"
end if
%>
&nbsp;</SPAN></TD>
						<TD class='MC_Pt_TD4'><div align="right">
【 字体：<A href="javascript:ContentSize(16)">大</A
><A href="javascript:ContentSize(14)">中</A
><A href="javascript:ContentSize(12)">小</A
> 】【<a href='javascript:window.print()'>打印此页</a>】
&nbsp;</div></TD>
						<TD class='MC_Pt_TD5'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_NewsNr_Title'>
					<TR> 
						<TD valign=top>
<%
if Action="Co" then
	Call M_NewsCo()
elseif Action="Ye" then
	Call M_NewsYe()
elseif Action="Pr" then
	Call M_NewsPr()
end if
%> 
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<!-- 新闻 END -->
		<TR> 
			<TD>
				<table cellpadding=0 cellspacing=0 align="center" class='MC_bottom_title'>
					<tr> 
						<td><div align="right">
【 字体：<A href="javascript:ContentSize(16)">大</A
><A href="javascript:ContentSize(14)">中</A
><A href="javascript:ContentSize(12)">小</A
> 】【<a href='javascript:window.print()'>打印此页</a
>】&nbsp;【<a href='javascript:history.back()'>返回</a
>】【<A href="javascript:window.scroll(0,-360)">顶部</A
>】【<a href="javascript:self.close()">关闭</a>】
&nbsp;</div></td>
					</tr>
					<tr><td  class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	</TABLE>
</TD>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_News() %>
			</td>
		</tr>
	</table> 
</TD>

</TR>
</TABLE>
<!--#include file="Include/Bs_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>

