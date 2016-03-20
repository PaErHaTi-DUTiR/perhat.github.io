<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/En_System.asp"-->
<!--#include file="Include/En_SysHonor.asp"-->
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
'请勿改动下面这二行代码
MaxPerPage=9   '每页显示多少条信息
strFileName="En_Honor.asp?Action="& Action &"&"
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
<% Call Left_Honor() %>
			</td>
		</tr>
	</table> 
</TD>
<%end if%> 
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 公司荣誉 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align='center' class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
<%
if Action="Honor" then
	Response.write "Company honor"
elseif Action="Img" then
	Response.write "Corporate image"
end if
%>
&nbsp;</SPAN></TD>
						<TD class='MC_Pt_TD4'><div align="right">
<%
if Action="Honor" then
	call ShowHonorTotal()
elseif Action="Img" then
	call ShowImgTotal()
end if
%>
&nbsp;</div></TD>
						<TD class='MC_Pt_TD5'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_HSJNr_Title'>
					<TR>
						<TD valign=top>
<%
if Action="Honor" then
	call ShowHonorInfo()
elseif Action="Img" then
	call ShowImgInfo()
end if
%>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<!-- 公司荣誉 END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0 align='center' class='MC_bottom_title'>
					<tr> 
						<td>
<%
	if totalput>0 then
	call ShowNewsPage(strFileName,totalput,MaxPerPage,false,true,"A info")
	end if
%>
						</td>
					</tr>
					<tr><td class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR> 
			<TD height=5><IMG height=1 src="img/1x1_pix.gif" width=10></TD>
		</TR>
	</TABLE>
</TD>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Honor() %>
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