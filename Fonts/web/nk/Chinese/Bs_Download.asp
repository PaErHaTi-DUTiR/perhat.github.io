<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
<!--#include file="../Inc/ubbcode.asp"-->
<!--#include file="Include/Bs_SysDown.asp"-->
<!--#include file="Include/Bs_StrLeft.asp"-->
<%
'������������������������������������������������������������
'�������������� COPYRIGHT ������������ ��
'���������ƣ���ҵ��վ����ϵͳMac3.0  (ASBDcorpweb Mac3.0)  �� 
'����Ȩ���У�WWW.ASBD.CN  BBS.ASBD.CN                      ��
'������������amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  �� 
'��Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ��
'������������������������������������������������������������
%>
<%
'����Ķ����������д���
ShowSmallClassType=ShowSmallClassType_Default
MaxPerPage=MaxPerPage_Downt
strFileName="Bs_download.asp?Bs_BigClassName=" & Bs_BigClassName & "&Bs_SmallClassName=" & Bs_SmallClassName 
%>
<!--#include file="Include/Bs_Head.asp" -->
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Download() %>
			</td>
		</tr>
	</table> 
</TD>
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<%end if%>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '����� %>

<!-- �������� -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align='center' class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
<% call ShowClass_DownGuide() %>
&nbsp;</SPAN></TD>
						<TD class='MC_Pt_TD4'><div align="right">
<% call ShowDownTotal() %>
&nbsp;</div></TD>
						<TD class='MC_Pt_TD5'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_DoNr_Title'>
					<tr> 
						<td align='center'valign=top style="line-height: 150%">
<% call ShowDown(32) %>
						</td>
					</tr>
				</table>
			</TD>
		</TR>
<!-- �������� END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_bottom_title'>
					<tr> 
						<td>
<%
if totalput>0 then
call showpage(strFileName,totalput,MaxPerPage,false,true,"������")
end if
%> 
						</td>
					</tr>
					<tr><td class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR>
		<TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	</TABLE>
</TD>
<%if strSkins <= 200 and strSkins >= 1 then%>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>