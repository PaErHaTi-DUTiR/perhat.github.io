<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/En_System.asp"-->
<!--#include file="../Inc/ubbcode.asp"-->
<!--#include file="Include/En_SysProduct.asp"-->
<!--#include file="Include/En_SysCoSaJo.asp"-->
<!--#include file="Include/En_StrLeft.asp"-->
<%
'������������������������������������������������������������
'�������������� COPYRIGHT ������������ ��
'���������ƣ���ҵ��վ����ϵͳMac3.0  (ASBDcorpweb Mac3.0)  �� 
'����Ȩ���У�WWW.ASBD.CN  BBS.ASBD.CN                      ��
'������������amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  �� 
'��Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ��
'������������������������������������������������������������
%>
<!--#include file="Include/En_Head.asp" -->
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR>
<%if strSkins <= 200 and strSkins >= 1 then%>
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Sale() %>
			</td>
		</tr>
	</table> 
</TD>
<%end if%>  
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '����� %>

<!-- �����г� -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
<%
dim Action
Action=trim(request("Action"))
if Action="Sale" then
	Response.write "China market"
elseif Action="Salea" then
	Response.write "International market"
end if
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
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_HSJNr_Title'>
					<tr> 
						<td valign=top style="line-height: 150%">
<%if Action="Sale" then%>
<%=CSJ_Sale%>
<%elseif Action="Salea" then%>
<%=CSJ_Salea%>
<%end if%>
						</td>
					</tr>
				</table>
			</TD>
		</TR>
<!-- �����г� END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_bottom_title'>
					<tr> 
						<td>
						</td>
					</tr>
					<tr><td class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR> 
			<TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
		</TR>
	</TABLE> 
</TD>

<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>

<%if strSkins <= 200 and strSkins >= 1 then%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Sale() %>
			</td>
		</tr>
	</table> 
</TD>

</TR>
</TABLE>
<!--#include file="Include/En_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot4.asp" -->
<%end if%>