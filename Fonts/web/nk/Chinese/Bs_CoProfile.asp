<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
<!--#include file="../Inc/ubbcode.asp"-->
<!--#include file="Include/Bs_SysCoSaJo.asp"-->
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
<!--#include file="Include/Bs_Head.asp" -->
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
<% Call Left_CoProfile() %>
			</td>
		</tr>
	</table> 
</TD>
<%end if%> 
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<%end if%>
<TD vAlign=top > 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '����� %>

<!-- ��˾���� -->
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
if Action="Profile" then
	Response.write "��˾����"
elseif Action="Ceo" then
	Response.write "�ܲý���"
elseif Action="Culture" then
	Response.write "��˾�Ļ�"
elseif Action="Organize" then
	Response.write "��֯����"
elseif Action="Principle" then
	Response.write "�쵼�ػ�"
elseif Action="Contact" then
	Response.write "��ϵ����"
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
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_CoNr_Title'>
					<tr> 
						<td style="line-height: 150%" valign=top>����
<%if Action="Profile" then%>
<%=CSJ_Profile%>
<%elseif Action="Ceo" then%>
<%=CSJ_Ceo%>
<%elseif Action="Culture" then%>
<%=CSJ_Culture%>
<%elseif Action="Organize" then%>
<%=CSJ_Organize%>
<%elseif Action="Principle" then%>
<%=CSJ_Principle%>
<%elseif Action="Contact" then%>
<%=CSJ_Contact%>
<%end if%>
						</td>
					</tr>
				</table>
			</TD>
		</TR>
<!-- ��˾���� END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0 class='MC_bottom_title' align='center'>
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
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'> 
		<tr>
			<td valign="top">
<% Call Left_CoProfile() %>
			</td>
		</tr>
	</table> 
</TD>

</TR>
</TABLE>
<!--#include file="Include/Bs_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Index() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>

