<!--#include file="Include/En_System.asp"-->
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
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<%end if%><TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '����� %>

<!-- ��Ϣ�ύ -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align='center' class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
Information refering
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
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_ServeNr_Title'>
		<TR> 
			<TD><div align="center"> 
				<table width="80%" border="0" cellspacing="0" cellpadding="0">
					<tr> 
						<td><div align="center"><strong><br>OK ��</strong></div>
<p align="center"><font color="#FF0000">Information is referred successfully��</font></p>
<p></p>
<p> <br>
You are so already successful as to refer to your information, we will do the best to reply within 24 hours, because you are not our member, we can only pass the way of E-mail , reply your information, please check your mailbox .</p>
<p>Meanwhile, we advise you to register and become our member at once, only register and become our member, you could see our reply in member administrative center , but needn check and accept through E-mail . 
<br> 
<br>
<br>
<br>
</p></td>
					</tr>
				</table><br>
			</div></TD>
		</TR>
				</TABLE>
			</TD>
		</TR>
<!-- ��Ϣ�ύ End -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_bottom_title'>
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
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot4.asp" -->
<%end if%>