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
ShowSmallClassType=ShowSmallClassType_Article
dim Bs_DownID,Bs_Content
Bs_DownID=trim(request("Bs_DownID"))
if Bs_DownID="" then
	response.Redirect("Bs_Download.asp")
end if

sql="select * from Bs_Download where Bs_DownID=" & Bs_DownID & ""
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	response.write"<SCRIPT language=JavaScript>alert('�Ҳ����������');"
  response.write"javascript:history.go(-1)</SCRIPT>"
else	
	rs("Bs_Hits")=rs("Bs_Hits")+1
	rs.update
	if rs("Bs_Hits")>=HitsOfHot then
		rs("Bs_Hits")=True
		rs.update
	end if
	Bs_BigClassName=rs("Bs_BigClassName")
	Bs_SmallClassName=rs("Bs_SmallClassName")
	Bs_Content=ubbcode(rs("Bs_Content"))
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
<% Call Left_Download() %>
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
<!-- ���ز鿴 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
<%
Response.Write  "<a href='Bs_Download.asp'>��������</a>"
Response.Write "&nbsp;&gt;&gt;&nbsp;<a href='Bs_Download.asp?Bs_BigClassName=" & rs("Bs_BigClassName") & "'>" & rs("Bs_BigClassName") & "</a>"
if rs("Bs_SmallClassName") & ""<>"" then
Response.Write "&nbsp;&gt;&gt;&nbsp;<a href='Bs_Download.asp?Bs_BigClassName=" & rs("Bs_BigClassName")&"&Bs_SmallClassName=" & rs("Bs_SmallClassName") & "'>" & rs("Bs_SmallClassName") & "</a>"
end if
'response.write "&nbsp;&gt;&gt;&nbsp;"& rs("Bs_Title")
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
				<table cellSpacing=1 cellpadding=0 align="center" class='MC_DoShow_Title1'>
					<TR class='MC_DoShow_T1Tr1'>
						<td colspan=2>&nbsp;<%=rs("Bs_Title")%></TD>
					</TR>
					<TR class='MC_DoShow_T1Tr2'>
						<td class='MC_DoShow_T1r2Td1'>&nbsp;������ԣ�<%=rs("Bs_Language")%></TD>
						<td rowspan=6><DIV align=center><IMG src='<%=rs("Bs_PhotoUrl")%>'height=150 width=150></DIV></TD>
					</TR>
					<TR><td class='MC_DoShow_T1rnTd1'>&nbsp;��Ȩ���ͣ�<%=rs("Bs_LicenceType")%></TD></TR>
					<TR><td class='MC_DoShow_T1rnTd1'>&nbsp;�ļ���С��<%=rs("Bs_FileSize")%></TD></TR>
					<TR><td class='MC_DoShow_T1rnTd1'>&nbsp;���л�����<%=rs("Bs_System")%></TD></TR>
					<TR><td class='MC_DoShow_T1rnTd1'>&nbsp;�������ڣ�<%=FormatDateTime(rs("Bs_UpDateTime"),2)%></TD></TR>
					<TR><td class='MC_DoShow_T1rnTd1'>&nbsp;���ش�����<%=rs("Bs_Hits")%></TD></TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<table cellSpacing=0 cellpadding=0 align="center" class='MC_DoShow_Title2'>
					<TR class='MC_DoShow_t2Tr1'><td>::���ص�ַ::</TD></TR>
					<TR>
						<td class='MC_DoShow_t2r2Td1'>
��<a href="<%=rs("Bs_DownloadUrl")%>" target=_blank>���ص�ַ1</a>��
<%
if rs("Bs_DownloadUrl2")<>"" then
	response.Write "<a href='" & rs("Bs_DownloadUrl2") & "'target=_blank>���ص�ַ2</a>��"
	else
	response.Write ""
end if
%>
						</TD>
					</TR>
					<TR class='MC_DoShow_t2Tr3'><td>::������::</TD></TR>
					<TR>
						<td class='MC_DoShow_t2r4Td1'>
<DIV>&nbsp;&nbsp;<%=Bs_Content%></DIV>
						</TD>
					</TR>
					<TR class='MC_DoShow_t2Tr5'><td>::����˵��::</TD></TR>
					<TR>
						<td class='MC_DoShow_t2r6Td1'><DIV>
						* Ϊ�˴ﵽ���������ٶȣ��Ƽ�ʹ��<a href="http://www.amazesoft.com/" target=_blank><FONT COLOR="#603811">���ʿ쳵</FONT></a>���ر�վ�����<BR> 
						* ��������ָ�����������أ���֪ͨ<a href="mailto:<%=BsEmail%>"><FONT COLOR="#603811">����Ա</FONT></a>��лл��<BR> 
						* δ����վ��ȷ��ɣ��κ���վ���÷Ƿ���������Ϯ��վ��Դ��������ҳ�棬��ע�����Ա�վ��лл����֧�֣�</DIV> 
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<%
end if
rs.close
set rs=nothing
%>
<!-- ���ز鿴 END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_bottom_title'>
					<tr> 
						<td><div align="right">
��<a href='javascript:history.back()'>����</a
>����<A href="javascript:window.scroll(0,-360)">����</A>��
&nbsp;</div></TD>
					</tr>
					<tr><td class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	</TABLE>
</TD>
<%if strSkins <= 200 and strSkins >= 1 then%>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<%
Call Style_Download() '
Call Style_Fellow() '��������
%>
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
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD><TD vAlign=top class='Mr_TitlePt'> 
<% Call Right_Download() %>
</TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>