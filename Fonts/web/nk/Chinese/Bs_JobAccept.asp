<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
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
dim job
job=trim(request("job"))
%>

<!--#include file="Include/Bs_Head.asp" -->
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR> 
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Job() %>
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

<!-- �˲�ӦƸ -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
�˲�ӦƸ�ǼǱ�
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
				<TABLE cellSpacing=0 cellPadding=0 class='MC_HSJNr_Title'>
					<TR> 
						<TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					</TR>
					<TR vAlign=top > 
						<TD  width="80%" height="18"> <form method="post" action="Bs_JobAcceptSave.asp">
							<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
								<tr> 
									<td width="100%"> <div align="center"> 
										<table width="100%" height="79" border="0" align="center" cellpadding="0" cellspacing="3">
											<tr> 
												<td width="23%" height="25" align="right">ӦƸ��λ��</td>
												<td width="77%" height="25"
><b><%=job%></b> <input name="jobname" type="hidden" value="<%=job%>"></td>
											</tr>
											<tr> 
												<td width="23%" height="25" align="right">�� ����</td>
												<td width="77%" height="-6"><font
><input name="mane" type="text" class="SmallInput" id="mane" size="12" maxlength="16"
></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">�� ��</td>
												<td width="77%" height="-2"><font
><select name="sex"><option value="��" selected>��</option
><option value="Ů">Ů</option></select></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">�������ڣ�</td>
												<td width="77%" height="-1"
><input name="birthday" type="text" class="SmallInput" id="birthday" size="14" maxlength="30">&nbsp;��ʽ��1976-02-02 </td>
											</tr>
											<tr> 
												<td height="25" align="right">����״����</td>
												<td width="77%" height="11"><font
><select name="marry"><option value="δ��" selected>δ��</option
><option value="�ѻ�">�ѻ�</option></select></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">��ҵԺУ��</td>
												<td width="77%" height="11"><font
><input name="school" type="text" class="SmallInput" id="school" size="30" maxlength="50"></font></td>
											</tr>
											<tr> 
												<td width="23%" height="25" align="right">ѧ ����</td>
												<td width="77%" height="-6"
><input name="studydegree" type="text" class="SmallInput" id="studydegree" size="10" maxlength="16"></td>
											</tr>
											<tr> 
												<td height="25" align="right">ר ҵ��</td>
												<td width="77%" height="-3"><font
><input name="specialty" type="text" class="SmallInput" id="specialty" size="10" maxlength="30"></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">��ҵʱ�䣺</td>
												<td width="77%" height="-2"
><input name="gradyear" type="text" class="SmallInput" id="gradyear" size="14" maxlength="30">&nbsp;��ʽ��1998-7��</td>
											</tr>
											<tr> 
												<td height="25" align="right">�� ����</td>
												<td width="77%" height="-1"><font
><input name="telephone" type="text" class="SmallInput" id="telephone" size="16" maxlength="30"></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">E-mail��</td>
												<td width="77%" height="-6"><font
><input name="email" type="text" class="SmallInput" id="email" size="20" maxlength="30"></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">��ϵ��ַ��</td>
												<td width="77%" height="-3"><font
><input name="address" type="text" class="SmallInput" id="address" size="30" maxlength="50"></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">ˮƽ��������</td>
												<td width="77%" height="-2"><font
><textarea name="ability" cols="35" rows="5" class="SmallInput" id="ability"></textarea></font></td>
											</tr>
											<tr> 
												<td height="25" align="right">���˼�����</td>
												<td width="77%" height="-1"><font
><textarea name="resumes" cols="35" rows="5" class="SmallInput" id="resumes"></textarea></font></td>
											</tr>
											<tr> 
												<td width="23%" height="0" valign="top"></td>
												<td width="77%" height="0" valign="top"
><input type="submit" value="�ύ����" name="cmdOk"> <input type="reset" value="��д" name="cmdReset"></td>
											</tr>
										</table>
									</div></td>
								</tr>
							</table>
						</form></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<!-- �˲�ӦƸ END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0 class='MC_bottom_title'>
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