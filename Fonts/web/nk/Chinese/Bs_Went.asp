<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
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
<%
dim id,rst,sql
Id=Session("UserName")
set rst = Server.CreateObject("ADODB.recordset")
sql="select * from Bs_User where UserName='"&Id&"'"
rst.open sql,conn,1,1
%>
<!--#include file="Include/Bs_Head.asp" -->
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
<%end if%><TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 信息反馈 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 align="center"  class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
信息反馈
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
					<TR vAlign=top > 
						<TD  width="80%" height="18"> <form method="post" action="Bs_NetBookSave.asp">
							<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
								<tr> 
									<td width="100%"> <div align="center"> 
										<table width="100%" height="79" border="0" align="center" cellpadding="0" cellspacing="3">
											<tr> 
												<td width="23%" height="25" align="right">用户名称：</td>
												<td width="77%" height="25"><font> 
												<input type="text" readonly name="Name" maxlength="36" value="<%if Session("UserName")="" then response.write"未注册用户" end if%><%=Session("UserName")%>" style="background-color: #FFFFFF; border-style: solid; border-color: #FFFFFF">
												</font></td>
											</tr>
											<tr> 
												<td width="23%" height="25" align="right">公司名称：</td>
												<td width="77%" height="-6"><font> 
												<input type="text" name="Comane" size="30" maxlength="36" value="<%=rst("Comane")%>" style="font-size: 14px">
												</font></td>
											</tr>
											<tr> 
												<td height="25" align="right">联系人：</td>
												<td width="77%" height="-2"><font> 
												<input type="text" name="Somane" size="12" maxlength="30" value="<%=rst("Somane")%>" style="font-size: 14px">
												</font></td>
											</tr>
											<tr> 
												<td height="25" align="right">联系电话：</td>
												<td width="77%" height="-1"><font> 
												<input type="text" name="Phone" size="24" maxlength="36" value="<%=rst("Phone")%>" style="font-size: 14px">
												</font></td>
											</tr>
											<tr> 
												<td height="25" align="right">联系传真：</td>
												<td width="77%" height="11"><font> 
												<input type="text" name="Fox" size="18" maxlength="36" value="<%=rst("Fox")%>" style="font-size: 14px">
												</font></td>
											</tr>
											<tr> 
												<td height="25" align="right">E-mail：</td>
												<td width="77%" height="11"><font> 
												<input type="text" name="Email" size="18" maxlength="36" value="<%=rst("Email")%>" style="font-size: 14px">
												</font></td>
											</tr>
											<tr> 
												<td width="23%" height="25" align="right">反馈主题：</td>
												<td width="77%" height="25"><input type="text" name="Title" size="42" maxlength="36" style="font-size: 14px"></td>
											</tr>
											<tr> 
												<td width="23%" height="1" valign="top" align="right">反馈内容：</td>
												<td width="77%" height="1" valign="top"><textarea rows="10" name="Content" cols="45" style="font-size: 14px"></textarea></td>
											</tr>
											<tr> 
												<td width="23%" height="0" valign="top"> </td>
												<td width="77%" height="0" valign="top"> 
												<input type="submit" value="提交留言" 
												name="cmdOk"> <input type="reset" value="重写" name="cmdReset"> </td>
											</tr>
										</table>
									</div></td>
								</tr>
							</table>
						</form></TD>
					</TR>
					<TR vAlign=top > 
						<TD  width="100%" height="15"> </TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<%
rst.close
set rst=nothing
%>
<!-- 信息反馈 -->
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
<!--#include file="Include/Bs_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-
mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>