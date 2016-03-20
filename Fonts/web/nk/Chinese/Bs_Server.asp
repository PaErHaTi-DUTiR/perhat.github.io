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
<%end if%>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 会员中心首页 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
会员中心
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
					<tr> 
						<TD> <%If Session("UserName")="" Then %> <br> <br>
							<table width="100%" border="0" cellpadding="0" cellspacing="3">
								<tr> 
									<td width="19%" align="right"></td>
									<td width="81%"> <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699"
									>对不起，您还没有登录，无法进入会员中心，请先登录！</font></p>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699"
									>如果您还不是我们的会员，请立即<a href="Bs_UserReg.asp"><font color="#006699">注册</font></a>。</font></p></td>
								</tr>
								<tr> 
									<td width="19%" align="right"></td>
									<td width="81%"> <div align="center"> 
										<table class="main" cellSpacing="0" cellPadding="2" width="100%" border="0" height="1">
											<tbody>
												<tr> 
													<td width="100%" height="1"><% call ShowUserLogina() %></td>
												</tr>
											</tbody>
										</table>
									</div></td>
								</tr>
								<tr> 
									<td width="19%" align="right"></td>
									<td width="81%"><font color="#990000"><b><font color="#006699">会员中心有以下功能：</font></b></font></td>
								</tr>
								<tr> 
									<td width="19%" align="right"><font color="#666666">1、</font></td>
									<td width="81%"><font color="#666666">修改您的会员注册资料</font></td>
								</tr>
								<tr> 
									<td width="19%" align="right" height="14"><font color="#666666">2、</font></td>
									<td width="81%" height="14"> <p><font color="#666666">查询您的订单，以及订单的处理情况</font></p></td>
								</tr>
								<tr> 
									<td width="19%" align="right"><font color="#666666">3、</font></td>
									<td width="81%"><font color="#666666">专用私人留言簿，您可在此给我们留言和查看我们的回复</font></td>
								</tr>
								<tr> 
									<td height="25" align="right">&nbsp;</td>
									<td>&nbsp;</td>
								</tr>
							</table>
<%else%>
							<TABLE width="85%" height="136" align="center" cellPadding=0 cellSpacing=0>
								<TR vAlign=top > 
									<TD  width="100%" height="18">
									<p><br><br><br></p>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> 
									<b><font color="#333333">会员中心首页：</font></b> <br><br><br>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><%=Session("UserName")%> 您好：<br>	<br></p>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">您现在已经进入会员服务中心，这里只有注册会员才能访问。您可在这里修改您的注册<BR>
									<BR>信息、给我们留言、查看我们对您留言的答复，也可以查询您的订单及订单处理情况。</p>
									</TD>
								</TR>
								<TR vAlign=top > 
									<TD  width="100%" height="15"> &nbsp;&nbsp;&nbsp; <p>&nbsp;</p></TD>
								</TR>
							</TABLE>
<%end if%> 
						</TD>
					</TR>
				</table>
			</TD>
		</TR>
<!-- 会员中心 END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0 align="center"  class='MC_bottom_title'>
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