<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/En_System.asp"-->
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
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<%end if%><TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 会员中心首页 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align='center' class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
Member centre
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
					<tr> 
						<TD> <%If Session("UserName")="" Then %> <br> <br>
							<table width="100%" border="0" cellpadding="0" cellspacing="3">
								<tr> 
									<td width="19%" align="right"></td>
									<td width="81%"> <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699"
									>sorry, you have no login, it is unable to enter the member centre, <BR>please log in first ！</font></p>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#006699"
									>If you are not still our member, invite immediately <a href="En_UserReg.asp"><font color="#006699">Register</font></a>。</font></p></td>
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
									<td width="81%"><font color="#990000"><b><font color="#006699">There is the following function in member centre：</font></b></font></td>
								</tr>
								<tr> 
									<td width="19%" align="right"><font color="#666666">1、</font></td>
									<td width="81%"><font color="#666666">Revise your member materials for registration </font></td>
								</tr>
								<tr> 
									<td width="19%" align="right" height="14"><font color="#666666">2、</font></td>
									<td width="81%" height="14"> <p><font color="#666666">Inquire about your order, and the treatment situation of the order </font></p></td>
								</tr>
								<tr> 
									<td width="19%" align="right"><font color="#666666">3、<BR>&nbsp;</font></td>
									<td width="81%"><font color="#666666">Special-purpose private visitor book, you can stay and make <BR>peace to look over our reply for us here</font></td>
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
									<div align="center"><center>
									</center></div>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> 
									<b><font color="#333333">Homepage of member centre：</font></b> <br><br><br>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><%=Session("UserName")%> Howdy：<br>	<br></p>
									<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
									You have already entered member service centre now, here can be visited<BR><BR>
									unless registering members. You can revise your registered information here ,<BR><BR>
									Give us leave a message , look over between we and you message reply ,<BR><BR>
									can inquire your order and order punish the situation too. </p>
									</TD>
								</TR>
								<TR vAlign=top > 
									<TD  width="100%" height="15"> &nbsp;&nbsp;&nbsp; <p>&nbsp;</p><p>　</p></TD>
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
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot4.asp" -->
<%end if%>