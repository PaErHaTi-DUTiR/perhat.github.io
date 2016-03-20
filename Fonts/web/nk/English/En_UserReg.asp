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

<!-- 用户注册 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_Pt_Title'>
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
				<TABLE cellSpacing=0 cellPadding=0 align="center"  class='MC_ServeNr_Title'>
					<TR> 
						<TD align="center"><FORM name='UserReg'action='En_UserRegPost.asp'method='post'>
							<table width=90% border=0 align="center" cellpadding=5 cellspacing=1 bordercolor="#FFFFFF" style="border-collapse: collapse">
								<TR align=center> 
									<TD height=20 colSpan=2><b>New user register</b></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><b>UserName：</b><BR>Can exceed 14 characters</TD>
									<TD width="63%"> <INPUT maxLength=14 size=30 name=UserName> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><B>Password (least 6 bit)：</B><BR>input pass，Distinguish big and small.</TD>
									<TD width="63%"> <INPUT type=password maxLength=12 size=30 name=Password> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>AffirmPass (least 6 bit)：</strong><BR>Please fail again and confirm</TD>
									<TD width="63%"> <INPUT type=password maxLength=12 size=30 name=PwdConfirm> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>PassQuestion：</strong><BR>Forget the suggestion question of the password</TD>
									<TD width="63%"> <INPUT   type=text maxLength=50 size=30 name="Question"> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>QuestionAnswer：</strong><BR>Forget the suggestion question answer to the password.</TD>
									<TD width="63%"> <INPUT type=text maxLength=20 size=30 name="Answer"> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>Sex ：</strong><BR>Please choose your sex</TD>
									<TD width="63%"> <INPUT type=radio CHECKED value="1" name=sex>Male &nbsp;&nbsp; <INPUT type=radio value="0" name=sex>Female</TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>Email：</strong><BR>Please input the effective mail address</TD>
									<TD width="63%"> <INPUT maxLength=50 size=30 name=Email> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>CompanyWebsite ：</strong></TD>
									<TD width="63%"><INPUT name=homepage id="homepage" value="http://" size=30   maxLength=50></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>CompanyName ：</strong><BR>Your company name</TD>
									<TD width="63%"> <INPUT name=Comane id="Comane" size=30   maxLength=100></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><p><strong>ReceiveAddress ：<br></strong>Your company address or the address of receiving</p></TD>
									<TD width="63%"> <input name=Add id="Add" size=30 maxlength=20> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>Consignee ：<br></strong>Receive the person to contact<strong> </strong></TD>
									<TD width="63%"><input name=Somane id="Somane" size=30 maxlength=20></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>Zipcode ：</strong></TD>
									<TD width="63%"><input name=Zip id="Zip" size=30 maxlength=20> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>Telephone ：<br></strong>Form +86-0755-888888<strong> </strong></TD>
									<TD width="63%"><input name=Phone id="Phone" value="+86" size=30 maxlength=20> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>Fax ：</strong></TD>
									<TD width="63%"> <INPUT name=Fox id="Fox" size=30 maxLength=50></TD>
								</TR>
							</TABLE>
							<div align="center"><INPUT type=submit value=" Registe " name=Submit>&nbsp; <INPUT name=Reset type=reset id="Reset" value=" Remove "></div>
						</form></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<!-- 用户注册 END -->
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