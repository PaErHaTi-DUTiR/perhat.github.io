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
					<TR> 
						<TD align="center"><FORM name='UserReg'action='Bs_UserRegPost.asp'method='post'>
							<table width=90% border=0 align="center" cellpadding=5 cellspacing=1 bordercolor="#FFFFFF" style="border-collapse: collapse">
								<TR align=center> 
									<TD height=20 colSpan=2><b>新用户注册</b></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><b>用户名：</b><BR>不能超过14个字符（7个汉字）</TD>
									<TD width="63%"> <INPUT maxLength=14 size=30 name=UserName> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><B>密码(至少6位)：</B><BR>请输入密码，区分大小写。 不要使用类似 '*'、''的特殊字符</TD>
									<TD width="63%"> <INPUT type=password maxLength=12 size=30 name=Password> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>确认密码(至少6位)：</strong><BR>请再输一遍确认</TD>
									<TD width="63%"> <INPUT type=password maxLength=12 size=30 name=PwdConfirm> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>密码问题：</strong><BR>忘记密码的提示问题</TD>
									<TD width="63%"> <INPUT   type=text maxLength=50 size=30 name="Question"> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>问题答案：</strong><BR>忘记密码的提示问题答案，用于取回密码</TD>
									<TD width="63%"> <INPUT type=text maxLength=20 size=30 name="Answer"> <font color="#FF0000">*</font> </TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>性别：</strong><BR>请选择您的性别</TD>
									<TD width="63%"> <INPUT type=radio CHECKED value="1" name=sex>男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex>女</TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>Email地址：</strong><BR>请输入有效的邮件地址</TD>
									<TD width="63%"> <INPUT maxLength=50 size=30 name=Email> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>公司网址：</strong></TD>
									<TD width="63%"><INPUT name=homepage id="homepage" value="http://" size=30   maxLength=50></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>公司名称：</strong><BR>您的公司名称</TD>
									<TD width="63%"> <INPUT name=Comane id="Comane" size=30   maxLength=100></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><p><strong>收货地址：<br></strong>您的公司地址或收货地址</p></TD>
									<TD width="63%"> <input name=Add id="Add" size=30 maxlength=20> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>收货人：<br></strong>收货联系人<strong> </strong></TD>
									<TD width="63%"><input name=Somane id="Somane" size=30 maxlength=20></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>邮政编码：</strong></TD>
									<TD width="63%"><input name=Zip id="Zip" size=30 maxlength=20> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD><strong>联系电话：<br></strong>格式+86-010-87654321<strong> </strong></TD>
									<TD width="63%"><input name=Phone id="Phone" value="+86" size=30 maxlength=20> <font color="#FF0000">*</font></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="37%"><strong>传 真：</strong></TD>
									<TD width="63%"> <INPUT name=Fox id="Fox" size=30 maxLength=50></TD>
								</TR>
							</TABLE>
							<div align="center"><INPUT type=submit value=" 注 册 " name=Submit>&nbsp; <INPUT name=Reset type=reset id="Reset" value=" 清 除 "></div>
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
<!--#include file="Include/Bs_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>