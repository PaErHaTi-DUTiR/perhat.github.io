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
Sub BsFootAddress1()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='FootAddress'>
	<TR> 
		<TD vAlign=top>
			<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
				<TR>
					<TD class='FootAddress_Td1'>
						<TABLE cellSpacing=0 cellPadding=0 class='FootAddressB'>
							<TR> 
								<TD height=25 style="padding-top:6px"><DIV align=center>Address：<%=BsEnAddress%>  Tel：<%=BsPhone%>  Email:<a href="mailto:<%=BsEmail%>"><%=BsEmail%></a></DIV></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR><TD class='FootAddress_Td2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
				<TR><TD class='FootAddress_Td3'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</DIV>
<%
End Sub
Sub BsFootAddress2()
%>
<DIV align="center">
</DIV>
<%
End Sub
%>
