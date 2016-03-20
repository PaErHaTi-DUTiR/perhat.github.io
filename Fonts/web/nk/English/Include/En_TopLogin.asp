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
Sub BsTopLogin1()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='TopLogin'>
	<TR> 
		<TD class='TopLogin_Td'>
			<TABLE cellPadding=0 cellSpacing=0 align='center' class='TopLogin'>
				<TR> 
					<TD class='TopLogin_Td1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					<TD class='TopLogin_Td2'>
						<TABLE border=0 align="left" cellPadding=0 cellSpacing=0>
							<TR> 
							 <TD align="center">
									<table cellpadding="0" cellspacing="0" align="center" class='TopLogin1'>
										<tr> 
											<td height=25><% call ShowUserLoginT() %></td>
										</tr>
									</table>
							 </TD>
							<TR> 
						</TABLE>
					</TD>
					<TD class='TopLogin_Td3'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					<TD class='TopLogin_Td4'>
						<TABLE cellPadding=0 cellSpacing=0 border=0 align="right">
							<TR> 
							 <TD height=10 colSpan=9> </TD> 
							<TR> 
							<TR> 
							 <TD> </TD>
							 <TD><a href="http://<%=BsWebMail%>" target="_blank">WebMail</a></TD>
							 <TD width=20 align="center">│</TD>
							 <TD><a href="En_CoProfile.asp?Action=Contact">Contact us</a></TD>
							 <TD width=20 align="center">│</TD>
							 <TD><a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsEnCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame">Favorite</a></TD>
							 <TD width=20 align="center">│</TD>
							 <TD><a href="../Chinese/index.asp">Chinese</a></TD>
							 <TD width=5> </TD>
							<TR> 
						</TABLE>
					</TD>
					<TD class='TopLogin_Td5'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</DIV>
<%
End Sub
Sub BsTopLogin2()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='TopLogin2'>
	<TR>
		<TD class='TopLogin2_Td1' align="left">&nbsp;&nbsp;&nbsp;Welcome to <%=BsWeb%></TD>
		<TD class='TopLogin2_Td2'></TD>
		<TD class='TopLogin2_Td3' align="right">
			<TABLE cellPadding=0 cellSpacing=0 align="right" class='TopLogin2B'>
				<TR> 
				 <TD class='TopLogin2B_Td1' align="center"> ⊙</TD>
				 <TD><a href="http://<%=BsWebMail%>" target="_blank" class='Menu1'>WebMail</a></TD>
				 <TD class='TopLogin2B_Td1' align="center"> ⊙</TD>
				 <TD><a href="en_CoProfile.asp?Action=Contact" class='Menu1'>Contact us</a></TD>
				 <TD class='TopLogin2B_Td1' align="center"> ⊙</TD>
				 <TD><a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsEnCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame" class='Menu1'>Favorite</a></TD>
				 <TD class='TopLogin2B_Td1' align="center"> ⊙</TD>
				 <TD><a href="../Chinese/index.asp" class='Menu1'>Chinese</a></TD>
				 <TD width=10> </TD>
				<TR> 
			</TABLE>
		</TD>
	</TR>
</TABLE>
</DIV>
<%
End Sub
Sub BsTopLogin3()
%>
<DIV align="center">
<TABLE cellSpacing=0 cellPadding=0 width=760 class='TopLogin3'>
	<TR>
		<TD class=TopLogin3_TD1>
<script LANGUAGE="JavaScript">
 today=new Date();
 function initArray(){
   this.length=initArray.arguments.length
   for(var i=0;i<this.length;i++)
   this[i+1]=initArray.arguments[i]  }
   var d=new initArray(
     "SUN",
     "MON",
     "TUE",
     "WED",
     "THU",
     "FRI",
     "SAT");
document.write ("",today.getMonth()+1,"-",today.getDate(),"-",today.getYear(),"&nbsp;",d[today.getDay()+1],"" ); 
</script>
		</TD>
		<TD class=TopLogin3_TD2>　</TD>
		<TD class=TopLogin3_TD3 align=right>
<IMG src='../Skin/<%=strSkins%>/playicqicon06.gif' width=11 height=10 border=0>
<a href="http://<%=BsWebMail%>" target="_blank" class='Menu1'>WebMail</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon03.gif' width=11 height=10 border=0>
<a href="en_CoProfile.asp?Action=Contact" class='Menu1'>Contact us</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon04.gif' width=11 height=10 border=0>
<a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsEnCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame" class='Menu1'>Favorite</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon01.gif' width=11 height=10 border=0>
<a href="../Chinese/index.asp" class='Menu1'>Chinese</a>
		&nbsp;</TD>
	</TR>
</TABLE>
</DIV>
<%
End Sub
Sub BsTopLogin4()
%>
<DIV align="center">
<TABLE width=1003 border="0" cellPadding=0 cellSpacing=0 class='TopLogin4'>
	<TR>
		<TD class=TopLogin4_TD1>
<script LANGUAGE="JavaScript">
 today=new Date();
 function initArray(){
   this.length=initArray.arguments.length
   for(var i=0;i<this.length;i++)
   this[i+1]=initArray.arguments[i]  }
   var d=new initArray(
     "SUN",
     "MON",
     "TUE",
     "WED",
     "THU",
     "FRI",
     "SAT");
document.write ("",today.getMonth()+1,"-",today.getDate(),"-",today.getYear(),"&nbsp;",d[today.getDay()+1],"" ); 
</script>
		</TD>
		<TD class=TopLogin4_TD2>　</TD>
		<TD class=TopLogin4_TD3 align=right>
<IMG src='../Skin/<%=strSkins%>/playicqicon06.gif' width=13 height=9 border=0>
<a href="http://<%=BsWebMail%>" target="_blank" class='Menu1'>WebMail</a>&nbsp;
<IMG src='../Skin/<%=strSkins%>/playicqicon03.gif' width=11 height=10 border=0>
<a href="en_CoProfile.asp?Action=Contact" class='Menu1'>Contact us</a>&nbsp;
<IMG src='../Skin/<%=strSkins%>/playicqicon04.gif' width=11 height=10 border=0>
<a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsEnCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame" class='Menu1'>Favorite</a>&nbsp;
<IMG src='../Skin/<%=strSkins%>/playicqicon01.gif' width=11 height=10 border=0>
<a href="../Chinese/index.asp" class='Menu1'>Chinese</a>
<img src='../Skin/<%=strSkins%>/playicqicon01.gif' width="11" height="10" border="0" /> <a href="../uy/uy/index.asp" class='Menu1'>维文</a>&nbsp;&nbsp;&nbsp;</TD>
	</TR>
</TABLE>
</DIV>
<%
End Sub
%>