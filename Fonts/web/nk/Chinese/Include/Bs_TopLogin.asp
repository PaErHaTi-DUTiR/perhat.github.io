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
							 <TD><a href="http://<%=BsWebMail%>" target="_blank">��˾�ʾ�</a></TD>
							 <TD align="center">��</TD>
							 <TD><a href="Bs_CoProfile.asp?Action=Contact">��ϵ����</a></TD>
							 <TD align="center">��</TD>
							 <TD><a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame">�ղر�վ</a></TD>
							 <TD align="center">��</TD>
							 <TD><a href="../English/index.asp">English</a></TD>
							 <TD>��<!--#include file="gb2big5.htm"--></TD>
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
				 <TD class='TopLogin2B_Td1' align="center"> ��</TD>
				 <TD><a href="http://<%=BsWebMail%>" target="_blank" class='Menu1'>��˾�ʾ�</a></TD>
				 <TD class='TopLogin2B_Td1' align="center"> ��</TD>
				 <TD><a href="Bs_CoProfile.asp?Action=Contact" class='Menu1'>��ϵ����</a></TD>
				 <TD class='TopLogin2B_Td1' align="center"> ��</TD>
				 <TD><a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame" class='Menu1'>�ղر�վ</a></TD>
				 <TD class='TopLogin2B_Td1' align="center"> ��</TD>
				 <TD><a href="../English/index.asp" class='Menu1'>English</a></TD>
				 <TD>&nbsp;��<!--#include file="gb2big5.htm"--></TD>
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
     "������",
     "����һ",
     "���ڶ�",
     "������",
     "������",
     "������",
     "������");
document.write ("",today.getYear(),"��",today.getMonth()+1,"��",today.getDate(),"��&nbsp;",d[today.getDay()+1],"" ); 
</script>
		</TD>
		<TD class=TopLogin3_TD2>��</TD>
		<TD class=TopLogin3_TD3 align=right>
<IMG src='../Skin/<%=strSkins%>/playicqicon06.gif' width=11 height=10 border=0>
<a href="http://<%=BsWebMail%>" target="_blank" class='Menu1'>��˾�ʾ�</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon03.gif' width=11 height=10 border=0>
<a href="Bs_CoProfile.asp?Action=Contact" class='Menu1'>��ϵ����</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon04.gif' width=11 height=10 border=0>
<a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame" class='Menu1'>�ղر�վ</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon01.gif' width=11 height=10 border=0>
<a href="../English/index.asp" class='Menu1'>English</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon01.gif' width=11 height=10 border=0>
<!--#include file="gb2big5.htm"--></TD>
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
     "������",
     "����һ",
     "���ڶ�",
     "������",
     "������",
     "������",
     "������");
document.write ("",today.getYear(),"��",today.getMonth()+1,"��",today.getDate(),"��&nbsp;",d[today.getDay()+1],"" ); 
</script>
		</TD>
		<TD class=TopLogin4_TD2>��</TD>
	  <TD class=TopLogin4_TD3 align=center><IMG src='../Skin/<%=strSkins%>/playicqicon03.gif' width=11 height=10 border=0>
<a href="Bs_CoProfile.asp?Action=Contact" class='Menu1'>��ϵ����</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon04.gif' width=11 height=10 border=0>
<a onclick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsCompanyName%>');return false;" href="http://<%=BsWeb%>/" target="mainFrame" class='Menu1'>�ղر�վ</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon01.gif' width=11 height=10 border=0>
<a href="../English/index.asp" class='Menu1'>English</a>
<IMG src='../Skin/<%=strSkins%>/playicqicon01.gif' width=11 height=10 border=0>
<!--#include file="gb2big5.htm"-->
<img src='../Skin/<%=strSkins%>/playicqicon01.gif' width="11" height="10" border="0" /> <a href="../uy/uy/index.asp" class='Menu1'>ά��</a></TD>
	</TR>
</TABLE>
</DIV>
<%
End Sub
%>