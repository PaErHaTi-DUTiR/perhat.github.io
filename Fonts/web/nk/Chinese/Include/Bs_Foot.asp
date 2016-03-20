<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
</TD></TR></TABLE>
</DIV>

<!--#include file="Bs_FootAddress.asp" -->
<SCRIPT LANGUAGE="JavaScript">FootLogoC1()</SCRIPT>

<%if strSkins <= 100 and strSkins >= 1 then%>
<%Call BsFootAddress1()%>
<SCRIPT LANGUAGE="JavaScript">CnFootMenu()</SCRIPT>

<%elseif strSkins <= 150 and strSkins >= 101 then%>
<SCRIPT LANGUAGE="JavaScript">CnFootMenu()</SCRIPT>
<%Call BsFootAddress1()%>

<%elseif strSkins <= 200 and strSkins >= 151 then%>
<%Call BsFootAddress1()%>
<SCRIPT LANGUAGE="JavaScript">CnFootMenu()</SCRIPT>
<%end if%>

<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='FootCopyright'>
	<TR><TD class='bottom_TdHR'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<%
Style_Copy = replace(replace(replace(Style_Copy,"$BsCompanyName",BsCompanyName),"$BsWeb",BsWeb),"$ExecutionTime",fix((timer()-StarTime)*1000))
Response.Write Style_Copy
%>
</TABLE>
</DIV>
<%
call CloseConn()
%>
</body>
</HEAD><HEAD> 
<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE"> <!-- 禁止浏览器缓存页面 -->
</HEAD> 
</html>
<!--#include file="../../QQ/QQShow.asp"-->