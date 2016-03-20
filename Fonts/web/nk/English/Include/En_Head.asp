<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<HTML><HEAD><TITLE>∷<%=BsEnCompanyName%>∷</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta name="Copyright" content="www.web300.cn">
<meta name="Keywords" content=" 公司(企业)网站管理系统">
<meta name="Description" content=" 公司(企业)网站管理系统">
<META http-equiv="Pragma" content="no-cache">
<META http-equiv="Cache-Control" content="no-cache, must-revalidate">
<META http-equiv="Expires" content="0">
<link href="../Skin/<%=Request.Cookies("BsSkins")%>/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
TD {
	FONT-SIZE: 12px; LINE-HEIGHT: 15px; FONT-FAMILY: "Arial,TFixedsys";
}
.menuitems {
	PADDING-RIGHT: 1px;
	PADDING-TOP: 1px;
	PADDING-LEFT: 1px;
	PADDING-BOTTOM: 1px;
	MARGIN: 2px;
	font-size:9pt;
	line-height:14pt;
}
.menuskin {
	BORDER-RIGHT: #0A2999 1px solid ;
	BORDER-TOP: #0A2999 1px solid;
	BORDER-LEFT: #0A2999 1px solid;
	BORDER-BOTTOM: #0A2999 1px solid;
	BACKGROUND-IMAGE: url(../Img/menubg.gif);
	POSITION: absolute;
	VISIBILITY: hidden;
}
#mouseoverstyle {
	PADDING-RIGHT: 0px;
	PADDING-LEFT: 0px;
	PADDING-BOTTOM: 0px;
	PADDING-TOP: 0px;
	BORDER-RIGHT: #0B008B 1px solid;
	BORDER-TOP: #0B008B 1px solid;
	BORDER-LEFT: #0B008B 1px solid;
	BORDER-BOTTOM: #0B008B 1px solid;
	BACKGROUND-COLOR: #FFEEC2 
}
.menuskin A {PADDING-RIGHT:10px;PADDING-LEFT:30px;}
-->
</style>
<SCRIPT LANGUAGE="JavaScript" src="../Js/BsTopLogo1.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="../js/BsTopMenu.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="../js/BsFootMenu.js"></SCRIPT>
</HEAD>

<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<div id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<!--#include file="En_TopLogin.asp" -->
<%
Dim strSkins
strSkins = CInt(Request.Cookies("BsSkins"))
%>
<%if strSkins <= 100 and strSkins >= 1 then%>

<%Call BsTopLogin1()%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA1()</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">EnTopMenu(<%=strSkins%>)</SCRIPT>

<%elseif strSkins <= 150 and strSkins >= 101 then%>
<%Call BsTopLogin2()%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA2()</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">EnTopMenu(<%=strSkins%>)</SCRIPT>

<%elseif strSkins <= 160 and strSkins >= 151 then%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA3()</SCRIPT>
<%Call BsTopLogin3()%>
<SCRIPT LANGUAGE="JavaScript" src="../js/EnTopMenu2.js"></SCRIPT>

<%elseif strSkins <= 200 and strSkins >= 161 then%>
<%Call BsTopLogin3()%>
<SCRIPT LANGUAGE="JavaScript" src="../js/EnTopMenu2.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">TopLogoA1()</SCRIPT>

<%elseif strSkins <= 250 and strSkins >= 201 then%>
<%Call BsTopLogin4()%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA4()</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="../js/EnTopMenu4.js"></SCRIPT>

<%end if%>

<SCRIPT LANGUAGE="JavaScript">TopLogoB1()</SCRIPT>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='MiddleTitle'>
<TR><TD vAlign=top> 
