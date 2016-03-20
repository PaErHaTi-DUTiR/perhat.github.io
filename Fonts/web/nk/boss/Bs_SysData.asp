<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<script type="text/JavaScript">
<!--
function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<!--#include file="bsconfig.asp"-->
<!--#include file="Include/Bs_SysInfo.asp"-->

<%
'=========================================================
'
'产品名称：公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有: www.web300.cn 
'程序制作：www.web300.cn开发团队
'Copyright 2003-2006 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>

<!-- #include file="Inc/Head.asp" -->
<body onmouseover="MM_displayStatusMsg('<%=rs("BsCompanyName")%>');return document.MM_returnValue">
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>网 站 基 本 资 料 设 置</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
<%
Str_SysData = replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(Str_SysData,"$S01",rs("Code")),"$S02",rs("BsCompanyName")),"$S03",rs("BsEnCompanyName")),"$S04",rs("BsAddress")),"$S05",rs("BsEnAddress")),"$S06",rs("BsLinkman")),"$S07",rs("BsEnLinkman")),"$S08",rs("BsPhone")),"$S09",rs("BsFax")),"$S10",rs("BsWeb")),"$S11",rs("BsEmail")),"$S12",rs("BsWebMail")),"$S13",rs("BsZipcode")),"$S14",rs("BsSkins_Default"))
Response.Write Str_SysData
%>
      <br>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<script language=Javascript>

	function checkchar(str)
	{
		str=str.toLowerCase()
		if (str.search("<"+"%")>0)  
		{
			window.alert("("+str+")中有非法字符,请检查!")
			return false;
		}
		if (str.search("<scrip"+"t")>0)
		{
			window.alert("("+str+")中有非法字符,请检查!")
			return false;
		}
		return true;
	}
</script>
</body>
</html>