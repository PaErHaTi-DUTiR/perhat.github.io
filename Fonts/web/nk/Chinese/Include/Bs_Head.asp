<% Response.CodePage=936%>
<% Response.Charset="gb2312" %>
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
Function tit()
	tit=""
	dim id,rs,sql
	id=Request.QueryString("id")
	If Request.QueryString("Action")="Ye" Then
		tit="|ҵ����Ѷ"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="Select [title] From Bs_NewsYe where id="&cint(id)
		rs.Open sql,conn,1,1
		If NOT rs.EOF Then tit="|ҵ����Ѷ|" & rs(0)
		rs.close
		set rs=nothing
	End if 

	If Request.QueryString("Action")="Co" Then
		tit="|��˾����"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="Select [title] From Bs_NewsCo where id="&cint(id)
		rs.Open sql,conn,1,1
		If NOT rs.EOF Then tit="|��˾����|" & rs(0)
		rs.close
		set rs=nothing
	End if 

	If Request.QueryString("Action")="Pr" Then
		tit="|��Ʒ��̬"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="Select [title] From Bs_NewsPr where id="&cint(id)
		rs.Open sql,conn,1,1
		If NOT rs.EOF Then tit="|��Ʒ��̬|" & rs(0)
		rs.close
		set rs=nothing
	End if 

	If Request.QueryString("Bs_DownID")<>"" Then
		tit="|��������"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="Select [Bs_Title] From [Bs_Download] where Bs_DownID=" & cint(Request.QueryString("Bs_DownID"))
		rs.Open sql,conn,1,1
		If NOT rs.EOF Then tit="|��������|" & rs(0)
		rs.close
		set rs=nothing
	End if 

	If Request.QueryString("ArticleID")<>"" Then
		tit="|��Ʒչʾ"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="Select [Title] From [Bs_Product] where ArticleID=" & cint(Request.QueryString("ArticleID"))
		rs.Open sql,conn,1,1
		If NOT rs.EOF Then tit="|��Ʒչʾ|" & rs(0)
		rs.close
		set rs=nothing
	End if 

	If Request.QueryString("Action")="Profile" Then
		tit="|��˾����"
	End if 

	If Request.QueryString("Action")="Ceo" Then
		tit="|�ܲ��´�"
	End if 

	If Request.QueryString("Action")="Culture" Then
		tit="|��˾�Ļ�"
	End if 

	If Request.QueryString("Action")="Principle" Then
		tit="|��������"
	End if 

	If Request.QueryString("Action")="Contact" Then
		tit="|��ϵ����"
	End if 

	If Request.QueryString("Action")="Honor" Then
		tit="|��˾����"
	End if 

	If Request.QueryString("Action")="Img" Then
		tit="|��˾����"
	End if 

	If Request.QueryString("Action")="Sale" Then
		tit="|�����г�"
	End if 

	If Request.QueryString("Action")="Salea" Then
		tit="|�����г�"
	End if 

End Function 
%>
<HTML><HEAD><TITLE>��<%=BsCompanyName & Server.HTMLEncode((tit))%>��</TITLE>
<script language="javascript">
	var s;
	s=document.location.href;
	s=s.substring(s.lastIndexOf("/")+1);
	if (s=="Bs_Product.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|��Ʒչʾ��";
	}

	if (s=="Bs_Products.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|��Ʒ�����";
	}

	if (s=="Bs_Search.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|��Ʒ������";
	}
	
	if (s=="Bs_Job.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|�˲���Ƹ��";
	}

	if (s=="Bs_Jobs.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|�˲Ų��ԡ�";
	}

	if (s=="Bs_Download.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|�������ġ�";
	}

	if (s=="Bs_Faq.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|����֧�֡�";
	}

	if (s=="Bs_Server.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|��Ա���ġ�";
	}

	if (s=="Bs_Went.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|��Ϣ������";
	}

	if (s=="Bs_NetBook.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|�������ġ�";
	}

	if (s=="Bs_E_shop.asp")
	{
		document.title=document.title.substring(0,document.title.length-1) + "|������ѯ��";
	}
</script>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta name="Author" content="Myszw">
<meta name="Keywords" content="www.asbd.cn ��ҵ��վ����ϵͳMac3.0">
<meta name="Description" content="www.asbd.cn ��ҵ��վ����ϵͳMac3.0">
<META http-equiv=Pragma content=no-cache,must-revalidate>
<META http-equiv=Cache-Control content=no-cache>
<META http-equiv=Expires content="1999-01-01 00:00:00">
<META http-equiv=Last-Modified content="1999-01-01 00:00:00">
<link href="../Skin/<%=Request.Cookies("BsSkins")%>/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
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
<SCRIPT LANGUAGE="JavaScript" src="../Js/BsTopLogo.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="../js/BsTopMenu.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="../js/BsFootMenu.js"></SCRIPT>
</HEAD>

<!--body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" oncontextmenu="return false" onSelectStart="return false"-->
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<div id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<!--#include file="Bs_TopLogin.asp" -->
<%
Dim strSkins
strSkins = CInt(Request.Cookies("BsSkins"))
%>

<%if strSkins <= 100 and strSkins >= 1 then%>
<%Call BsTopLogin1()%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA1()</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">CnTopMenu(<%=strSkins%>)</SCRIPT>

<%elseif strSkins <= 150 and strSkins >= 101 then%>
<%Call BsTopLogin2()%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA2()</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">CnTopMenu(<%=strSkins%>)</SCRIPT>

<%elseif strSkins <= 160 and strSkins >= 151 then%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA3()</SCRIPT>
<%Call BsTopLogin3()%>
<SCRIPT LANGUAGE="JavaScript" src="../js/BsTopMenu2.js"></SCRIPT>

<%elseif strSkins <= 200 and strSkins >= 161 then%>
<%Call BsTopLogin3()%>
<SCRIPT LANGUAGE="JavaScript" src="../js/BsTopMenu2.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">TopLogoA1()</SCRIPT>

<%elseif strSkins <= 250 and strSkins >= 201 then%>
<%Call BsTopLogin4()%>
<SCRIPT LANGUAGE="JavaScript">TopLogoA4()</SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="../js/BsTopMenu4.js"></SCRIPT>


<%end if%>

<SCRIPT LANGUAGE="JavaScript">TopLogoB1()</SCRIPT>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='MiddleTitle'>
<TR><TD vAlign=top> 