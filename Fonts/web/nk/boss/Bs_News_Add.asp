<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->

<%
'=========================================================
'
'产品名称： 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有: www.web300.cn 
'程序制作：www.web300.cn开发团队
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>

<%
dim strTitleName,strDataName,strUrlFile,UrlName
UrlName=trim(request("UrlName"))

if UrlName="Co" then
	strTitleName="添 加 公 司 新 闻"
	strDataName="Bs_NewsCo"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
elseif UrlName="Ye" then
	strTitleName="添 加 业 内 资 讯"
	strDataName="Bs_NewsYe"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
elseif UrlName="Pr" then
	strTitleName="添 加 产 品 动 态"
	strDataName="Bs_NewsPr"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
elseif UrlName="Faq" then
	strTitleName="添 加 问 题 解 答"
	strDataName="Bs_Faq"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
end if
%>
