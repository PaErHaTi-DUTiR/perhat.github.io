<!--#include file="bsconfig.asp"-->
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
dim sqltext,rs,contentID,CurrentPage
dim strTitleName,strDataName,strUrlFile,UrlName,strSeeUrl
UrlName=trim(request("UrlName"))

if UrlName="Co" then
	strTitleName="公 司 新 闻 管 理"
	strDataName="Bs_NewsCo"
  strSeeUrl="Bs_NewsInfo.asp?Action=Co&"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
elseif UrlName="Ye" then
	strTitleName="业 内 资 讯 管 理"
	strDataName="Bs_NewsYe"
  strSeeUrl="Bs_NewsInfo.asp?Action=Ye&"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
elseif UrlName="Pr" then
	strTitleName="产 品 动 态 管 理"
	strDataName="Bs_NewsPr"
  strSeeUrl="Bs_NewsInfo.asp?Action=Pr&"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
elseif UrlName="Faq" then
	strTitleName="问 题 解 答 管 理"
	strDataName="Bs_Faq"
  strSeeUrl="Bs_FaqInfo.asp?"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
end if
%>
