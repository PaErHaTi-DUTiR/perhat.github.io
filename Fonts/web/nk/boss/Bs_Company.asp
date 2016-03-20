<!--#include file="bsconfig.asp"-->
<%
'=========================================================
'
'产品名称：科技 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有：www.web300.cn
'程序制作：QQ812256@hotmail.com
'联系 方式：QQ ：812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>

<%
dim sql,rs,rs2
dim content,strContent,EnContent,strEnContent
dim strTitleName,strDataName,strUrlFile,UrlName,strSeeUrl
UrlName=trim(request("UrlName"))

if UrlName="Profile" then
	strTitleName="公 司 介 绍 设 置"
	strContent="Profile"
	strEnContent="EnProfile"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
if yzcv<>zcv then response.end
elseif UrlName="Ceo" then
	strTitleName="总 裁 介 绍 设 置"
	strContent="Ceo"
	strEnContent="EnCeo"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Culture" then
	strTitleName="公 司 文 化 设 置"
	strContent="Culture"
	strEnContent="EnCulture"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Organize" then
	strTitleName="组 织 机 构 设 置"
	strContent="Organize"
	strEnContent="EnOrganize"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Principle" then
	strTitleName="领 导 关 怀 设 置"
	strContent="Principle"
	strEnContent="EnPrinciple"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Contact" then
	strTitleName="联 系 信 息 设 置"
	strContent="Contact"
	strEnContent="EnContact"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Sale" then
	strTitleName="国 内 市 场 设 置"
	strContent="Sale"
	strEnContent="EnSale"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Salea" then
	strTitleName="国 际 市 场 设 置"
	strContent="Salea"
	strEnContent="EnSalea"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
end if
%>
