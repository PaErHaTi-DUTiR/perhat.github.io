<!--#include file="bsconfig.asp"-->
<%
'=========================================================
'
'��Ʒ���ƣ��Ƽ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ ��ʽ��QQ ��812256
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
	strTitleName="�� ˾ �� �� �� ��"
	strContent="Profile"
	strEnContent="EnProfile"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
if yzcv<>zcv then response.end
elseif UrlName="Ceo" then
	strTitleName="�� �� �� �� �� ��"
	strContent="Ceo"
	strEnContent="EnCeo"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Culture" then
	strTitleName="�� ˾ �� �� �� ��"
	strContent="Culture"
	strEnContent="EnCulture"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Organize" then
	strTitleName="�� ֯ �� �� �� ��"
	strContent="Organize"
	strEnContent="EnOrganize"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Principle" then
	strTitleName="�� �� �� �� �� ��"
	strContent="Principle"
	strEnContent="EnPrinciple"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Contact" then
	strTitleName="�� ϵ �� Ϣ �� ��"
	strContent="Contact"
	strEnContent="EnContact"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Sale" then
	strTitleName="�� �� �� �� �� ��"
	strContent="Sale"
	strEnContent="EnSale"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
elseif UrlName="Salea" then
	strTitleName="�� �� �� �� �� ��"
	strContent="Salea"
	strEnContent="EnSalea"
%>
<!--#include file="Include/Bs_SysCompany.asp"-->
<%
end if
%>
