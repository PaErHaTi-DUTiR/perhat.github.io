<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->

<%
'=========================================================
'
'��Ʒ���ƣ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ����: www.web300.cn 
'����������www.web300.cn�����Ŷ�
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>

<%
dim strTitleName,strDataName,strUrlFile,UrlName
UrlName=trim(request("UrlName"))

if UrlName="Co" then
	strTitleName="�� �� �� ˾ �� ��"
	strDataName="Bs_NewsCo"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
elseif UrlName="Ye" then
	strTitleName="�� �� ҵ �� �� Ѷ"
	strDataName="Bs_NewsYe"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
elseif UrlName="Pr" then
	strTitleName="�� �� �� Ʒ �� ̬"
	strDataName="Bs_NewsPr"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
elseif UrlName="Faq" then
	strTitleName="�� �� �� �� �� ��"
	strDataName="Bs_Faq"
%>
<!--#include file="Include/Bs_SysNewsAdd.asp"-->
<%
end if
%>
