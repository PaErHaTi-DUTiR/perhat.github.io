<!--#include file="bsconfig.asp"-->
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
dim sqltext,rs,contentID,CurrentPage
dim strTitleName,strDataName,strUrlFile,UrlName,strSeeUrl
UrlName=trim(request("UrlName"))

if UrlName="Co" then
	strTitleName="�� ˾ �� �� �� ��"
	strDataName="Bs_NewsCo"
  strSeeUrl="Bs_NewsInfo.asp?Action=Co&"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
elseif UrlName="Ye" then
	strTitleName="ҵ �� �� Ѷ �� ��"
	strDataName="Bs_NewsYe"
  strSeeUrl="Bs_NewsInfo.asp?Action=Ye&"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
elseif UrlName="Pr" then
	strTitleName="�� Ʒ �� ̬ �� ��"
	strDataName="Bs_NewsPr"
  strSeeUrl="Bs_NewsInfo.asp?Action=Pr&"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
elseif UrlName="Faq" then
	strTitleName="�� �� �� �� �� ��"
	strDataName="Bs_Faq"
  strSeeUrl="Bs_FaqInfo.asp?"
%>
<!--#include file="Include/Bs_SysNews.asp"-->
<%
end if
%>
