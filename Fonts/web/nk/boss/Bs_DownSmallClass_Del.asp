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
dim Bs_SmallClassID,sql

Bs_SmallClassID=trim(Request("Bs_SmallClassID"))
if Bs_SmallClassID<>"" then
	sql="delete from Bs_DownSmallClass where Bs_SmallClassID="&Cint(Bs_SmallClassID)&""
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_DownClass.asp"
%>


