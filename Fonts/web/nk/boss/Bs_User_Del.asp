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
dim UserID,sql,rs

UserID=trim(Request("UserID"))
if UserID<>"" then
	sql="delete from [Bs_User] where UserID=" & Clng(UserID)
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_User.asp"
%>
