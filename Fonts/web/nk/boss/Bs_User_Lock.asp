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
dim UserID,Action,sql

UserID=trim(Request("UserID"))
Action=trim(request("Action"))
if UserID<>"" then
	if Action="Lock" then
		sql="Update [Bs_User] set LockUser=true where UserID=" & CLng(UserID)
	else
		sql="Update [Bs_User] set LockUser=false where UserID=" & CLng(UserID)
	end if
	conn.Execute sql
    call CloseConn()      
end if
response.redirect "Bs_User.asp"
%>
