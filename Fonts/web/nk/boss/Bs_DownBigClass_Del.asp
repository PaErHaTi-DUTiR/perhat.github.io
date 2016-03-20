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
dim Bs_BigClassName,sql

Bs_BigClassName=trim(Request("Bs_BigClassName"))
if Bs_BigClassName<>"" then
	sql="delete from Bs_DownBigClass where Bs_BigClassName='" & Bs_BigClassName & "'"
	conn.Execute sql
	sql="delete from Bs_DownSmallClass where Bs_BigClassName='" & Bs_BigClassName & "'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_DownClass.asp"
%>


