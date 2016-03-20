<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
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
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin where Id=" & request.querystring("uid")
rs.open sqltext,conn,3,3

'更新管理员密码到数据库
rs("UserName")=request.form("UserName")
rs("PassWord")=md5(request.form("pwd1"))
rs.update
rs.close
conn.close
response.redirect "Bs_Admin.asp"
%>
