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
dim SQL, Rs, contentID,CurrentPage

CurrentPage = request("Page")
contentID=request("ID")

set rs=server.createobject("adodb.recordset")
sqltext="delete from Bs_OrderList where Form_Id="& contentID
rs.open sqltext,conn,3,3
set rs=nothing

set rs=server.createobject("adodb.recordset")
sqltext="delete from Bs_ShopList where Form_Id="& contentID
rs.open sqltext,conn,3,3
set rs=nothing
conn.close

response.redirect "Bs_Eshop.asp?page="&CurrentPage
%>
