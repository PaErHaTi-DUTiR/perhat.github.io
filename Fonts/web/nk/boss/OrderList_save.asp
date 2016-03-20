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
djfc="已经发货"
set rs=server.createobject("adodb.recordset")
sqltext="select Flag from Bs_OrderList where Form_Id=" & request("Form_Id")
rs.open sqltext,conn,1,1
if rs("Flag")="已经发货" then
rs.close
Response.Redirect "OrderList_Save.asp?msg=此订货单您已经做过了发货处理！"
Else
set rs=server.createobject("adodb.recordset")
sqltext="update Bs_OrderList set Flag='"&djfc&"'where Form_Id=" & request("Form_Id")
rs.open sqltext,conn,3,3
response.redirect "OrderMessageBox.asp?msg=订货单处理处理完毕，请按客户详细地址发货！"
end if
%>
