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
