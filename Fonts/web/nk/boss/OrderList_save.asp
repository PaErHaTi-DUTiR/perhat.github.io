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
djfc="�Ѿ�����"
set rs=server.createobject("adodb.recordset")
sqltext="select Flag from Bs_OrderList where Form_Id=" & request("Form_Id")
rs.open sqltext,conn,1,1
if rs("Flag")="�Ѿ�����" then
rs.close
Response.Redirect "OrderList_Save.asp?msg=�˶��������Ѿ������˷�������"
Else
set rs=server.createobject("adodb.recordset")
sqltext="update Bs_OrderList set Flag='"&djfc&"'where Form_Id=" & request("Form_Id")
rs.open sqltext,conn,3,3
response.redirect "OrderMessageBox.asp?msg=��������������ϣ��밴�ͻ���ϸ��ַ������"
end if
%>
