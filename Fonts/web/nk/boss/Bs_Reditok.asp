<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
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
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin where Id=" & request.querystring("uid")
rs.open sqltext,conn,3,3

'���¹���Ա���뵽���ݿ�
rs("UserName")=request.form("UserName")
rs("PassWord")=md5(request.form("pwd1"))
rs.update
rs.close
conn.close
response.redirect "Bs_Admin.asp"
%>
