<!--#include file="bsconfig.asp"-->
<%
'=========================================================
'
'��Ʒ���ƣ��Ƽ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ ��ʽ��QQ ��812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%

dim sql
dim rs
on error resume next

set rs=server.createobject("adodb.recordset")
sql="delete from Bs_Jobbook where ID="&request("id")
rs.open sql,conn,1,1
set rs=nothing
conn.close
set conn=nothing
response.redirect "Bs_Job_Book.asp"
%>
