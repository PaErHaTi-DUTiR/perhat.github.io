<!--#include file="../Inc/Conn.asp"-->

<%
dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="delete from email where email='"&request("email")&"'"
rs.open sql,conn,3,3
set rs=nothing
conn.close
set conn=nothing
response.write"<SCRIPT language=JavaScript>alert('���ѳɹ��˶���վ���ʼ��б�����ӭ�´��ٶ��ģ�');"
  response.write"javascript:history.go(-1)</SCRIPT>"
%>