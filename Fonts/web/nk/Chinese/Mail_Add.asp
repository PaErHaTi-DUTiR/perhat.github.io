<!--#include file="../Inc/Conn.asp"-->

<%
'������������������������������������������������������������
'�������������� COPYRIGHT ������������ ��
'���������ƣ���ҵ��վ����ϵͳMac3.0  (ASBDcorpweb Mac3.0)  �� 
'����Ȩ���У�WWW.ASBD.CN  BBS.ASBD.CN                      ��
'������������amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  �� 
'��Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ��
'������������������������������������������������������������
%>
<%
dim rs
dim sql

set rs=server.createobject("adodb.recordset")
sql="select * from email where email='"&email&"'"
rs.open sql,conn,3,3
rs.addnew
rs("email")=request("email")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
  response.write"<SCRIPT language=JavaScript>alert('��л�����ı�վ���ʼ��б�');"
  response.write"javascript:history.go(-1)</SCRIPT>"
%>
