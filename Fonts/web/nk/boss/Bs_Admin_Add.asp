<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
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
password=replace(trim(Request("pwd1")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin where UserName='" & request.form("UserName") & "'"'or RealName='" & request.form("realname") & "'and PassWord='" & password & "'"
rs.open sqltext,conn,1,1

'�������ݿ⣬���˹���Ա�Ƿ��Ѿ�����
if rs.recordcount >= 1 then
if rs("UserName")=request.form("UserName") then
%>
<script language='javascript'>alert ("�˹���Ա�ʺ��Ѿ����ڣ���ѡ����������!"); location='Bs_Admin.asp';</script>
<%
response.end
rs.close
end if
end if


set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin"
rs.open sqltext,conn,3,3

'���һ������Ա�ʺŵ����ݿ�
rs.addnew
rs("UserName")=request.form("UserName")
rs("RealName")=request.form("RealName")
rs("PassWord")=password
rs.update
Response.Redirect "Bs_Admin.asp"

%>
