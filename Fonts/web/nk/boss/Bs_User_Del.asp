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
dim UserID,sql,rs

UserID=trim(Request("UserID"))
if UserID<>"" then
	sql="delete from [Bs_User] where UserID=" & Clng(UserID)
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_User.asp"
%>
