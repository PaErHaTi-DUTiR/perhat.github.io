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
dim UserID,Action,sql

UserID=trim(Request("UserID"))
Action=trim(request("Action"))
if UserID<>"" then
	if Action="Lock" then
		sql="Update [Bs_User] set LockUser=true where UserID=" & CLng(UserID)
	else
		sql="Update [Bs_User] set LockUser=false where UserID=" & CLng(UserID)
	end if
	conn.Execute sql
    call CloseConn()      
end if
response.redirect "Bs_User.asp"
%>
