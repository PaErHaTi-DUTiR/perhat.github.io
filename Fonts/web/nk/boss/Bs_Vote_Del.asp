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
dim VoteID,sql
VoteID=trim(Request("ID"))
if VoteID<>"" then
	sql="delete from Vote where ID="&Clng(VoteID)&""
	conn.Execute sql
    call CloseConn()      
end if
response.redirect "Bs_Vote.asp"
%>
