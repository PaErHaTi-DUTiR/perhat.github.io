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
dim BigClassName,sql

BigClassName=trim(Request("BigClassName"))
if BigClassName<>"" then
	sql="delete from Bs_PrBigClasss where BigClassName='" & BigClassName & "'"
	conn.Execute sql
	sql="delete from Bs_PrSmallClass where BigClassName='" & BigClassName & "'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_Class.asp"
%>


