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
dim Bs_BigClassName,sql

Bs_BigClassName=trim(Request("Bs_BigClassName"))
if Bs_BigClassName<>"" then
	sql="delete from Bs_DownBigClass where Bs_BigClassName='" & Bs_BigClassName & "'"
	conn.Execute sql
	sql="delete from Bs_DownSmallClass where Bs_BigClassName='" & Bs_BigClassName & "'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_DownClass.asp"
%>


