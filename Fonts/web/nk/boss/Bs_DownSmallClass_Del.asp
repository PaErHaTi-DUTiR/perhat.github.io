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
dim Bs_SmallClassID,sql

Bs_SmallClassID=trim(Request("Bs_SmallClassID"))
if Bs_SmallClassID<>"" then
	sql="delete from Bs_DownSmallClass where Bs_SmallClassID="&Cint(Bs_SmallClassID)&""
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_DownClass.asp"
%>


