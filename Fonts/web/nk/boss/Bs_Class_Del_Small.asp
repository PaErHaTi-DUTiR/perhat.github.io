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
dim SmallClassID,sql

SmallClassID=trim(Request("SmallClassID"))
if SmallClassID<>"" then
	sql="delete from Bs_PrSmallClass where SmallClassID="&Cint(SmallClassID)&""
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Bs_Class.asp"
%>


