<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/md5.asp"-->
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
dim sql
dim rs
dim username
dim password
username=replace(trim(request("username")),"'","")
password=replace(trim(Request("password")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sql="select * from [Bs_User] where LockUser=False and username='" & username & "'and password='" & password &"'"
rs.open sql,conn,1,1
if not(rs.bof and rs.eof) then
	if password=rs("password") then
		session("UserName")=rs("username")
		'session("purview")=999
		Response.Redirect "Bs_Ment1.asp"
	end if
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
<html>
<head>
<style type="text/css">
A:link,A:active,A:visited{TEXT-DECORATION:none ;Color:#000000}
A:hover{Color:#4455aa}
BODY{
FONT-SIZE: 12px;
COLOR: #000000;
FONT-FAMILY:  ����;
background-color: #FFFFFF; 
scrollbar-face-color: #C0C0C0;
scrollbar-highlight-color: #C0C0C0;
scrollbar-shadow-color: #C0C0C0;
scrollbar-3dlight-color: #E0E0E0;
scrollbar-arrow-color:  #000000;
scrollbar-track-color: #E0E0E0;
scrollbar-darkshadow-color: #C0C0C0;
}
TD{
font-family: ����;
font-size: 12px;
line-height: 15px;
}
td.TableBody1
{
background-color: #137171;
}
.tableBorder1
{
width:97%;
border: 1px; 
background-color: #0D4F4F;
}
.table { border-collapse: collapse; border-left: 1px solid #000000; border-right: 1px solid #000000; background-color:#ffffff }
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #cccccc; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #cccccc; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #cccccc; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #cccccc}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body>
<br><br><br>
<table width='300'border='1' align='center'cellpadding='4'cellspacing='0' bordercolor="#666666" bgcolor="#CCCCCC" class='border'>
	<tr >
		<td height='15'colspan='2' align="center" background="images/b1.gif" class='title'>����: ȷ������ʧ��!</td>
	</tr>
	<tr>
		<td height='23'colspan='2' align="center" class='tdbg'
> <br><br>�û�����������󣡣���<br><br
> <a href='javascript:onclick=history.go(-1)'>�����ء�</a> <br><br></td>
	</tr>
</table>
</body></html>
<%
rs.close
set rs=nothing
call CloseConn()
%>