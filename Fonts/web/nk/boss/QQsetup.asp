<!--#include file="Inc/Function.asp"-->

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
Dim StarTime,Style_Copy
Dim AdminName
startime=timer()
AdminName=replace(session("AdminName"),"'","")
if AdminName="" then
	call CloseConn()
%>
<script language='javascript'>top.location='Login.asp';</script>
<%
'	response.redirect "login.asp"
end if
%>
<!--  -->
<%
sub htmlend
%>
<p align=center>�Ƽ����޹�˾</p>
<%
end sub
%>



<script language=javascript src=../inc/mouse_on_title.js></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY>
<!-- #include file="Inc/Head.asp" -->
<%
Dim connQQstr
connQQstr = "Provider=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE="&SERVER.MAPPATH("../QQ/QQ.mdb")
Set connQQ=server.CreateObject("adodb.connection")
connQQ.Open connQQstr
action=request("ok")
call aspsql()
if action="" then 
Set rs = connQQ.Execute("select * from QQsetup") 
%>

<table width="98%" border="1" style="border-collapse: collapse; border-style: dotted; border-width: 0px" bordercolor="#333333" cellspacing="0" cellpadding="2">
<form action=QQsetup.asp method=post name=setup>
<tr class=backs><td colspan=2 class=td height=18>������ѯ���� &nbsp;</td></tr>
<tr><td width=20% align=right>��������ѯ��״̬ &nbsp;</td>
<td >
<input type="radio" name="qqonline" value="1" <%if rs("qqonline")=1 then%>checked<%end if%>>��ʾ
<input type="radio" name="qqonline" value="0" <%if rs("qqonline")=0 then%>checked<%end if%>>���� 
</td></TR>
<tr><td width=20% align=right>��ʾλ�� &nbsp;</td>
<td >
<input type="radio" name="whereqq" value="1" <%if rs("whereqq")=1 then%>checked<%end if%>>���
<input type="radio" name="whereqq" value="0" <%if rs("whereqq")=0 then%>checked<%end if%>>�Ҳ� &nbsp; <img src=../images/admin/memo.gif alt="��������Ҳ�"></td></TR>
<tr><td width=20% align=right>�ͷ�QQ�� &nbsp;</td><td > <input type=text value="<%=rs("adm_qq")%>" size=35 name="adm_qq" maxlength="50" size=50> &nbsp; <img src=../images/admin/memo.gif alt="����QQ��<br>���QQ�������Ķ��Ÿ���<br>ע�⣺QQ��Ӧ��������ǳ�һһ��Ӧ"> </td></TR>
<tr><td width=20% align=right>�ͷ�QQ��Ӧ�ǳ� &nbsp;</td><td > <input type=text value="<%=rs("adm_qq_name")%>" size=35 name="adm_qq_name" maxlength="50" size=50> &nbsp; <img src=../images/admin/memo.gif alt="����QQ�ǳ�<br>����ǳ������Ķ��Ÿ���<br>ע�⣺�ǳ�Ӧ�������QQ��һһ��Ӧ"> </td></TR>
<tr><td width=20% align=right>��ѡ��ͷ�� &nbsp;</td>
<td>
<table width=50% border=0>
<tr><td><input type="radio" name="qqskin" value="1" <%if rs("qqskin")=1 then%>checked<%end if%>>���ߣ�</td><td><img src=../images/qq/qqon1.gif border=0> </td><td> ���ߣ�</td><td><img src=../images/qq/qqoff1.gif border=0> &nbsp; &nbsp; &nbsp; &nbsp; <img src=../images/admin/memo.gif alt="��������QQ�ǳƽϳ�ʱ<br>��ʹ�õ�һ��Сͷ��<br>����ᳬ�����ڿ���"></td></tr>
<tr><td><input type="radio" name="qqskin" value="2" <%if rs("qqskin")=2 then%>checked<%end if%>>���ߣ�</td><td><img src=../images/qq/qqon2.gif border=0> </td><td> ���ߣ� </td><td> <img src=../images/qq/qqoff2.gif border=0> </td></tr> 
<tr><td><input type="radio" name="qqskin" value="3" <%if rs("qqskin")=3 then%>checked<%end if%>>���ߣ�</td><td><img src=../images/qq/qqon3.gif border=0> </td><td> ���ߣ� </td><td> <img src=../images/qq/qqoff3.gif border=0> </td></tr> 
<tr><td><input type="radio" name="qqskin" value="4" <%if rs("qqskin")=4 then%>checked<%end if%>>���ߣ�</td><td><img src=../images/qq/qqon4.gif border=0> </td><td> ���ߣ� </td><td> <img src=../images/qq/qqoff4.gif border=0> </td></tr> 
<tr><td><input type="radio" name="qqskin" value="5" <%if rs("qqskin")=5 then%>checked<%end if%>>���ߣ�</td><td><img src=../images/qq/qqon5.gif border=0> </td><td> ���ߣ� </td><td> <img src=../images/qq/qqoff5.gif border=0> </td></tr> 
</table>
</td></TR>
<tr><td width=20% align=right>��ѡ����ʽ &nbsp;</td>
<td >
<input type="radio" name="kefuskin" value="1" <%if rs("kefuskin")=1 then%>checked<%end if%>>��ʽһ	<input type="radio" name="kefuskin" value="2" <%if rs("kefuskin")=2 then%>checked<%end if%>>��ʽ��	<input type="radio" name="kefuskin" value="3" <%if rs("kefuskin")=3 then%>checked<%end if%>>��ʽ��	<input type="radio" name="kefuskin" value="4" <%if rs("kefuskin")=4 then%>checked<%end if%>>��ʽ��	<input type="radio" name="kefuskin" value="5" <%if rs("kefuskin")=5 then%>checked<%end if%>>��ʽ��	<a href=../images/help/qq.gif target=_blank><img src=../images/admin/memo.gif border=0 alt="���Ԥ��Ч��"></a>
</td></TR>
<tr><td width=20% align=right>������ʾ�� &nbsp;</td><td > 
<TEXTAREA NAME="qqmsg_on" ROWS="3" COLS="34" style="overflow:auto;"><%=rs("qqmsg_on")%></TEXTAREA> &nbsp; <img align=top src=../images/admin/memo.gif alt="�ͷ�����ʱ�����ָ��ͷ����ǳ�<br>����ʾ�����Ի���ʾ��<br>������ӡ�br����"></td></tr>
<tr><td width=20% align=right>������ʾ�� &nbsp;</td><td >
<TEXTAREA NAME="qqmsg_off" ROWS="3" COLS="34" style="overflow:auto;"><%=rs("qqmsg_off")%></TEXTAREA> &nbsp; <img align=top src=../images/admin/memo.gif alt="�ͷ�����ʱ�����ָ��ͷ����ǳ�<br>����ʾ�����Ի���ʾ��<br>������ӡ�br����"></td></tr>
<tr><td colspan=2 width="568" ><INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="��������"> &nbsp;</td></tr>
</form>
</table>
<%
rs.close
set rs=nothing
connQQ.close
set connQQ=Nothing
end If
%>
<%
if action="ok" Then
if request.form("qqonline")="" or request.form("whereqq")="" then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ�������д���������д���');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if request.form("qqonline")=1 then
if trim(request.form("adm_qq"))="" or trim(request.form("qqmsg_on"))="" or trim(request.form("qqmsg_off"))="" then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ�������д���������д���');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
end if
if request.form("adm_qq_name")<>"" and request.form("adm_qq")<>"" then
adm_qq_name=replace(request.form("adm_qq_name"),"��",",")
adm_qq=replace(request.form("adm_qq"),"��",",")
if UBound(split(adm_qq_name,","))<> UBound(split(adm_qq,",")) then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ��ͷ�QQ�������ǳ�δһһ��Ӧ��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from QQsetup"
rs.open sql,connQQ,1,3
rs("qqonline")=request.form("qqonline")
rs("whereqq")=request.form("whereqq")
if request.form("qqonline")=1 then 
rs("kefuskin")=request.form("kefuskin")
rs("qqskin")=request.form("qqskin")
qqmsg_off=replace(trim(request.form("qqmsg_off")),"'","")
qqmsg_off=replace(qqmsg_off,"&quot;","")
rs("qqmsg_off")=qqmsg_off
qqmsg_on=replace(trim(request.form("qqmsg_on")),"'","")
qqmsg_on=replace(qqmsg_on,"&quot;","")
rs("qqmsg_on")=qqmsg_on
rs("adm_qq")=replace(request.form("adm_qq"),"��",",")
rs("adm_qq_name")=replace(request.form("adm_qq_name"),"��",",")
if request.form("adm_qq_name")="" then rs("adm_qq_name")=replace(request.form("adm_qq"),"��",",") 
end if
rs.update
url="QQsetup.asp"
rs.close
set rs=nothing
connQQ.close
set connQQ=nothing
response.write "<script language='javascript'>"
response.write "alert('�����ɹ��������õ���Ϣ�ѱ��棡');"
response.write "location.href='"&url&"';"
response.write "</script>"
end If

Function msg(msgstr)
	Response.Write msgstr
	Response.End 
End Function 

sub aspsql()
SQL_injdata = "'|;|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare"
SQL_inj = split(SQL_Injdata,"|")
If Request.Form<>"" Then
For Each Sql_Post In Request.Form
For SQL_Data=0 To Ubound(SQL_inj)
if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then
response.write "<script language='javascript'>"
response.write "alert('��վ��ȫ��ʾ���벻Ҫ�ڲ����а����Ƿ��ַ���');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
next
next
end if
If Request.QueryString<>"" Then
For Each SQL_Get In Request.QueryString
For SQL_Data=0 To Ubound(SQL_inj)
if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_DATA))>0 Then
response.write "<script language='javascript'>"
response.write "alert('��վ��ȫ��ʾ���벻Ҫ�ڲ����а����Ƿ��ַ���');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
next
Next
end If
end sub
%>