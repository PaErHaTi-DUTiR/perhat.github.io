<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
'=========================================================
'
'��Ʒ���ƣ���˾(��ҵ)��վ����ϵͳ
'��Ȩ����: www.web300.cn
'����������web300Դ������
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
	  <% 
'����
set rs=server.createobject("adodb.recordset") 
sql="select * from email "
rs.open sql,conn,1,3  

'��ȡĬ�ϵ��ʼ����⼰���� 
set rs1=server.createobject("adodb.recordset")
sql1="select * from maildefault "
rs1.open sql1,conn,1,3  

'���÷�����
frommail=request("frommail")
if frommail="" then
frommail=rs1("frommail")
end if

'�����ʼ�����
mailsubject=request("mailsubject")
if mailsubject="" then
mailsubject=rs1("mailsubject")
end if

'�����ʼ�����
mailbody=request("mailbody")
if mailbody="" then
mailbody=rs1("mailbody")
end if

'�ж϶�˭����
tomail=request("tomail")
'д������Ϣ
response.write "�����˵�ַ: "&frommail
response.write "<br><br><br>"
if tomail<>"" then
response.write "�����˵�ַ��"&tomail
else
response.write "���ڽ����ʼ�Ⱥ����"
end if

if tomail<>"" then
'���ڵ�һ�û�����
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From = frommail
objCDOMail.To = tomail
objCDOMail.Subject = mailsubject  
objCDOMail.Body = mailbody   
objCDOMail.Send
Set objCDOMail = Nothing
else

'�������û����ݿ��е�ȫ���û�����
for i=1 to rs.recordcount
tomail=rs("email")
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From = frommail
objCDOMail.To = tomail
objCDOMail.Subject = mailsubject  
objCDOMail.Body = mailbody   
objCDOMail.Send
Set objCDOMail = Nothing
rs.movenext
next
end if
response.write "<br><br><br>"
response.write "�ʼ����ͳɹ���^&^"
'response.write "<br><br><br>"
'response.write rs1("mailsubject")
%> <a href="Bs_Mail_Send.asp">����</a>
			<BR><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
