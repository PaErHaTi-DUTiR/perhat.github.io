<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
'=========================================================
'
'产品名称：公司(企业)网站管理系统
'版权所有: www.web300.cn
'程序制作：web300源码中心
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>发 送 邮 件</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
	  <% 
'发送
set rs=server.createobject("adodb.recordset") 
sql="select * from email "
rs.open sql,conn,1,3  

'读取默认的邮件标题及内容 
set rs1=server.createobject("adodb.recordset")
sql1="select * from maildefault "
rs1.open sql1,conn,1,3  

'设置发信人
frommail=request("frommail")
if frommail="" then
frommail=rs1("frommail")
end if

'设置邮件主题
mailsubject=request("mailsubject")
if mailsubject="" then
mailsubject=rs1("mailsubject")
end if

'设置邮件内容
mailbody=request("mailbody")
if mailbody="" then
mailbody=rs1("mailbody")
end if

'判断对谁发信
tomail=request("tomail")
'写发信信息
response.write "发信人地址: "&frommail
response.write "<br><br><br>"
if tomail<>"" then
response.write "收信人地址："&tomail
else
response.write "正在进行邮件群发！"
end if

if tomail<>"" then
'对于单一用户发信
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From = frommail
objCDOMail.To = tomail
objCDOMail.Subject = mailsubject  
objCDOMail.Body = mailbody   
objCDOMail.Send
Set objCDOMail = Nothing
else

'对于在用户数据库中的全体用户发信
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
response.write "邮件发送成功！^&^"
'response.write "<br><br><br>"
'response.write rs1("mailsubject")
%> <a href="Bs_Mail_Send.asp">返回</a>
			<BR><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
