<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<%
password=replace(trim(Request("pwd1")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin where UserName='" & request.form("UserName") & "'"'or RealName='" & request.form("realname") & "'and PassWord='" & password & "'"
rs.open sqltext,conn,1,1

'查找数据库，检查此管理员是否已经存在
if rs.recordcount >= 1 then
if rs("UserName")=request.form("UserName") then
%>
<script language='javascript'>alert ("此管理员帐号已经存在，请选用其它名称!"); location='Bs_Admin.asp';</script>
<%
response.end
rs.close
end if
end if


set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin"
rs.open sqltext,conn,3,3

'添加一个管理员帐号到数据库
rs.addnew
rs("UserName")=request.form("UserName")
rs("RealName")=request.form("RealName")
rs("PassWord")=password
rs.update
Response.Redirect "Bs_Admin.asp"

%>
