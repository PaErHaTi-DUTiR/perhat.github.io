<!--#include file="../Inc/Conn.asp"-->

<!--#include file="../Inc/articleCHAR.INC"-->
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
If request.form("Comane")="" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入公司名称，请返回检查！！"");history.go(-1);</script>")
response.end
end if
If request.form("Somane")="" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入联系人，请返回检查！！"");history.go(-1);</script>")
response.end
end if
If request.form("Phone")="" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入联系电话，请返回检查！！"");history.go(-1);</script>")
response.end
end if
If request.form("title")="" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入标题，请返回检查！！"");history.go(-1);</script>")
response.end
end if
If request.form("content")="" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入留言内容，请返回检查！！"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Book"
rs.open sql,conn,1,3


rs.addnew
rs("name")=htmlencode2(request.form("name"))
rs("Comane")=htmlencode2(request.form("Comane"))
rs("Somane")=htmlencode2(request.form("Somane"))
rs("Phone")=htmlencode2(request.form("Phone"))
rs("Fox")=htmlencode2(request.form("Fox"))
rs("email")=htmlencode2(request.form("email"))
rs("title")=htmlencode2(request.form("title"))
rs("content")=htmlencode2(request.form("content"))
rs("time")=date()
rs.update

rs.close
set rs=nothing
call CloseConn()

if request.form("name")="未注册用户" then
response.redirect "Bs_Wtok.asp"
else
response.redirect "Bs_NetBook.asp"
end if
%>
