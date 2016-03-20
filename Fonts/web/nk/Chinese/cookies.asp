<%
select case Request("menu")

case "BsSkins"
Response.Cookies("BsSkins")=""&Request("no")&""
Response.Cookies("BsSkins").Expires=date+3650


url=Request.ServerVariables("http_referer")
if url<>empty and instr(url,"left.asp")=0 then
response.redirect url
else
response.write "<SCRIPT>location='index.asp';</SCRIPT>"
end if

case "eremite"
Response.Cookies("eremite")="1"
Response.Cookies("eremite").Expires=date+3650
response.redirect Request.ServerVariables("http_referer")


case "online"
Response.Cookies("eremite")="0"
Response.Cookies("eremite").Expires=date+3650
response.redirect Request.ServerVariables("http_referer")


end select
%>
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
