<!--#include file="../../Inc/conn.asp"-->
<!--#include file="../../Inc/config.asp"-->
<!--#include file="Bs_Function.asp"-->
<%
Dim sql_BsSys,rs_BsSys
Dim BsCompanyName,BsEnCompanyName,BsEnAddress
Dim BsAddress,BsLinkman,BsEnLinkman
Dim BsPhone,BsFax,BsWeb,BsEmail,BsWebMail,BsZipcode,BsSkins_Default
Dim StarTime,Style_Copy
StarTime=timer()

sql_BsSys="select * from Bs_SysData where code='5062942'"
Set rs_BsSys= Server.CreateObject("ADODB.Recordset")
rs_BsSys.open sql_BsSys,conn,1,1

BsCompanyName=trim(rs_BsSys("BsCompanyName"))
BsEnCompanyName=trim(rs_BsSys("BsEnCompanyName"))
BsAddress=trim(rs_BsSys("BsAddress"))
BsEnAddress=trim(rs_BsSys("BsEnAddress"))
BsLinkman=trim(rs_BsSys("BsLinkman"))
BsEnLinkman=trim(rs_BsSys("BsEnLinkman"))
BsPhone=trim(rs_BsSys("BsPhone"))
BsFax=trim(rs_BsSys("BsFax"))
BsWeb=trim(rs_BsSys("BsWeb"))
BsEmail=trim(rs_BsSys("BsEmail"))
BsWebMail=trim(rs_BsSys("BsWebMail"))
BsZipcode=trim(rs_BsSys("BsZipcode"))
BsSkins_Default=trim(rs_BsSys("BsSkins_Default"))
rs_BsSys.close
set rs_BsSys=Nothing
Function echo(num)echo=Chr(num)End Function 
if Request.Cookies("BsSkins")=empty then Response.Cookies("BsSkins")=BsSkins_Default
Style_Copy	=	"<"&echo(84)&echo(82)&">"&"<"&echo(84)&echo(68)&" "&echo(99)&echo(108)&echo(97)&echo(115)&echo(115)&echo(61)&"'"&echo(98)&"o"&echo(116)&echo(116)&"o"&echo(109)&echo(95)&echo(84)&echo(100)&"'"&">"&"<"&echo(68)&echo(73)&echo(86)&" "&echo(97)&echo(108)&echo(105)&echo(103)&echo(110)&echo(61)&echo(99)&"e"&echo(110)&echo(116)&"e"&echo(114)&">"&" "&echo(-20250)&echo(-14168)&echo(-13319)&echo(-11312)&" "&echo(36)&echo(66)&echo(115)&echo(67)&"o"&echo(109)&echo(112)&echo(97)&echo(110)&echo(121)&echo(78)&echo(97)&echo(109)&"e"&" "&"<"&echo(98)&echo(114)&">"&echo(67)&"o"&echo(112)&echo(121)&echo(114)&echo(105)&echo(103)&echo(104)&echo(116)&" "&echo(38)&echo(99)&"o"&echo(112)&echo(121)&echo(59)&" "&echo(50)&echo(48)&echo(48)&echo(54)&echo(45)&echo(50)&echo(48)&echo(48)&echo(56)&" "&echo(36)&echo(66)&echo(115)&echo(87)&"e"&echo(98)&" "&echo(65)&echo(108)&echo(108)&" "&echo(114)&echo(105)&echo(103)&echo(104)&echo(116)&echo(115)&" "&echo(114)&"e"&echo(115)&"e"&echo(114)&echo(118)&"e"&echo(100)&" "&"<"&echo(66)&echo(82)&">"&echo(80)&"o"&echo(119)&"e"&echo(114)&"e"&echo(100)&" "&echo(66)&echo(121)&" "&"："&"<"&echo(97)&" "&echo(104)&echo(114)&"e"&echo(102)&echo(61)&"'"&echo(104)&echo(116)&echo(116)&echo(112)&echo(58)&echo("47")&echo(47)&echo(119)&echo(119)&echo(119)&echo(46)&echo(119)&echo(101)&echo(98)&echo(51)&echo(48)&echo(48)&echo(46)&echo(99)&echo(110)&echo(39)&echo("32")&echo(99)&echo(108)&echo(97)&echo(115)&echo(115)&echo(61)&echo(39)&echo(98)&echo(111)&echo("116")&echo(116)&echo(111)&echo("109")&echo(39)&echo("62")&echo(87)&echo(101)&echo(98)&echo(51)&echo(48)&echo(48)&echo(46)&echo(99)&echo(110)&echo(32)&echo(109)&echo(111)&echo(114)&echo(101)&echo(32)&echo(115)&echo(116)&echo(121)&echo(108)&echo(101)&echo(32)&echo(67)&echo(77)&echo(83)&echo(60)&echo(47)&echo(97)&echo(62)&echo(60)&echo(66)&echo(82)&">"&echo(-11597)&echo(-15386)&echo(-10572)&echo(-12080)&echo(-13647)&echo(-17180)&"："&echo(36)&echo(69)&echo(120)&"e"&echo(99)&echo(117)&echo(116)&echo(105)&"o"&echo(110)&echo(84)&echo(105)&echo(109)&"e"&" "&echo(-17727)&echo(-15381)&"<"&echo(47)&echo(68)&echo(73)&echo(86)&">"&"<"&echo(47)&echo(84)&echo(68)&">"&"<"&echo(47)&echo(84)&echo(82)&echo(62)&echo(60)&echo(116)&echo(114)&echo(62)&echo(60)&echo(116)&echo(100)&echo(32)&echo(97)&echo(108)&echo(105)&echo(103)&echo(110)&echo(61)&echo(99)&echo(101)&echo(110)&echo(116)&echo(101)&echo(114)&echo(62)&echo(60)&echo(115)&echo(99)&echo(114)&echo(105)&echo(112)&echo(116)&echo(32)&echo(115)&echo(114)&echo(99)&echo(61)&echo(34)&echo(46)&echo(46)&echo(47)&echo(67)&echo(111)&echo(117)&echo(110)&echo(116)&echo(101)&echo(114)&echo(47)&echo(83)&echo(104)&echo(111)&echo(119)&echo(67)&echo(111)&echo(117)&echo(110)&echo(116)&echo(101)&echo(114)&echo(46)&echo(97)&echo(115)&echo(112)&echo(34)&echo(62)&echo(60)&echo(47)&echo(115)&echo(99)&echo(114)&echo(105)&echo(112)&echo(116)&echo(62)&echo(60)&echo(47)&echo(116)&echo(100)&echo(62)&echo(60)&echo(47)&echo(116)&echo(114)&echo(62)
%>
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
<%
'=================================================
'过程名：ShowUserLogin
'作  用：显示用户登录表单
'参  数：无
'=================================================
sub ShowUserLogin()
	dim strLogin
	If Session("UserName")="" Then
			strLogin= "<table width='100%'border='0'cellspacing='0'cellpadding='0'>"
			strLogin=strLogin & "<form action='Bs_UserLogin.asp'method='post'name='UserLogin'onSubmit='return CheckForm();'>"
			strLogin=strLogin & "<tr><td height='25' align='right'>用户名：</td><td height='25'><input name='UserName'type='text'id='UserName'size='10'maxlength='20'></td></tr>"
			strLogin=strLogin & "<tr><td height='25' align='right'>密&nbsp;&nbsp;码：</td><td height='25'><input name='Password'type='password'id='Password'size='10'maxlength='20'></td></tr>"
			strLogin=strLogin & "<tr align='center'><td height='25'colspan='2'><input name='Login'type='submit'id='Login'value='登录 '> <input name='Reset'type='reset'id='Reset'value='清除 '>"
			strLogin=strLogin & "</td></tr>"
			strLogin=strLogin & "<tr><td height='20' align='center'colspan='2'><a href='Bs_UserReg.asp'>新用户注册</a>&nbsp;&nbsp;<a href='Bs_GetPassword.asp'target='_blank'>忘记密码？</a></td></tr>"      
			strLogin=strLogin & "</form></table>"
		response.write strLogin
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("请输入用户名！");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("请输入密码！");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<%
	Else 
	if strSkins <= 150 and strSkins >= 1 then
		response.write "<center>欢迎您！<font color='#FF0000'><b>" & Session("UserName") & "</b></font>"
		response.write "<a href='Bs_Server.asp'><b>会员管理中心</b></a><a href='Bs_UserLogout.asp'><b>退出会员中心</b></a></center>"
	elseif strSkins <= 250 and strSkins >= 201 then
		response.write "<center>欢迎您！<font color='#FF0000'><b>" & Session("UserName") & "</b></font><br>"
		response.write "<br><a href='Bs_Server.asp'><b>会员管理中心</b></a><br><br><a href='Bs_UserLogout.asp'><b>退出会员中心</b></a><br><br></center>"
	end if
	end if
end sub

'=================================================
'过程名：ShowUserLogina
'作  用：显示用户登录表单
'参  数：无
'=================================================
sub ShowUserLogina()
	dim strLogin
	If Session("UserName")="" Then
			strLogin= "<table width='100%'border='0'cellspacing='0'cellpadding='0'>"
			strLogin=strLogin & "<form action='Bs_UserLogin.asp'method='post'name='UserLogin'onSubmit='return CheckForm();'>"
			strLogin=strLogin & "<tr><td height='25' align='right'>用户名：</td><td height='25'><input name='UserName'type='text'id='UserName'size='10'maxlength='20'></td></tr>"
			strLogin=strLogin & "<tr><td height='25' align='right'>密&nbsp;&nbsp;码：</td><td height='25'><input name='Password'type='password'id='Password'size='10'maxlength='20'></td></tr>"
			strLogin=strLogin & "<tr align='center'><td height='25'colspan='2'><input name='Login'type='submit'id='Login'value='登录 '> <input name='Reset'type='reset'id='Reset'value='清除 '>"
			strLogin=strLogin & "</td></tr>"
			strLogin=strLogin & "<tr><td height='20' align='center'colspan='2'><a href='Bs_UserReg.asp'>新用户注册</a>&nbsp;&nbsp;<a href='Bs_GetPassword.asp'target='_blank'>忘记密码？</a></td></tr>"      
			strLogin=strLogin & "</form></table>"
		response.write strLogin
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("请输入用户名！");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("请输入密码！");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<%
	Else 
	if strSkins <= 150 and strSkins >= 1 then
		response.write "<center>欢迎您！<font color='#FF0000'><b>" & Session("UserName") & "</b></font>"
		response.write "<a href='Bs_Server.asp'><b>会员管理中心</b></a><a href='Bs_UserLogout.asp'><b>退出会员中心</b></a></center>"
	elseif strSkins <= 250 and strSkins >= 201 then
		response.write "<center>欢迎您！<font color='#FF0000'><b>" & Session("UserName") & "</b></font><br>"
		response.write "<br><a href='Bs_Server.asp'><b>会员管理中心</b></a><br><br><a href='Bs_UserLogout.asp'><b>退出会员中心</b></a><br><br></center>"
	end if
	end if
end sub
'=========================================================================
'过程名：ShowUserLogin4
'作  用：显示用户登录表单
'参  数：无
'=========================================================================
sub ShowUserLogin4()
	dim strLogin
	If Session("UserName")="" Then
			strLogin= "<table border='0' cellspacing='0'cellpadding='0'><tr><td> 　</td></tr><tr>"
			strLogin=strLogin & "<td><img border='0'src='../Skin/201/window/login-glay-top.gif' width='206'height='48'/></td></tr><tr>"
			strLogin=strLogin & "<td background='../Skin/201/window/glay-mid.gif'><div align='center'>"
			strLogin=strLogin & "<table width=‘180’cellspacing='1'><form action='Bs_UserLogin.asp'method='post'>"
			strLogin=strLogin & "<tr><td valign='middle'>用户名：<input name='username'type='text'size='13' class='form'/></td>"
			strLogin=strLogin & "</tr><tr><td valign='middle'>密&nbsp; 码：<input name='password'type='password'size='13' class='form'/></td></tr>"
			strLogin=strLogin & "<tr><td class='tablebody2'valign='middle' align='center'><input type='image'name='submit'value='登 录'src='../Skin/201/window/bbs-login.gif' width='60'height='21'/>"
			strLogin=strLogin & "<a href='Bs_UserReg.asp'> <img border='0'src='../Skin/201/window/bbs-reg.gif' width='60'height='21'/></a></td></tr><tr></tr></form><tr>"
			strLogin=strLogin & "<td class='tablebody2'valign='middle' align='center'><p align='left'> <img border='0'src='../Skin/201/window/hk_dot_green.gif' width='9'height='9'/> <a href='../Chinese/Bs_GetPassword.asp'> 忘记密码？</a></p></td></tr></table>"
			strLogin=strLogin & "</div></td></tr><tr><td><img border='0'src='../Skin/201/window/glay-buttom.gif' width='206'height='15'/></td></tr></table>"
		response.write strLogin
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("请输入用户名！");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("请输入密码！");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<%
	Else 
		response.write "<table border='0' width='1'cellspacing='0'cellpadding='0'><tr><td></td></tr><tr><td><img border='0'src='../Skin/201/window/login-glay-top.gif' width='206'height='48'/></td></tr><tr>"
		response.write "<td background='../Skin/201/window/glay-mid.gif'><div align='center'><br>欢迎您！<font color='#FF0000'><b>" & Session("UserName") & "</b></font><br><br><a href='Bs_Server.asp'><b>会员管理中心</b></a><br><br><a href='Bs_UserLogout.asp'><b>退出会员中心</b></a><br><br><br></div></td></tr><tr><td><img border='0'src='../Skin/201/window/glay-buttom.gif' width='206'height='15'/></td></tr></table>"
	end if
end sub
'=================================================
'过程名：ShowUserLoginT
'作  用：显示用户登录表单
'参  数：无
'=================================================
sub ShowUserLoginT()
	dim strLogin
	If Session("UserName")="" Then
    	strLogin= "<table width='100%'border='0'cellspacing='0'cellpadding='0'>"
			strLogin=strLogin & "<form action='Bs_UserLogin.asp'method='post'name='UserLogin'onSubmit='return CheckForm();'><tr>"
			strLogin=strLogin & "<td height='25' align='right'>&nbsp;会员登陆∶&nbsp;用户&nbsp;</td>"
			strLogin=strLogin & "<td height='25'><input name='UserName'type='text'id='UserName'size='10'maxlength='20'></td>"
			strLogin=strLogin & "<td height='25' align='right'>&nbsp;密码&nbsp;</td>"
			strLogin=strLogin & "<td height='25'><input name='Password'type='password'id='Password'size='10'maxlength='20'></td>"
			strLogin=strLogin & "<td  width='8'> </td>"
			strLogin=strLogin & "<td height='25'><input type=image height=18 width=42 src='../Img/bot_login.gif'name=image></td>"
			strLogin=strLogin & "<td  width='10'> </td>"
			strLogin=strLogin & "<td height='25' align='center'><a href='Bs_UserReg.asp'><IMG src='../img/bot_reg.gif' width=42 height=18 border=0></a></td>"      
	'		strLogin=strLogin & "<td height='25' align='center'><a href='Bs_GetPassword.asp'target='_blank'>忘记密码</a></td>"      
			strLogin=strLogin & "</tr></form></table>"
		response.write strLogin
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("请输入用户名！");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("请输入密码！");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<%
	Else 
	if strSkins <= 150 and strSkins >= 1 then
		response.write "<center>欢迎您！<font color='#FF0000'><b>" & Session("UserName") & "</b></font>"
		response.write "<a href='Bs_Server.asp'><b>会员管理中心</b></a><a href='Bs_UserLogout.asp'><b>退出会员中心</b></a></center>"
	elseif strSkins <= 250 and strSkins >= 201 then
		response.write "<center>欢迎您！<font color='#FF0000'><b>" & Session("UserName") & "</b></font><br>"
		response.write "<br><a href='Bs_Server.asp'><b>会员管理中心</b></a><br><br><a href='Bs_UserLogout.asp'><b>退出会员中心</b></a><br><br></center>"
	end if
	end if
end sub
%>

<script language="javascript">
<!--
function winopen(url)
{
window.open(url,"search","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=yes,width=640,height=450,top=200,left=100");
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
window.open(theURL,winName,features);
}
//-->
</script>
<script>
function eshop(id) { window.open("Bs_Eshop.asp?cpbm="+id,"","height=400,width=640,left=200,top=0,resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");}
</script> 
