<%@language=vbscript codepage=936 %>
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
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'��Ҫ��ʹ������ֵ�ͼƬ�������
%>
<html>
<head>
<title>����Ա��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
BODY
{
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	text-decoration: none;
	line-height: 150%;
	background-color: #FBFDFF;
	FONT-SIZE: 9pt;background:#ffffff;
text-decoration: none;
SCROLLBAR-FACE-COLOR: #cccccc;
SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #333333; SCROLLBAR-3DLIGHT-COLOR: #333333; SCROLLBAR-ARROW-COLOR: #cccccc; SCROLLBAR-TRACK-COLOR: #cccccc; SCROLLBAR-DARKSHADOW-COLOR: #ffffff;
}
TD{	FONT-FAMILY: "����";	FONT-SIZE: 9pt;}
Input{	FONT-SIZE: 9pt;	HEIGHT: 20px;}
Button{	FONT-SIZE: 9pt;	HEIGHT: 20px; }
Select{	FONT-SIZE: 9pt;	HEIGHT: 20px;}
A{	TEXT-DECORATION: none;	color: #333333;}
A:hover{	COLOR: #428EFF;	text-decoration: underline;}
.title{	background:url(Images/topBar_bg.gif);}
.border{	border: 1px solid #39867B;}
.tdbg{	background:#E1F4EE;	line-height: 120%;}
.topbg{	background:url(Images/topbg.gif);	color: #FFFFFF;}
.bgcolor {	background-color: #ffffff;}
.STYLE1 {font-family: Arial, Helvetica, sans-serif}
-->
</style>
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("�������û�����");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("���������룡");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("������������֤�룡");
       document.Login.CheckCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("������ʾ��\n    ��ʹ�õ���Netscape����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("������ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  }
}
//-->
</script>
</head>
<body class="bgcolor">
<p>&nbsp;</p>
<center>
<table border=1 borderColor=#666666 cellPadding=4 cellSpacing=1 width=500 style="border-collapse: collapse" align=center >
	<tr valign="baseline"> 
		<td align="right" background=images/b1.gif>
        <div align="center"><font color="#333333" style="font-size: 10.5pt;"> 
          <strong>��վ��̨����ϵͳ��¼</strong></font></div>
      </td>
	</tr>
	<tr valign="baseline"> 
		<td bgcolor=#EFF1F3 align=center valign=middle height=60> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width=20> </td>
					
            <td width=150><div align="center"><img src="images/setup.gif" width="184" height="130" border="0" usemap="#Map"> 
            </div></td>
					<td width=280>
<form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
	<table width="100%" border="0" cellspacing="8" cellpadding="0" align="center">
		<tr align="center"> 
			<td height="38" colspan="2"><font color="#333333" size="3"><strong>����Ա��¼</strong></font> 
			</td>
		</tr>
		<tr> 
			<td align="right"><font color="#333333">�û����ƣ�</font></td>
			<td><input name="UserName"  type="text"  id="UserName4" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onFocus="this.select(); " onMouseOver="this.style.background='#cccccc';" onMouseOut="this.style.background='#ffffff'" maxlength="20"></td>
		</tr>
		<tr> 
			<td align="right"><font color="#333333">�û����룺</font></td>
			<td><input name="Password"  type="password" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onFocus="this.select(); " onMouseOver="this.style.background='#cccccc';" onMouseOut="this.style.background='#ffffff'" maxlength="20"></td>
		</tr>
		<tr> 
			<td align="right"><font color="#00000">�� ֤ �룺</font></td>
			<td><input name="CheckCode" size="6" maxlength="4" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#cccccc';" onMouseOut="this.style.background='#ffffff'" onFocus="this.select(); ">
				<font color="#333333">�����������</font>            <img src="inc/checkcode.asp"></td>
		</tr>
		<tr> 
			<td colspan="2"> <div align="center"> &nbsp;&nbsp;&nbsp;&nbsp;
				<input   type="submit" name="Submit" value=" ȷ&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #333333; background-color: #cccccc; border: 1 solid #333333" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#cccccc'">
				&nbsp; 
				<input name="reset" type="reset"  id="reset" value=" ��&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #333333; background-color: #cccccc; border: 1 solid #333333" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#cccccc'"><br>
				</div></td>
		</tr>
	</table>
</form>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</center>
<p align="center"><span class="STYLE1">ASBDcorpWEB Mac3.0 For ASBD.Co.,LTD Copyright(C):<a href="http://BBS.ASBD.CN">BBS.ASBD.CN</a><br>
Powered By��<a href="http://www.asbd.cn">ASPIRATION SCHEME &amp; BRAND DESIGN Co.,LTD</a></span><br>
  <br>
</p>

<script language="JavaScript" type="text/JavaScript">
CheckBrowser();
SetFocus(); 
</script>

</body>
</html>
