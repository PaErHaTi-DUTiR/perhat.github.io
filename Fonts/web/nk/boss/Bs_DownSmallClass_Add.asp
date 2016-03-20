<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->

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
dim Action,Bs_BigClassName,Bs_SmallClassName,sqlBig,rsBig,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
Bs_BigClassName=trim(request("Bs_BigClassName"))
Bs_SmallClassName=trim(request("Bs_SmallClassName"))

sqlBig="select * from Bs_DownBigClass where Bs_BigClassName='" & Bs_BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
rsBig.close

if Action="Add" then
	if Bs_BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>下载大类名不能为空！</li>"
	end if
	if Bs_SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>下载小类名不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From Bs_DownSmallClass Where Bs_BigClassName='" & Bs_BigClassName & "'AND Bs_SmallClassName='" & Bs_SmallClassName & "'",conn,1,3
		if not rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>“" & Bs_BigClassName & "”中已经存在下载小类“" & Bs_SmallClassName & "”！</li>"
		else
     		rs.addnew
				rs("Bs_BigClassName")=Bs_BigClassName
    	 	rs("Bs_SmallClassName")=Bs_SmallClassName
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "Bs_DownClass.asp"  
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<script language="JavaScript" type="text/JavaScript">
function checkSmall()
{
  if (document.form2.Bs_BigClassName.value=="")
  {
    alert("请先添加大类名称！");
	document.form1.Bs_BigClassName.focus();
	return false;
  }

  if (document.form2.Bs_SmallClassName.value=="")
  {
    alert("小类名称不能为空！");
	document.form2.Bs_SmallClassName.focus();
	return false;
  }
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="360" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>下 载 类 别 设 置</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<BR>
			<form name="form2" method="post" action="Bs_DownSmallClass_Add.asp" onsubmit="return checkSmall()">
			<table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
				<tr bgcolor="#999999" class="title"> 
					<td height="25" colspan="2" align="center"><strong>添加小类</strong></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>所属大类：</strong></td>
					<td width="218" bgcolor="#E3E3E3"> 
<select name="Bs_BigClassName">
<%
dim rsBigClass
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * From Bs_DownBigClass",conn,1,1
if yzcv<>zcv then response.end
if rsBigClass.bof and rsBigClass.bof then
	response.write "<option>请先添加文章大类</option>"
else
	do while not rsBigClass.eof
		if rsBigClass("Bs_BigClassName")=Bs_BigClassName then
			response.write "<option value='"& rsBigClass("Bs_BigClassName") & "'selected>" & rsBigClass("Bs_BigClassName") & "</option>"
		else
			response.write "<option value='"& rsBigClass("Bs_BigClassName") & "'>" & rsBigClass("Bs_BigClassName") & "</option>"
		end if
		rsBigClass.movenext
	loop
end if
rsBigClass.close
set rsBigClass=nothing
%>
</select>
					</td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>小类名称：</strong></td>
					<td bgcolor="#E3E3E3"><input name="Bs_SmallClassName" type="text" size="20" maxlength="30"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" align="center" bgcolor="#C0C0C0">&nbsp; </td>
					<td height="22" align="center" bgcolor="#E3E3E3"><div align="left"
><input name="Action" type="hidden" id="Action3" value="Add"><input name="Add" type="submit" value=" 添 加 "></div></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
end if
call CloseConn()
%>
