<!--#include file="bsconfig.asp"-->
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
set rs=server.createobject("adodb.recordset") 
sql="select * from maildefault order by id desc" 
rs.open sql,conn,1,1   
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>邮 件 列 表 设 置</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td width="36%">
					<form name="addcat" method="post" action="Bs_Mail_default.asp">
						<table width="98%" border="0" align="center" cellpadding="0" cellspacing="2">
							<tr> 
								<td height="25" bgcolor="#C0C0C0"><b><font color="#000000">&nbsp;设置邮件的默认发信人</font></b></td>
							</tr>
							<tr> 
								<td bgcolor="#E3E3E3" align="center"> 
									<table border="0" cellspacing="0" cellpadding="4">
										<tr> 
											<td><font color="#000000">发信人:</font> </td>
										</tr>
										<tr> 
											<td align="right"><input name="frommail" type="text" value="<%=rs("frommail")%>" size="25" maxlength="50"></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr> 
								<td height="30" align="center" bgcolor="#C0C0C0"> 
									<input type="submit" name="Submit" value="设置" class="Tips_bo"> 
									<input type="hidden" name="action" value="frommail">
								</td>
							</tr>
						</table>
					</form>
					</td>
          <td rowspan="2" width="64%">
					<form name="form1" method="post" action="Bs_Mail_default.asp">
						<table border="0" cellspacing="2" cellpadding="0" width="98%">
							<tr> 
								<td height="30" bgcolor="#C0C0C0"> 
									<p><b><font color="#000000">&nbsp;设置邮件默认发送内容</font></b></p></td>
							</tr>
							<tr> 
								<td bgcolor="#E3E3E3" align="center"> 
									<textarea name="mailbody" cols="46" rows="12"><%=rs("mailbody")%></textarea>
								</td>
							</tr>
							<tr> 
								<td bgcolor="#C0C0C0" align="center" height="30"> 
									<input type="submit" name="Submit2" value="设置" class="Tips_bo"> 
									<input type="hidden" name="action" value="mailbody">
								</td>
							</tr>
						</table>
						<div align="center"></div>
					</form>
					</td>
        </tr>
        <tr> 
          <td width="36%">
					<form name="form1" method="post" action="Bs_Mail_default.asp">
						<table width="98%" border="0" align="center" cellpadding="0" cellspacing="2">
							<tr> 
								<td height="25" bgcolor="#C0C0C0"> 
									<p><b><font color="#000000">&nbsp;设置邮件默认标题</font></b></p></td>
							</tr>
							<tr> 
								<td align="center" bgcolor="#E3E3E3"> 
									<table border="0" cellspacing="0" cellpadding="4">
										<tr> 
											<td><font color="#000000">邮件标题:</font> </td>
										</tr>
										<tr> 
											<td align="right"><div align="center"> 
													<input name="mailsubject" type="text" value="<%=rs("mailsubject")%>" size="25" maxlength="50">
												</div></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr> 
								<td height="30" align="center" bgcolor="#C0C0C0"> 
									<input type="submit" name="Submit22" value="设置" class="Tips_bo"> 
									<input type="hidden" name="action" value="mailsubject">
								</td>
							</tr>
						</table>
					</form>
					</td>
        </tr>
      </table>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
'添加功能和修改功能

if request("action")<>"frommail" and request("action")<>"mailsubject" and request("action")<>"mailbody" then 
response.end
else

if request("action")="frommail" then
frommail=request("frommail")
conn.execute "update maildefault set frommail='"&frommail&"'"
response.write"<SCRIPT language=JavaScript>alert('发信人地址更改成功！');"
response.write "</SCRIPT>"
end if

if request("action")="mailsubject" then
mailsubject=request("mailsubject")
conn.execute "update maildefault set mailsubject='"&mailsubject&"'"
response.write"<SCRIPT language=JavaScript>alert('邮件默认标题更改成功！');"
response.write "</SCRIPT>"
end if

if request("action")="mailbody" then
mailbody=request("mailbody")
conn.execute "update maildefault set mailbody='"&mailbody&"'"
response.write"<SCRIPT language=JavaScript>alert('邮件默认内容更改成功！');"
response.write "</SCRIPT>"
end if

end if
%> 
