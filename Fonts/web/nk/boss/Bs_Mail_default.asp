<!--#include file="bsconfig.asp"-->
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
<% 
set rs=server.createobject("adodb.recordset") 
sql="select * from maildefault order by id desc" 
rs.open sql,conn,1,1   
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� �� �� ��</strong></td>
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
								<td height="25" bgcolor="#C0C0C0"><b><font color="#000000">&nbsp;�����ʼ���Ĭ�Ϸ�����</font></b></td>
							</tr>
							<tr> 
								<td bgcolor="#E3E3E3" align="center"> 
									<table border="0" cellspacing="0" cellpadding="4">
										<tr> 
											<td><font color="#000000">������:</font> </td>
										</tr>
										<tr> 
											<td align="right"><input name="frommail" type="text" value="<%=rs("frommail")%>" size="25" maxlength="50"></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr> 
								<td height="30" align="center" bgcolor="#C0C0C0"> 
									<input type="submit" name="Submit" value="����" class="Tips_bo"> 
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
									<p><b><font color="#000000">&nbsp;�����ʼ�Ĭ�Ϸ�������</font></b></p></td>
							</tr>
							<tr> 
								<td bgcolor="#E3E3E3" align="center"> 
									<textarea name="mailbody" cols="46" rows="12"><%=rs("mailbody")%></textarea>
								</td>
							</tr>
							<tr> 
								<td bgcolor="#C0C0C0" align="center" height="30"> 
									<input type="submit" name="Submit2" value="����" class="Tips_bo"> 
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
									<p><b><font color="#000000">&nbsp;�����ʼ�Ĭ�ϱ���</font></b></p></td>
							</tr>
							<tr> 
								<td align="center" bgcolor="#E3E3E3"> 
									<table border="0" cellspacing="0" cellpadding="4">
										<tr> 
											<td><font color="#000000">�ʼ�����:</font> </td>
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
									<input type="submit" name="Submit22" value="����" class="Tips_bo"> 
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
'��ӹ��ܺ��޸Ĺ���

if request("action")<>"frommail" and request("action")<>"mailsubject" and request("action")<>"mailbody" then 
response.end
else

if request("action")="frommail" then
frommail=request("frommail")
conn.execute "update maildefault set frommail='"&frommail&"'"
response.write"<SCRIPT language=JavaScript>alert('�����˵�ַ���ĳɹ���');"
response.write "</SCRIPT>"
end if

if request("action")="mailsubject" then
mailsubject=request("mailsubject")
conn.execute "update maildefault set mailsubject='"&mailsubject&"'"
response.write"<SCRIPT language=JavaScript>alert('�ʼ�Ĭ�ϱ�����ĳɹ���');"
response.write "</SCRIPT>"
end if

if request("action")="mailbody" then
mailbody=request("mailbody")
conn.execute "update maildefault set mailbody='"&mailbody&"'"
response.write"<SCRIPT language=JavaScript>alert('�ʼ�Ĭ�����ݸ��ĳɹ���');"
response.write "</SCRIPT>"
end if

end if
%> 
