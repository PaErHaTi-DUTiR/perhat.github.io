<!--#include file="bsconfig.asp"-->
<%
'=========================================================
'
'��Ʒ���ƣ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ����: www.web300.cn 
'����������www.web300.cn�����Ŷ�
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%if Request.QueryString("no")="eshop" then%>
<%
Duix=server.htmlencode(Trim(Request("Duix")))
Rens=server.htmlencode(Trim(Request("Rens")))
Did=server.htmlencode(Trim(Request("Did")))
Daiy=server.htmlencode(Trim(Request("Daiy")))
Yaoq=server.htmlencode(Trim(Request("Yaoq")))
Qix=server.htmlencode(Trim(Request("Qix")))
%>
<%
If Duix="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If Rens="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If Did="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If Daiy="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If Yaoq="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If Qix="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Job "
rs.open sql,conn,1,3
rs.addnew
rs("Duix")=Duix
rs("Rens")=Rens
rs("Did")=Did
rs("Daiy")=Daiy
rs("Yaoq")=Yaoq
rs("Qix")=Qix
rs.update
rs.close
response.redirect "Bs_Job.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>������Ƹ��Ϣ</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
					<form method="post" action="Bs_Job_Add.asp?no=eshop"><td><div align="center"> 
							<table width="100%" border="0" cellspacing="2" cellpadding="3">
								<tr> 
									<td width="19%" height="25" bgcolor="#C0C0C0"> <div align="center">��Ƹ����</div></td>
									<td width="81%" bgcolor="#E3E3E3">
									<input name="Duix" type="text" class="inputtext" id="Duix" size="25" maxlength="30"></td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"> 
										<div align="center">��Ƹ����</div></td>
									<td bgcolor="#E3E3E3"><input name="Rens" type="text" class="inputtext" id="Rens" size="5" maxlength="30">
										��</td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"> <div align="center">�����ص�</div></td>
									<td height="-4" bgcolor="#E3E3E3">
									<input name="Did" type="text" class="inputtext" id="Did" size="30" maxlength="50"> 
									</td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"><div align="center">���ʴ���</div></td>
									<td height="-1" bgcolor="#E3E3E3">
									<input name="Daiy" type="text" class="inputtext" id="Daiy" size="20" maxlength="50"></td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0">
									<div align="center">��������</div></td>
									<td height="-3" bgcolor="#E3E3E3"> <%=Date()%></td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"><div align="center">��Ч����</div></td>
									<td height="0" bgcolor="#E3E3E3"><input name="Qix" type="text" class="inputtext" id="Qix" size="5" maxlength="30">
										��</td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"><div align="center">��ƸҪ��</div></td>
									<td height="10" bgcolor="#E3E3E3">
									<textarea name="Yaoq" cols="40" rows="5" class="inputtext" id="Yaoq"></textarea></td>
								</tr>
								<tr> 
									<td height="25" colspan="2" bgcolor="#E3E3E3"> <div align="center"> 
											<input type="submit" value="ȷ��">
											&nbsp; 
											<input type="reset" value="ȡ��">
									</div></td>
								</tr>
							</table>
					</div></td></form>
        </tr>
      </table>
      <br>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
