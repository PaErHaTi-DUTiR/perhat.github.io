<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<%
'=========================================================
'
'��Ʒ���ƣ��Ƽ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ ��ʽ��QQ ��812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%if Request.QueryString("no")="eshop" then
id=request("id")

if request.form("html")="on" then
rebook=request.form("rebook")
else
rebook=htmlencode2(request.form("rebook"))
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Book where id="&id
rs.open sql,conn,1,3

rs("rebook")=rebook
rs.update
rs.close
response.redirect "Bs_Book.asp"
end if

id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * from Bs_Book where id="&id, conn,3,3
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�ظ�����</strong></td>
	</tr>
	<tr class="a4">
    <form method="post" action="Bs_Book_Re.asp?no=eshop"><input type=hidden name=id value=<%=id%>>
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
					<td><div align="center"> 
						<table width="100%" border="0" cellspacing="2" cellpadding="3">
							<tr bgcolor="#C0C0C0"> 
								<td height="25" colspan="2" bgcolor="#DCDDDE"> <div align="center"></div>
							  &nbsp;&nbsp;����׼���ظ�<b><%=rs("name")%></b>�����ԣ����������ǣ�</td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
							  <td width="18%" height="22" bgcolor="#DCDDDE"> <div align="center">�� ��</div></td>
								<td width="82%" bgcolor="#DCDDDE">&nbsp;&nbsp;<%=rs("title")%></td>
							</tr>
							<tr bgcolor="#C0C0C0"> 
							  <td height="25" bgcolor="#DCDDDE"><div align="center">�� ��</div></td>
								<td height="25" bgcolor="#DCDDDE">&nbsp;&nbsp;<%=rs("content")%></td>
							</tr>
						</table>
					</div></td>
        </tr>
      </table>
      <br>
			<table width="500" border="0" cellspacing="0" cellpadding="5">
				<tr> 
					<td><div align="center">����ظ�����</div></td>
				</tr>
				<tr>
					<td><div align="center">
							<textarea name="rebook" cols="60" rows="10"><%=rs("rebook")%></textarea>
					</div></td>
				</tr>
				<tr>
					<td><div align="center">�Ƿ�֧��Html 
							<input type="checkbox" name="html" value="on">
							<input type="submit" value="ȷ��">
							<input type="reset" value="����">
					</div></td>
        </tr>
      </table>
      <br>
		</td>
		</form>
	</tr>
</table>
<BR>
<%htmlend%>
