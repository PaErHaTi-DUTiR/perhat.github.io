<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<%
'=========================================================
'
'��Ʒ���ƣ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ��ʽ��QQ 812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%
if Request.QueryString("no")="eshop" then
If request.form("title")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������⣬�뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if


If request.form("content")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����������ݣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Book"
rs.open sql,conn,1,3


rs.addnew
if request.form("html")="on" then
rs("content")=request.form("content")
else
rs("content")=htmlencode2(request.form("content"))
end if
rs("name")=request.form("name")
rs("title")=request.form("title")
rs("time")=date()
rs.update
rs.close
response.redirect "Bs_Book.asp"
end if

%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>��������Ա����</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <form method="POST" action="Bs_Book_Add.asp?no=eshop">
						<input type=hidden readonly name="Name"   value="����Ա">
						<td><div align="center"> 
							<table width="80%" border="0" cellspacing="2" cellpadding="3">
								<tr bgcolor="#C0C0C0"> 
								  <td width="8%" height="25" bgcolor="#DCDDDE"> <div align="center">��&nbsp;&nbsp;��</div></td>
								  <td width="62%" bgcolor="#DCDDDE"><input type="text" name="title" size="45"></td>
								</tr>
								<tr bgcolor="#E3E3E3"> 
									<td height="22" bgcolor="#E3E3E3"> <div align="center">��&nbsp;&nbsp;��</div></td>
									<td bgcolor="#E3E3E3"><textarea rows="8" name="Content" cols="45"></textarea></td>
								</tr>
								<tr bgcolor="#C0C0C0"> 
									<td height="25" colspan="2" bgcolor="#DCDDDE"> 
								  <div align="center">�Ƿ�֧��html <input type="checkbox" name="html" value="on"> <input type="submit" value="�ύ����" name="B1"> <input type="reset" value="ȫ����д" name="B2"></div></td>
								</tr>
							</table>
						</div></td>
					</form>
        </tr>
      </table>
      <br>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
