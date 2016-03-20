<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<%
'=========================================================
'
'产品名称：科技 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有：www.web300.cn
'程序制作：QQ812256@hotmail.com
'联系 方式：QQ ：812256
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
		<td class="a1" height="25" align="center"><strong>回复留言</strong></td>
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
							  &nbsp;&nbsp;你正准备回复<b><%=rs("name")%></b>的留言，他的留言是：</td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
							  <td width="18%" height="22" bgcolor="#DCDDDE"> <div align="center">主 题</div></td>
								<td width="82%" bgcolor="#DCDDDE">&nbsp;&nbsp;<%=rs("title")%></td>
							</tr>
							<tr bgcolor="#C0C0C0"> 
							  <td height="25" bgcolor="#DCDDDE"><div align="center">内 容</div></td>
								<td height="25" bgcolor="#DCDDDE">&nbsp;&nbsp;<%=rs("content")%></td>
							</tr>
						</table>
					</div></td>
        </tr>
      </table>
      <br>
			<table width="500" border="0" cellspacing="0" cellpadding="5">
				<tr> 
					<td><div align="center">输入回复内容</div></td>
				</tr>
				<tr>
					<td><div align="center">
							<textarea name="rebook" cols="60" rows="10"><%=rs("rebook")%></textarea>
					</div></td>
				</tr>
				<tr>
					<td><div align="center">是否支持Html 
							<input type="checkbox" name="html" value="on">
							<input type="submit" value="确定">
							<input type="reset" value="从来">
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
