<!--#include file="bsconfig.asp"-->
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
<%if Request.QueryString("no")="eshop" then

id=request("id")
name=request("name")
note=request("note")
link=request("link")


If name="" Then
response.write "SORRY <br>"
response.write "请输入更新内容"
response.end
end if
If note="" Then
response.write "SORRY <br>"
response.write "内容不能为空"
response.end
end if
If link="" Then
response.write "SORRY <br>"
response.write "内容不能为空"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Links where id="&id
rs.open sql,conn,1,3

rs("name")=name
rs("note")=note
rs("link")=link
rs.update
rs.close
response.redirect "Bs_Link.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * from Bs_Links where id="&id, conn,3,3
%>

<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>修改友情链接</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <form method="post" action="Bs_Link_edit.asp?no=eshop">
            <input type=hidden name=id value=<%=id%>>
            <td><div align="center">
							<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2">
                <tr> 
                  <td width="21%" height="22" bgcolor="#C0C0C0"> <div align="center">网站名称</div></td>
                  <td width="79%" bgcolor="#E3E3E3"><input type="text" name="name" class="inputtext" id="name" value="<%=rs("name")%>" size="35" maxlength="40"></td>
               </tr>
                <tr> 
                  <td height="22" bgcolor="#C0C0C0"><div align="center">网站说明</div></td>
                  <td bgcolor="#E3E3E3"><input type="text" name="note" class="inputtext" id="note" value="<%=rs("note")%>" size="50" maxlength="120"></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#C0C0C0"> <div align="center">连接地址</div></td>
                  <td bgcolor="#E3E3E3"><input type="text" name="link" class="inputtext" id="link" value="<%=rs("link")%>" size="40" maxlength="50"></td>
                </tr>
                <tr bgcolor="#C0C0C0"> 
                  <td height="22" colspan="2"> <div align="center"> 
											<input name="submit" type="submit" value="确定">
											&nbsp; 
											<input name="reset" type="reset" value="取消">
									</div></td>
                </tr>
              </table>
            </td>
          </form>
        </tr>
      </table><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
