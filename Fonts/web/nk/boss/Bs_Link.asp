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
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = Conn
Rs.Open "SELECT * from Bs_Links Order BY id"

if Request.QueryString("no")="yes" then
id= Trim(Request("id"))
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = Conn
Rs.Open "DELETE from Bs_Links Where id="&id,Conn,2,3,1


Set Rs= Nothing
Set Conn = Nothing
Response.Redirect "Bs_Link.asp"
end if



if Request.QueryString("no")="eshop" then
note=request.form("note")
link=request.form("link")
name=request.form("name")

If name=""  or note="http://"Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入网站名称，请返回检查！！"");history.go(-1);</script>")
response.end
end if

If note=""  or note="http://"Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入图标连接，请返回检查！！"");history.go(-1);</script>")
response.end
end if

If link="" or link="http://" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没有输入超连接，请返回检查！！"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Links "
rs.open sql,conn,1,3
rs.addnew
rs("name")=name
rs("note")=note
rs("link")=link

rs.update
Response.Redirect "Bs_Link.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>友情链接管理</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <form method="post" action="Bs_Link.asp?no=eshop">
            <td>
							<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2">
                <tr> 
                  <td width="21%" height="22" bgcolor="#C0C0C0"> <div align="center">网站名称</div></td>
                  <td width="79%" bgcolor="#E3E3E3"><input type="text" name="name" size="35" maxlength="40"></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#C0C0C0"><div align="center">网站说明</div></td>
                  <td bgcolor="#E3E3E3"><input type="text" name="note" size="50" maxlength="120"></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#C0C0C0"> <div align="center">连接地址</div></td>
                  <td bgcolor="#E3E3E3"><input type="text" name="link" size="40" maxlength="50" value="http://"></td>
                </tr>
                <tr bgcolor="#C0C0C0"> 
                  <td height="22" colspan="2"> <div align="center"> 
                      <input type="submit" name="Submit" value="提交">
                      <input type="reset" name="Submit2" value="重置">
                    </div></td>
                </tr>
              </table>
              <br> 
              <table width="100%" border="0" cellspacing="2" cellpadding="3">
                <tr bgcolor="#C0C0C0"> 
                  <td width="22%" height="25" bgcolor="#C0C0C0"><div align="center">网站名称</div></td>
                  <td width="46%" bgcolor="#C0C0C0"><div align="center">网站说明</div></td>
                  <td width="14%" bgcolor="#C0C0C0"><div align="center">加入时间</div></td>
                  <td width="9%" bgcolor="#C0C0C0"><div align="center">操作</div></td>
                  <td width="9%" bgcolor="#C0C0C0"><div align="center">操作</div></td>
                </tr>
                <% do while not Rs.eof %>
                <tr bgcolor="#E3E3E3"> 
                  <td height="22" bgcolor="#E3E3E3"><div align="center"><a href="<%=Rs("link")%>" target="_blank"><%=rs("name")%></a></div></td>
                  <td bgcolor="#E3E3E3"><div align="center"><%=Rs("note")%></div></td>
                  <td bgcolor="#E3E3E3"><div align="center"><%= FormatDateTime(rs("time"),2) %></div></td>
                  <td bgcolor="#E3E3E3"><div align="center"><a href="Bs_link_edit.asp?id=<%=Rs("id")%>">修改</a></div></td>
                  <td bgcolor="#E3E3E3"><div align="center"><a href="Bs_Link.asp?id=<%=Rs("id")%>&amp;no=yes">删除</a></div></td>
                </tr>
<%
Rs.MoveNext 
loop 
%>
              </table>
<%
Set Rs = Nothing 
Set Conn = Nothing 
%>
            </td>
          </form>
        </tr>
        <tr>
          <td><BR>&nbsp;首页最多可显示10条友情链接</td>
        </tr>
      </table><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
