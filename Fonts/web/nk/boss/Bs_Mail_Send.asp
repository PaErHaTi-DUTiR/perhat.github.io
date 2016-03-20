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
		<td class="a1" height="25" align="center"><strong>发 送 邮 件</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
			<form name="sendmail" action="Bs_Mail_send_to.asp" method="post" >
        <table border="0" cellspacing="2" cellpadding="0" width="550" align="center">
          <tr> 
            <td height="25" bgcolor="#C0C0C0">
						<div align="center"><font color="#000000">发信人地址：</font></div></td>
            <td height="30" bgcolor="#E3E3E3"> <input type="text" name="frommail" value="<%=rs("frommail")%>"> 
            </td>
          </tr>
          <tr> <%email=request("email")%>
            <td bgcolor="#C0C0C0">
						<div align="center"><font color="#000000">收信人地址：</font></div></td>
            <td bgcolor="#E3E3E3"> 
              <input name="tomail" type="text" value="<%=email%>">
              <br>
              <br>
              <font color="#000000">（如果为空，从数据库取地址群发！） </font></td>
          </tr>
          <tr> 
            <td bgcolor="#C0C0C0">
						<div align="center"><font color="#000000">信件标题：</font></div></td>
            <td bgcolor="#E3E3E3"> 
              <input type="text" name="mailsubject" size="50">
              <br> <br>
              <font color="#000000">（如果为空，则显示<font color="#FF0000">“<%=rs("mailsubject")%>”</font>！） 
              </font></td>
          </tr>
          <tr> 
            <td bgcolor="#C0C0C0">
						<div align="center"><font color="#000000">信件内容：</font></div></td>
            <td bgcolor="#E3E3E3"> 
              <textarea name="mailbody" cols="50" rows="8"></textarea>
              <br>
              <br>
            </td>
          </tr>
          <tr bgcolor="#336699"> 
            <td height="25" bgcolor="#C0C0C0"> 
              <div align="center"></div></td>
            <td height="22" bgcolor="#E3E3E3"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="submit" name="Submit" value="发送"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="reset" name="Submit" value="取消"> <input type="hidden" name="action" value="我发"> 
            </td>
          </tr>
        </table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
