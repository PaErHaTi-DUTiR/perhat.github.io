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
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If Rens="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If Did="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If Daiy="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If Yaoq="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If Qix="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
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
		<td class="a1" height="25" align="center"><strong>发布招聘信息</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
					<form method="post" action="Bs_Job_Add.asp?no=eshop"><td><div align="center"> 
							<table width="100%" border="0" cellspacing="2" cellpadding="3">
								<tr> 
									<td width="19%" height="25" bgcolor="#C0C0C0"> <div align="center">招聘对象</div></td>
									<td width="81%" bgcolor="#E3E3E3">
									<input name="Duix" type="text" class="inputtext" id="Duix" size="25" maxlength="30"></td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"> 
										<div align="center">招聘人数</div></td>
									<td bgcolor="#E3E3E3"><input name="Rens" type="text" class="inputtext" id="Rens" size="5" maxlength="30">
										人</td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"> <div align="center">工作地点</div></td>
									<td height="-4" bgcolor="#E3E3E3">
									<input name="Did" type="text" class="inputtext" id="Did" size="30" maxlength="50"> 
									</td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"><div align="center">工资待遇</div></td>
									<td height="-1" bgcolor="#E3E3E3">
									<input name="Daiy" type="text" class="inputtext" id="Daiy" size="20" maxlength="50"></td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0">
									<div align="center">发布日期</div></td>
									<td height="-3" bgcolor="#E3E3E3"> <%=Date()%></td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"><div align="center">有效期限</div></td>
									<td height="0" bgcolor="#E3E3E3"><input name="Qix" type="text" class="inputtext" id="Qix" size="5" maxlength="30">
										天</td>
								</tr>
								<tr> 
									<td height="22" bgcolor="#C0C0C0"><div align="center">招聘要求</div></td>
									<td height="10" bgcolor="#E3E3E3">
									<textarea name="Yaoq" cols="40" rows="5" class="inputtext" id="Yaoq"></textarea></td>
								</tr>
								<tr> 
									<td height="25" colspan="2" bgcolor="#E3E3E3"> <div align="center"> 
											<input type="submit" value="确定">
											&nbsp; 
											<input type="reset" value="取消">
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
