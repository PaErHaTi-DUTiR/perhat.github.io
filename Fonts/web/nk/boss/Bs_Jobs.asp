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
<%content=server.htmlencode(Trim(Request("content"))) %>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("adodb.recordset")
sql="select * from Bs_Company"
rs.open sql,conn,3,3
rs("Job")=content
rs.update
rs.close
response.redirect "Bs_Jobs.asp"
end if
sql="select * from Bs_Company"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>人 才 策 略 设 置</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <form method="POST" action="Bs_Jobs.asp?no=eshop">
					<td>
						<table width="100%" border="0" cellspacing="2" cellpadding="0">
							<tr> 
								<td width="19%" height="300" bgcolor="#C0C0C0"> <div align="center"><br> 支持UBB代码 </div></td>
								<td width="81%" bgcolor="#E3E3E3"> <script src=Inc/Eshopcode.js></script> 
<!--#include file=Inc/Ubb.inc -->
<textarea name="content" cols="58" rows="15">
<%
if rs_home("html")=false then
content=replace(rs_home("Job"),"<br>",chr(13))
content=replace(content,"&nbsp;"," ")
else
content=rs_home("Job")
end if
response.write content
%>
</textarea>
								</td>
							</tr>
							<tr> 
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr> 
								<td colspan="2"><div align="center"> 
									<input type="submit" value=" 修 改 " name="cmdok">
									&nbsp; 
									<input type="reset" value=" 重 写 " name="cmdcancel">
								</div></td>
							</tr>
							<tr>
								<td colspan="2">图片上传</td>
							</tr>
							<tr>
								<td colspan="2"><div align="center">
								<iframe name="ad" frameborder=0 width=100% height=48 scrolling=no src='../WebEdit/Upload_File.asp'></iframe></div></td>
							</tr>
						</table>
					</td>
					</form>
        </tr>
      </table>
			<br>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
