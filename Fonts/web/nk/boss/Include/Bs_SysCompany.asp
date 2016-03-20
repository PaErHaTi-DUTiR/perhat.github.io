<%
'=========================================================
'
'
'产品名称：科技 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有：www.web300.cn
'程序制作：QQ812256@hotmail.com
'联系 方式：QQ ：812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'
'========================================================
%>
<%content=server.htmlencode(Trim(Request("content"))) %>
<%EnContent=server.htmlencode(Trim(Request("EnContent"))) %>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("adodb.recordset")
sql="select * from Bs_Company"
rs.open sql,conn,3,3
rs(strContent)=content
rs(strEnContent)=EnContent
rs.update
rs.close
response.redirect "Bs_Company.asp?UrlName="& UrlName &""
end if
sql="select * from Bs_Company"
Set rs2= Server.CreateObject("ADODB.Recordset")
rs2.open sql,conn,1,1
%>
<!-- #include file="../Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong><%=strTitleName%></strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <form method="POST" action="Bs_Company.asp?UrlName=<%=UrlName%>&no=eshop">
			<table width="550" border="0" cellspacing="2" cellpadding="0">
				<tr> 
				<td width="19%" height="300" bgcolor="#C0C0C0"> <div align="center">中文信息设置<BR><BR>支持UBB代码</div></td>
				<td width="81%" bgcolor="#E3E3E3">
<script src='Inc/Eshopcode.js'></script> 
<!--#include file="../Inc/Ubb.inc" -->
<textarea name="content" cols="58" rows="15">
<%
if rs2("html")=false then
content=replace(rs2(strContent),"<br>",chr(13))
content=replace(content,"&nbsp;"," ")
else
content=rs2(strContent)
end if
response.write content
%>
</textarea></td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr> 
				<td width="19%" height="300" bgcolor="#C0C0C0"> <div align="center">English信息设置<BR><BR>支持UBB代码</div></td>
				<td width="81%" bgcolor="#E3E3E3">
<textarea name="EnContent" cols="58" rows="15">
<%
if rs2("html")=false then
EnContent=replace(rs2(strEnContent),"<br>",chr(13))
EnContent=replace(EnContent,"&nbsp;"," ")
else
EnContent=rs2(strEnContent)
end if
response.write EnContent
%>
</textarea>
<BR>&nbsp;英文栏的UBB代码只能直接输入.
<BR>&nbsp;如:图片连接[img]http://www.xxx.com/xxx.gif[/img],[b]文字加粗体效果[/b].
<BR>&nbsp;或从中文栏中复制UBB代码,或查看系统帮助.</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2"><div align="center"
><input type="submit" value=" 修 改 "	 name="cmdok">&nbsp;<input type="reset" value=" 重 写 " name="cmdcancel"></div></td>
				</tr>
				<tr>
					<td colspan="2">图片上传</td>
				</tr>
				<tr>
					<td colspan="2"><div align="center"
><iframe name="ad" frameborder=0 width=100% height=48 scrolling=no src='../WebEdit/Upload_File.asp'></iframe></div></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
