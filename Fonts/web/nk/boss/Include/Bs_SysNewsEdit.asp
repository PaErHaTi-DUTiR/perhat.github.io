<%if Request.QueryString("no")="eshop" then

id=request("id")
title=request("title")
content=htmlencode2(Request("content"))

If title="" Then
response.write "SORRY <br>"
response.write "请输入更新内容的标题"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "内容不能为空"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from "& strDataName &" where id="&id
rs.open sql,conn,1,3

rs("title")=title
rs("content")=content
rs.update
rs.close
response.redirect "Bs_News.asp?UrlName="& UrlName &""
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * from "& strDataName &" where id="&id, conn,3,3
%>
<!-- #include file="../Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong><%=strTitleName%></strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
					<form method="post" action="Bs_News_edit.asp?UrlName=<%=UrlName%>&no=eshop">
						<input type=hidden name=id value=<%=id%>>
						<td><div align="center"> 
<table width="100%" border="0" cellspacing="2" cellpadding="3">
	<tr> 
		<td width="8%" height="25" bgcolor="#C0C0C0"> <div align="center">标&nbsp;&nbsp;题</div></td>
		<td width="62%" bgcolor="#C0C0C0"><input type="text" name="title" maxlength="80" size="50" value="<%=rs("title")%>"></td>
	</tr>
	<tr> 
		<td height="22" bgcolor="#E3E3E3"> 
			<div align="center">UBB代码</div></td>
		<td bgcolor="#E3E3E3"><script src='Inc/eshopcode.js'></script> 
			<!--#include file=../Inc/Ubb.inc -->
		</td>
	</tr>
	<tr> 
		<td height="25" bgcolor="#C0C0C0"><div align="center">内&nbsp;&nbsp;容</div></td>
		<td height="25" bgcolor="#C0C0C0">
<textarea name="content" cols="57" rows="12">
<%if rs("html")=false then
content=replace(rs("content"),"<br>",chr(13))
content=replace(content,"&nbsp;"," ")
else
content=rs("content")
end if
response.write content%></textarea></td>
	</tr>
	<tr> 
		<td height="25" colspan="2" bgcolor="#E3E3E3"> 
			<div align="center"> 
				<input type="submit" value="确定">
				&nbsp;
				<input type="reset" value="从来">
			</div></td>
	</tr>
	<tr> 
		<td colspan="2">图片上传</td>
	</tr>
	<tr> 
		<td colspan="2"><div align="center"> 
				<iframe name="ad" frameborder=0 width=100% height=48 scrolling=no src='../WebEdit/Upload_File.asp'></iframe>
			</div></td>
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
