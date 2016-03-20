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
title=server.htmlencode(Trim(Request("title")))
Entitle=server.htmlencode(Trim(Request("Entitle")))
img=server.htmlencode(Trim(Request("img")))
%>
<%


If title="" Then
response.write "SORRY <br>"
response.write "请输入更新内容的标题!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If img="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Img "
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("Entitle")=Entitle
rs("img")=img
rs.update
rs.close
response.redirect "Bs_Img.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>添加形象</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <form method="post" name="myform" action="Bs_Img_add.asp?no=eshop">
            <td><div align="center"> 
							<table width="100%" border="0" cellspacing="2" cellpadding="3">
								<tr> 
									<td width="18%" bgcolor="#C0C0C0"><div align="right">形象名称</div></td>
									<td width="62%" bgcolor="#C0C0C0"><input type="text" name="title" class="inputtext" maxlength="300000" size="25"></td>
								</tr>
								<tr> 
									<td width="18%" bgcolor="#C0C0C0"><div align="right">English名称</div></td>
									<td width="62%" bgcolor="#C0C0C0"><input type="text" name="Entitle" class="inputtext" maxlength="300000" size="25"></td>
								</tr>
								<tr> 
									<td bgcolor="#E3E3E3"><div align="right">形象图片</div></td>
									<td bgcolor="#E3E3E3"><input name="img" type="text" class="inputtext" id="img" size="30" maxlength="500000">
										<font color="#FF0000"> * 图片地址</font></td>
								</tr>
								<tr bgcolor="#C0C0C0"> 
									<td colspan="2"> 
										<div align="center"> 
											<input name="submit" type="submit" value="确定">
											&nbsp; 
											<input name="reset" type="reset" value="取消">
										</div></td>
								</tr>
								<tr> 
									<td colspan="2">图片上传</td>
								</tr>
								<tr> 
									<td colspan="2"> 
										<div align="center"> 
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
