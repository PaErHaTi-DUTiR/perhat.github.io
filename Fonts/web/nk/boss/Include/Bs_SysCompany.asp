<%
'=========================================================
'
'
'��Ʒ���ƣ��Ƽ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ ��ʽ��QQ ��812256
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
				<td width="19%" height="300" bgcolor="#C0C0C0"> <div align="center">������Ϣ����<BR><BR>֧��UBB����</div></td>
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
				<td width="19%" height="300" bgcolor="#C0C0C0"> <div align="center">English��Ϣ����<BR><BR>֧��UBB����</div></td>
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
<BR>&nbsp;Ӣ������UBB����ֻ��ֱ������.
<BR>&nbsp;��:ͼƬ����[img]http://www.xxx.com/xxx.gif[/img],[b]���ּӴ���Ч��[/b].
<BR>&nbsp;����������и���UBB����,��鿴ϵͳ����.</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2"><div align="center"
><input type="submit" value=" �� �� "	 name="cmdok">&nbsp;<input type="reset" value=" �� д " name="cmdcancel"></div></td>
				</tr>
				<tr>
					<td colspan="2">ͼƬ�ϴ�</td>
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
