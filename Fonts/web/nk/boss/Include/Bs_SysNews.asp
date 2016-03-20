<%
if Request.QueryString("no")="eshop" then

CurrentPage = request("Page")
contentID=request("ID")

set rs=server.createobject("adodb.recordset")
sqltext="delete from "& strDataName &" where Id="& contentID
rs.open sqltext,conn,3,3
set rs=nothing
conn.close
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('成功删除！');" & Chr(13)		
		response.write "</script>" & Chr(13)
		response.redirect "Bs_News.asp?UrlName="& UrlName &""
end if
%>
<!-- #include file="../Inc/Head.asp" -->
<%
'call showpages
Call list()
%>
<%htmlend%>
<%
Sub list()

set rs=server.createobject("adodb.recordset")
sqltext="select * from "& strDataName &" order by id desc"
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=20

dim text,checkpage
text="0123456789"
Rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if
%>
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
<td><div align="center"> 
<table width="100%" border="0" cellspacing="2" cellpadding="3">
	<tr bgcolor="#C0C0C0"> 
		<td width="8%" height="25" bgcolor="#C0C0C0"> <div align="center">ID</div></td>
		<td width="62%"> <div align="center">新闻标题</div></td>
		<td width="10%"> <div align="center">操作</div></td>
		<td width="10%"> <div align="center">操作</div></td>
		<td width="10%"> <div align="center">操作</div></td>
	</tr>
<%
if not rs.eof then
i=0
do while not rs.eof
%>
	<tr bgcolor="#E3E3E3"> 
		<td height="22"> <div align="center"><%=rs("id")%></div></td>
		<td>&nbsp;&nbsp;<%=rs("title")%></td>
		<td> <div align="center"><a href='../Chinese/<%=strSeeUrl%>Id=<%=rs("id")%>'target="_blank">查看</a></div></td>
		<td> <div align="center"><a href="Bs_News_edit.asp?UrlName=<%=UrlName%>&id=<%=rs("id")%>">修改</a></div></td>
		<td> <div align="center"><a href="Bs_News.asp?UrlName=<%=UrlName%>&id=<%=rs("id")%>&amp;no=eshop">删除</a></div></td>
	</tr>
<% 
i=i+1 
if i >= MaxPerpage then exit do 
rs.movenext 
loop 
end if 
%>
	<tr bgcolor="#C0C0C0"> 
		<td height="25" colspan="5">&nbsp;&nbsp; 
<%
Response.write "<strong><font color='#000000'>-> 全部-</font>"
Response.write "共</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>条新闻</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>第</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='Bs_News.asp?UrlName="& UrlName &"&page="+cstr(1)+"'><font color='#000000'>首页</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='Bs_News.asp?UrlName="& UrlName &"&page="+Cstr(currentpage-1)+"'><font color='#000000'>上一页</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>上一页 </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='Bs_News.asp?UrlName="& UrlName &"&page="+Cstr(currentPage+1)+"'><font color='#000000'>下一页</font></a><font color='#ffffff'> </font>"
Response.write "<a href='Bs_News.asp?UrlName="& UrlName &"&page="+Cstr(Rs.PageCount)+"'><font color='#000000'>尾页</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>下一页</font></strong>&nbsp;&nbsp;"
End if
%>
		</td>
	</tr>
</table>
</div></td>
</tr>
</table>
      <br>
		</td>
	</tr>
</table>
<BR>
<%
'If Rs.recordcount > MaxPerPage then
'call showpages
'end if
rs.close
%>
<%End sub%>
