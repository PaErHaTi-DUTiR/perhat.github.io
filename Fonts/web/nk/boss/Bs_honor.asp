<!--#include file="bsconfig.asp"-->
<%
'=========================================================
'
'��Ʒ���ƣ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ����: www.web300.cn 
'����������www.web300.cn�����Ŷ�
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%if Request.QueryString("no")="eshop" then

dim SQL, Rs, contentID,CurrentPage
CurrentPage = request("Page")
contentID=request("ID")

set rs=server.createobject("adodb.recordset")
sqltext="delete from Bs_Honor where Id="& contentID
rs.open sqltext,conn,3,3
set rs=nothing
conn.close
response.redirect "Bs_honor.asp"
end if

%>
<%

set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Honor order by id desc"
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=10

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

call showpages
call list

If Rs.recordcount > MaxPerPage then
call showpages
end if

'��ʾ���ӵ��ӳ���
Sub list()%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>��˾��������</strong></td>
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
								<td width="25%"> <div align="center">��������</div></td>
								<td width="47%"> <div align="center">����֤��</div></td>
								<td width="10%"> <div align="center">����</div></td>
								<td width="10%"> <div align="center">����</div></td>
							</tr>
<%
if not rs.eof then
i=0
do while not rs.eof
%>
							<tr bgcolor="#E3E3E3"> 
							  <td height="22"> <div align="center"><%=rs("id")%></div></td>
								<td>&nbsp;&nbsp;<%=rs("title")%></td>
								<td> <div align="center"><img name="" src="<%=rs("img")%>" width="120" height="120" alt=""></div></td>
							  <td bgcolor="#E3E3E3"> <div align="center"><a href="Bs_honor_edit.asp?id=<%=rs("id")%>">�޸�</a></div></td>
								<td> <div align="center"><a href="Bs_honor.asp?id=<%=rs("id")%>&no=eshop">ɾ��</a></div></td>
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
Response.write "<strong><font color='#000000'>-> ȫ��-</font>"
Response.write "��</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>���������Ŀ</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>��</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='Bs_Honor.asp?&page="+cstr(1)+"'><font color='#000000'>��ҳ</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='Bs_Honor.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>��һҳ </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='Bs_Honor.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font>"
Response.write "<a href='Bs_Honor.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'>βҳ</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>��һҳ</font></strong>&nbsp;&nbsp;"
End if
'response.write "</td><td align='right'>"
'response.write "<font color='#000000'>ת����</font><input type='text'name='page'size=4 maxlength=4 class=smallInput value="&Currentpage&">&nbsp;"
'response.write "<input class=buttonface type='submit' value='Go' name='cndok'>&nbsp;&nbsp;"
%>
								</td>
							</tr>
						</table>
<%
End sub
rs.close

%>
					</div></td>
        </tr>
      </table>
      <br>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
