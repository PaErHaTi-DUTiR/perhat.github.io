<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<%
'=========================================================
'
'��Ʒ���ƣ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ��ʽ��QQ 812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%
if Request.QueryString("no")="eshop" then
If request.form("title")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������⣬�뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if


If request.form("content")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����������ݣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Book"
rs.open sql,conn,1,3


rs.addnew
if request.form("html")="on" then
rs("content")=request.form("content")
else
rs("content")=htmlencode2(request.form("content"))
end if
rs("name")=request.form("name")
rs("title")=request.form("title")
rs("time")=date()
rs.update
rs.close
response.redirect "Bs_Book.asp"
end if

%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>���Թ���</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
					<td><div align="center"> 
			<% 
const MaxPerPage=5 '��ҳ��ʾ�ļ�¼���� 
dim sql 
dim rs 
dim totalPut 
dim CurrentPage 
dim TotalPages 
dim i,j 
 
 
if not isempty(request("page")) then 
currentPage=cint(request("page")) 
else 
currentPage=1 
end if 
set rs=server.createobject("adodb.recordset") 
sql="select * from Bs_Book order by id desc" 
rs.open sql,conn,1,1 
 
if rs.eof and rs.bof then 
response.write "<p align='center'>��û���κ�����!</p>" 
else 
totalPut=rs.recordcount 
totalPut=rs.recordcount 
if currentpage<1 then 
currentpage=1 
end if 
if (currentpage-1)*MaxPerPage>totalput then 
if (totalPut mod MaxPerPage)=0 then 
currentpage= totalPut \ MaxPerPage 
else 
currentpage= totalPut \ MaxPerPage + 1 
end if 
end if 
if currentPage=1 then 
showpages 
showContent 
showpages 
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showpages 
showContent 
showpages 
else 
currentPage=1 
showContent 
end if 
end if 
rs.close 
end if 
set rs=nothing 
conn.close 
set conn=nothing 
sub showContent 
dim i 
i=0 
do while not (rs.eof or err) %>
                
              <table width="100%" border="0" cellspacing="2" cellpadding="3">
                <tr bgcolor="#C0C0C0"> 
                  <td width="18%" height="25" bgcolor="#C0C0C0"> <div align="center"> 
                      <%if rs("name")="����Ա"then%>
                      ���̳ǹ��桻 
                      <%else%>
                      �û��� 
                      <%end if%>
                    </div></td>
                  <td width="82%" bgcolor="#C0C0C0">&nbsp;&nbsp;<%=rs("name")%> 
                    &nbsp; <a href="Bs_Book_Del.asp?id=<%=rs("ID")%>">ɾ��</a>&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href="Bs_Book_Re.asp?id=<%=rs("id")%>"> 
                    <%if rs("name")<>"����Ա" then%>
                    <%if rs("name")<>"δע���û�" then 
response.write "�ظ�" 
end if 
%>
                    </a> 
                    <%if rs("name")="δע���û�" then 
response.write "û��ע��Ϊ��Աͨ��E-amil�ظ���" 
end if%>
                    <%end if%>
                  </td>
                </tr>
                <%if rs("email")<>"" then%>
                <tr bgcolor="#E3E3E3"> 
                  <td height="10"> <div align="center">��˾����</div></td>
                  <td bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("Comane")%></td>
                </tr>
                <tr bgcolor="#E3E3E3"> 
                  <td height="10"><div align="center">��ϵ��</div></td>
                  <td bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("Somane")%></td>
                </tr>
                <tr bgcolor="#E3E3E3"> 
                  <td height="22"><div align="center">��ϵ�绰</div></td>
                  <td bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("Phone")%></td>
                </tr>
                <tr bgcolor="#E3E3E3"> 
                  <td height="22"><div align="center">��ϵ����</div></td>
                  <td bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("Fox")%></td>
                </tr>
                <tr bgcolor="#E3E3E3"> 
                  <td height="22"><div align="center">E-mail</div></td>
                  <td bgcolor="#E3E3E3"><a href="Bs_Mail_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></td>
                </tr>
                <%end if%>
                <tr bgcolor="#C0C0C0"> 
                  <td height="25" bgcolor="#C0C0C0"><div align="center">�� ��</div></td>
                  <td height="25">&nbsp;&nbsp;<%=rs("title")%>&nbsp;&nbsp;[<%=rs("time")%>]</td>
                </tr>
                <tr bgcolor="#E3E3E3"> 
                  <td height="25"> <div align="center">�� ��</div></td>
                  <td height="25">&nbsp;&nbsp;<%=rs("content")%></td>
                </tr>
                <%if rs("rebook")<>""then%>
                <tr bgcolor="#E3E3E3"> 
                  <td height="25"> <div align="center">�ظ�����</div></td>
                  <td height="25">&nbsp;&nbsp;<%=rs("rebook")%></td>
                </tr>
                <%end if%>
              </table>
              </div></td>
        </tr>
      </table>
      <br>
      <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><div align="center"><% i=i+1 
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
end sub 
sub showpages() 
dim n 
if (totalPut mod MaxPerPage)=0 then 
n= totalPut \ MaxPerPage 
else 
n= totalPut \ MaxPerPage + 1 
end if 
if n=1 then 
response.write "���Բ��������" 
 
exit sub 
end if 
dim k 
response.write "<p align='left'>&gt;&gt; ���Է�ҳ " 
for k=1 to n 
if k=currentPage then 
response.write "[<b>"+Cstr(k)+"</b>] " 
else 
response.write "[<b>"+"<a href='Bs_Book.asp?page="+cstr(k)+"'>"+Cstr(k)+"</a></b>] " 
end if 
next 
end sub 
%></div></td>
        </tr>
      </table> 
      <br>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
