<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!-- #include file="Inc/Head.asp" -->
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
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>应聘管理</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
					<td><div align="center"> 
<% 
const MaxPerPage=5 '分页显示的纪录个数 
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
sql="select * from Bs_Jobbook order by id desc" 
rs.open sql,conn,1,1 
 
if rs.eof and rs.bof then 
response.write "<p align='center'>还没有任何应聘信息!</p>" 
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
do while not (rs.eof or err) 
%>
                
						<table width="100%" border="0" cellspacing="2" cellpadding="3">
							<tr bgcolor="#C0C0C0"> 
								<td width="17%" height="25" bgcolor="#C0C0C0"> <div align="center">应聘岗位</div></td>
								<td colspan="2" bgcolor="#C0C0C0">&nbsp;&nbsp;<b><%=rs("jobname")%></b> 
									&nbsp; </td>
								<td bgcolor="#C0C0C0"><a href="Bs_Job_Book_Del.asp?id=<%=rs("ID")%>">删除</a>&nbsp;&nbsp;&nbsp;
								<a href="Bs_Mail_Send.asp?email=<%=rs("email")%>">发信</a> 
								</td>
							</tr>
							<%if rs("email")<>"" then%>
							<tr bgcolor="#E3E3E3"> 
								<td height="10"> <div align="center">姓 名</div></td>
								<td width="34%" bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("mane")%></td>
								<td width="13%" bgcolor="#E3E3E3"><div align="center">性 别</div></td>
								<td width="36%" bgcolor="#E3E3E3"><%=rs("sex")%></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="22" bgcolor="#E3E3E3"> <div align="center">出生日期</div></td>
								<td bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("birthday")%></td>
								<td bgcolor="#E3E3E3"><div align="center">婚姻状况</div></td>
								<td bgcolor="#E3E3E3"><%=rs("marry")%></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="22"><div align="center">毕业院校</div></td>
								<td colspan="3" bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("school")%></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="22"><div align="center">学 历</div></td>
								<td bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("studydegree")%></td>
								<td bgcolor="#E3E3E3"><div align="center">专 业</div></td>
								<td bgcolor="#E3E3E3"><%=rs("specialty")%></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="22"><div align="center">毕业时间</div></td>
								<td bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("gradyear")%></td>
								<td bgcolor="#E3E3E3"><div align="center"><font color="#FF0000">应聘日期</font></div></td>
								<td bgcolor="#E3E3E3"><font color="#FF0000"><%=rs("time")%></font></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="22"><div align="center">电 话</div></td>
								<td colspan="3" bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("telephone")%></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="22"><div align="center">E-mail</div></td>
								<td colspan="3" bgcolor="#E3E3E3">&nbsp;&nbsp;<a href="Bs_Mail_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="22"><div align="center">联系地址</div></td>
								<td colspan="3" bgcolor="#E3E3E3">&nbsp;&nbsp;<%=rs("address")%></td>
							</tr>
							<%end if%>
							<tr bgcolor="#C0C0C0"> 
								<td height="25" bgcolor="#C0C0C0"><div align="center">水平与能力</div></td>
								<td height="25" colspan="3">&nbsp;&nbsp;<%=rs("ability")%></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="25" bgcolor="#C0C0C0"> <div align="center">个人简历</div></td>
								<td height="25" colspan="3" bgcolor="#C0C0C0">&nbsp;&nbsp;<%=rs("resumes")%></td>
							</tr>
						</table>
					</div></td>
        </tr>
      </table>
      <br>
      <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><div align="center">
<%
i=i+1 
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
response.write "留言簿管理界面" 
 
exit sub 
end if 
dim k 
response.write "<p align='left'>&gt;&gt; 留言分页 " 
for k=1 to n 
if k=currentPage then 
response.write "[<b>"+Cstr(k)+"</b>] " 
else 
response.write "[<b>"+"<a href='Bs_Job_Book.asp?page="+cstr(k)+"'>"+Cstr(k)+"</a></b>] " 
end if 
next 
end sub 
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
