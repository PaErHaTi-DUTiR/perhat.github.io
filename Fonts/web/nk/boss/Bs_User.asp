<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->
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
<%
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim rs, sql
strFileName="Bs_User.asp"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from [Bs_User] order by UserID desc"
rs.Open sql,conn,1,1
%>
<script language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除此用户吗？"))
     return true;
   else
     return false;
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>注册会员管理</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
<%
  	if rs.eof and rs.bof then
		response.write "目前共有 0 个注册用户"
	else
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
        	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
        	else
	        	currentPage=1
        		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
      <table width="550" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
        <tr bgcolor="#C0C0C0" class="title"> 
          <td width="30" height="25" align="center"><strong> 序号</strong></td>
          <td width="59" height="25" align="center"><strong> 用户名</strong></td>
          <td width="29" height="25" align="center"><strong> 性别</strong></td>
          <td width="87" height="25" align="center"><strong>Email</strong></td>
          <td width="185" height="25" align="center"><strong>公司名称</strong></td>
          <td width="39" height="25" align="center"><strong> 状态</strong></td>
          <td width="104" height="25" align="center" bgcolor="#C0C0C0"><strong>操作</strong></td>
        </tr>
        <%do while not rs.EOF %>
        <tr bgcolor="#E3E3E3" class="tdbg"> 
          <td height="22" align="center"><%=rs("UserID")%></td>
          <td align="center"><%=rs("username")%></td>
          <td align="center"> 
            <%if rs("Sex")=1 then
	    response.write "男"
	  else
	    response.write "女"
	  end if%>
          </td>
          <td><a href="Bs_Mail_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></td>
          <td> <%=rs("Comane")%> </td>
          <td align="center"> 
            <%
	  if rs("LockUser")=true then
	  	response.write "已锁定"
	  else
	  	response.write "正常"
	  end if
	  %>
          </td>
          <td align="center"><a href="Bs_User_edit.asp?UserID=<%=rs("UserID")%>">修改</a>&nbsp; 
            <%if rs("LockUser")=False then %>
            <a href="Bs_User_Lock.asp?Action=Lock&UserID=<%=rs("UserID")%>">锁定</a> 
            <%else%>
            <a href="Bs_User_Lock.asp?Action=CancelLock&UserID=<%=rs("UserID")%>">解锁</a> 
            <%end if%>
            &nbsp;<a href="Bs_User_Del.asp?UserID=<%=rs("UserID")%>" onClick="return ConfirmDel();">删除</a></td>
        </tr>
        <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
      </table>  
<%
end sub 
%>
		<BR></td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
rs.Close
set rs=Nothing
call CloseConn()  
%>
