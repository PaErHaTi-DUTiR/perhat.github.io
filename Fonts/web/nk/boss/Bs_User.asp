<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->
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
   if(confirm("ȷ��Ҫɾ�����û���"))
     return true;
   else
     return false;
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>ע���Ա����</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
<%
  	if rs.eof and rs.bof then
		response.write "Ŀǰ���� 0 ��ע���û�"
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
        	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
        	else
	        	currentPage=1
        		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
      <table width="550" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
        <tr bgcolor="#C0C0C0" class="title"> 
          <td width="30" height="25" align="center"><strong> ���</strong></td>
          <td width="59" height="25" align="center"><strong> �û���</strong></td>
          <td width="29" height="25" align="center"><strong> �Ա�</strong></td>
          <td width="87" height="25" align="center"><strong>Email</strong></td>
          <td width="185" height="25" align="center"><strong>��˾����</strong></td>
          <td width="39" height="25" align="center"><strong> ״̬</strong></td>
          <td width="104" height="25" align="center" bgcolor="#C0C0C0"><strong>����</strong></td>
        </tr>
        <%do while not rs.EOF %>
        <tr bgcolor="#E3E3E3" class="tdbg"> 
          <td height="22" align="center"><%=rs("UserID")%></td>
          <td align="center"><%=rs("username")%></td>
          <td align="center"> 
            <%if rs("Sex")=1 then
	    response.write "��"
	  else
	    response.write "Ů"
	  end if%>
          </td>
          <td><a href="Bs_Mail_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></td>
          <td> <%=rs("Comane")%> </td>
          <td align="center"> 
            <%
	  if rs("LockUser")=true then
	  	response.write "������"
	  else
	  	response.write "����"
	  end if
	  %>
          </td>
          <td align="center"><a href="Bs_User_edit.asp?UserID=<%=rs("UserID")%>">�޸�</a>&nbsp; 
            <%if rs("LockUser")=False then %>
            <a href="Bs_User_Lock.asp?Action=Lock&UserID=<%=rs("UserID")%>">����</a> 
            <%else%>
            <a href="Bs_User_Lock.asp?Action=CancelLock&UserID=<%=rs("UserID")%>">����</a> 
            <%end if%>
            &nbsp;<a href="Bs_User_Del.asp?UserID=<%=rs("UserID")%>" onClick="return ConfirmDel();">ɾ��</a></td>
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
