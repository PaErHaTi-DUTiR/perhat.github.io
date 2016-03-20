<!--#include file="../inc/conn.asp"-->
<!--#include file="config.asp"-->
<%
imgpath="../counter/num/style"
set rs=server.createobject("adodb.recordset")
sqltext="select querycount from bs_sysdata where code='5062942'"
rs.open sqltext,conn,1,1
totcount=rs(0)
totcount=totcount+fake
rs.close
%>
<%if show="on" then %>
document.write("共有");
<%if style=0 then %>
document.write("<font color=red><%=totcount%></font>");
<%else %>
<%for i=1 to len(totcount)%>
document.write("<img src=<%=imgpath%><%=style%>/<%=mid(totcount,i,1)%>.gif width=12 height=12>");
<%next%>
<%end if %>
document.write("人次访问");
<%end if%>
<%set rs=nothing %>