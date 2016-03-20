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
<%
method=Request.QueryString("method")
pageno=Request.QueryString("pageno")

htmltitle="网站访问情况"
htmlname="sysadmin_count_list.asp?flag=1"
set rs=server.CreateObject("ADODB.Recordset")

if method="clearall" then
	sqltext="delete from webcount"
	conn.Execute sqltext
	sqltext="update bs_sysdata set querycount=0"
	conn.Execute sqltext
end if 
sqltext="select querycount from Bs_SysData where code='5062942'"
rs.Open sqltext,conn,1,1
totcount=rs(0)
rs.Close
%>
<html>
<head>
<title><%=htmltitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {font-size: 12px; color: #000000; font-family: 宋体}
td {font-size: 12px; color: #000000; font-family: 宋体;line-height:130%}
.t1 {font:12px 宋体;color=000000} 
.t2 {font:12px 宋体;color:ffffff} 
.t3 {font:12px 宋体;color:336699} 
.t4 {font:12px 宋体;color:ff0000} 

.bt1 {font:14px 宋体;color=000000} 
.bt2 {font:12px 宋体;color:ffffff} 
.bt3 {font:14px 宋体;color:336699} 
.bt4 {font:14px 宋体;color:ff0000} 
.bt5 {font:14px 宋体;color:0000ff}
.bt10 {font:14px 宋体;color:0000ff}

.td1 {font-size:12px;background-color:#3388bb;color:#ffffff}
.td2 {font-size:12px;background-color:#ffffff;color:#000000;}

A:link {color: #333333}
A:visited {color: #333333}
A:hover {color: #ff0000}

A.r:link {color: #333333}
A.r:visited {color: #333333}
A.r:hover {color: #ff0000}

A.title1:link {text-decoration:none;color:#333333;}
A.title1:visited {text-decoration:none;color:#333333;}
A.title1:active {text-decoration:none;color:#333333;}
A.title1:hover {text-decoration:none;color:#333333;}
.STYLE1 {color: #333333}
-->
</style>
</head>
<body topmargin=5>
<table width="100%" align=center>
	<tr>
	<td valign="top" align="left" width="100%">
		<form name="form1">
		<table width="100%" height="20" border="0">
			<tr>
			<td style="font-size:12px;"></td>
			<td><font color="navy" style="font-size:14px"><b><%=htmltitle%></b>(共有<font color=red><%=totcount%></font>人次访问)</font></td>
			<td align="right">
			</td>
			</tr>
		</table>
		<table width="100%" height="20" border="0">
			<tr>
		<%
		sqltext="select * from webcount order by keyno desc"
		'Response.Write sqltext
		rs.Open sqltext,conn,1,1

	    colzs=2
		rowzs=30	
  		rs.PageSize=colzs*rowzs
		if pageno="" then 
			pageno=1
		else
			pageno=cint(pageno)	
		end if	
		if rs.PageCount>0 then rs.AbsolutePage=pageno
	%>
		<td class=t1 align="left" width=200>
			<a href="Javascript:window.location.reload()">[刷新列表]</a>
			<!-- <a href="sysadmin_count_report.asp">[日访问量]</a> -->
			<a href="Javascript:delrec()">[清空记录]</a>
			</td>
		<td  align="right">
		<%if rs.PageCount>1 then%>
			[第<b><%=pageno%></b>页, 共<b><%=rs.PageCount%></b>页<b><%=rs.RecordCount%></b>条记录]		              
			<%if pageno>1 then%>
				<a class=r href="<%=htmlname%>&pageno=1">首页</a>              
				<a class=r href="<%=htmlname%>&pageno=<%=pageno-1%>">上页</a>              
			<%end if              
			if pageno<rs.PageCount then%>              
				<a class=r href="<%=htmlname%>&pageno=<%=pageno+1%>">下页</a>              
				<a class=r href="<%=htmlname%>&pageno=<%=rs.PageCount%>">末页</a>              
			<%end if%>              
		<%end if              
		%>              
		</td>
		</tr>
		</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" style="font-size:14.5px;line-height:100%">
				<tr bgcolor="336699" style="color:#ffffff" align="center">
				<td bgcolor="#666666"></td>
				<td bgcolor="#666666" class=bt2>登录时间</td>
				<td bgcolor="#666666" class=bt2>IP地址</td>
				<td bgcolor="#666666"></td>
				<td bgcolor="#666666" class=bt2>登录时间</td>
				<td bgcolor="#666666" class=bt2>IP地址</td>
				<td bgcolor="#666666"></td>
				</tr>
		<%
		i=1
	    do while i<=rowzs and not rs.EOF
			if (i mod 2)=0 then
				Response.Write("<tr bgcolor=fefefe>")
			else
				Response.Write("<tr bgcolor=efefef>")
			end if	
		    for j=1 to colzs
				if not rs.EOF then
			%>			
				<td width=30 align=right><%=((pageno-1)*colzs*rowzs+(i-1)*colzs+j)%></td>
				<td width="150" align=right><%=rs("logindate")%>
				</td>
				<td width="120" align=right><%=trim(rs("ipaddress"))%></td>
				<%
					rs.MoveNext
				else
				%>
				<td width=30>&nbsp;</td>
				<td width="120" align=left></td>
				<td width="120" align=right></td>
				<%
				end if
			next	
			%>
			<td></td>
			</tr>
		<%	
			i=i+1
		loop
		%>
			</table>
		</td></tr>
	<tr><td>
		<table width="100%" cellspacing="0" cellpadding="0" border=0> 
		<tr><td><hr size=1></td></tr>
			<tr><td align=right>
			[第<b><%=pageno%></b>页, 共<b><%=rs.PageCount%></b>页<b><%=rs.RecordCount%></b>条记录]			
			<%if pageno>1 then%>
				<a href="<%=htmlname%>&pageno=1" class=r1 STYLE1>[首页]</a>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=pageno-1%>">[上页]</a>              
			<%end if              
			if pageno<rs.PageCount then%>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=pageno+1%>">[下页]</a>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=rs.PageCount%>">[末页]</a>              
			<%end if%>              
			</td></tr>              
		</table>
	</td></tr>
		</form>
</td></tr><table>
<%htmlend%>

<script language="Javascript">
	function delrec()
	{
		if (confirm('您确认要清空网站访问记录吗？')==true)
		{
		window.location.href="<%=htmlname%>&method=clearall"
		return false;
		}
	}

</script>

