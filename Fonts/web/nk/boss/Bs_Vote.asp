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
dim sql,rs,Action,ID
Action=Trim(Request("Action"))
ID=Trim(Request("VoteID"))
if Action="Set" and ID<>"" then
	conn.execute "Update Vote set IsSelected=False where IsSelected=True"
	conn.execute "Update Vote set IsSelected=True Where ID=" & ID
	response.Write "<script language='JavaScript'type='text/JavaScript'>alert('设置成功！');</script>"
end if
sql="select * from Vote order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>调 查 管 理</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<br>
			<form method="POST" action="Bs_Vote.asp">
        <table width="500" border="0" align="center" cellpadding="0" cellspacing="2" Class="border">
          <tr bgcolor="#C0C0C0" class="title"> 
            <td width="37" height="25" align="center">选择</td>
            <td width="37" align="center">ID</td>
            <td width="328" align="center">主题</td>
            <td width="88" align="center">操作</td>
					</tr>
<%
if not (rs.bof and rs.eof) then
	do while not rs.eof
%>
          <tr bgcolor="#E3E3E3" class="tdbg"> 
            <td width="37" height="22" align="center"> 
              <input type="radio" value=<%=rs("ID")%><%if rs("IsSelected")=true then%> checked<%end if%> name="VoteID"></td>
            <td width="37" height="22" align="center"><%=rs("ID")%></td>
            <td height="22">&nbsp;<%=rs("Title")%></td>
            <td width="88" height="22" align="center"><a href="Bs_Vote_edit.asp?ID=<%=rs("ID")%>">修改</a> 
              <a href="Bs_Vote_Del.asp?ID=<%=rs("ID")%>">删除</a></td>
					</tr>
<%
rs.movenext
loop
%>
					<tr class="tdbg"> 
						<td colspan=4 align=center><BR> <input name="Action" type="hidden" id="Action" value="Set"> 
							<input type="submit" value="将选定的调查设为最新调查" name="submit">
						</td>
					</tr>
<% end if%>
				</table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
