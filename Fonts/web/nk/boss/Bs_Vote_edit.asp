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
<%
dim ID,Title,VoteTime,IsSelected
dim rs,sql
ID=request("id")
Title=trim(request.form("Title"))
VoteTime=trim(request.form("VoteTime"))
if VoteTime="" then VoteTime=now()
IsSelected=trim(request("IsSelected"))
if IsSelected="True" then
	conn.execute "Update Vote set IsSelected=False where IsSelected=True"
end if
dim i
if ID="" then
	Response.Redirect "Bs_Vote.asp"
end if
sql="select * from Vote where ID="&Cint(ID)
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if not rs.eof then
	if Title<>"" then
		rs("Title")=Title
		for i=1 to 8
			if trim(request("select"&i))<>"" then
				rs("select"&i)=trim(request("select"&i))
				if request("answer"&i)="" then
					rs("answer"&i)=0
				else
					rs("answer"&i)=clng(request("answer"&i))
				end if
			else
				rs("select"&i)=""
				rs("answer"&i)=0
			end if
		next
		rs("VoteTime")=VoteTime
		rs("VoteType")=request("VoteType")
		if IsSelected="" then IsSelected=false
		rs("IsSelected")=IsSelected
		rs.update
		rs.close
		set rs=nothing
		call CloseConn()
		Response.Redirect "Bs_Vote.asp"
	end if
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
			<form method="POST" action="Bs_Vote_edit.asp">
        <table width="500" border="0" align="center" cellpadding="0" cellspacing="2" Class="border">
          <tr class="tdbg"> 
						<td width="15%" align="right" bgcolor="#C0C0C0">���⣺</td>
						<td width="85%" colspan="3" bgcolor="#E3E3E3"><input name="Title" type="text" value="<%=rs("Title")%>" size="52"></td>
					</tr>
					<% for i=1 to 8%>
					<tr class="tdbg"> 
						<td align="right" bgcolor="#C0C0C0">ѡ��<%=i%>��</td>
						<td bgcolor="#E3E3E3"><input name="select<%=i%>" type="text" value="<%=rs("select"& i)%>" size="36"> 
						</td>
						<td align="right" bgcolor="#E3E3E3">Ʊ����</td>
						<td width="80" bgcolor="#E3E3E3"><input name="answer<%=i%>" type="text" value="<%=rs("answer"&i)%>" size="5"></td>
					</tr>
					<%next%>
					<tr class="tdbg"> 
						<td align="right" bgcolor="#C0C0C0">�������ͣ�</td>
						<td colspan="3" bgcolor="#E3E3E3"><select name="VoteType" id="VoteType">
							<option value="Single" <% if rs("VoteType")="Single" then %> selected <% end if%>>��ѡ</option>
							<option value="Multi" <% if rs("VoteType")="Multi" then %> selected <% end if%>>��ѡ</option>
						</select></td>
					</tr>
					<tr class="tdbg">
						<td align="right" bgcolor="#C0C0C0">&nbsp;</td>
						<td colspan="3" bgcolor="#E3E3E3">
							<input name="IsSelected" type="checkbox" id="IsSelected" value="True" <% if rs("IsSelected")=true then response.write "checked"%>>
							��Ϊ���µ���</td>
					</tr>
					<tr class="tdbg">
						<td colspan=4 align=center> <BR>
							<input name="ID" type="hidden" id="ID" value="<%=rs("ID")%>"> 
							&nbsp;<input name="Submit" type="submit" id="Submit" value="�����޸Ľ��"> </td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
