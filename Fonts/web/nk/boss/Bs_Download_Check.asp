<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
'=========================================================
'
'��Ʒ���ƣ��Ƽ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ ��ʽ��QQ ��812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<%
dim strFileName
const MaxPerPage=20
dim totalPut   
dim CurrentPage
dim TotalPages
dim i,j
dim Bs_DownID
dim Bs_Title
dim Bs_BigClassName,Bs_SmallClassName
dim Bs_Passed
dim PurviewChecked
dim strAdmin,arrAdmin

PurviewChecked=false
Bs_Passed=trim(request("Bs_Passed"))
strFileName="Bs_Download_Check.asp?Bs_BigClassName=" & Bs_BigClassName & "&Bs_SmallClassName=" & Bs_SmallClassName

if Bs_Passed="" then
	if Session("Bs_Passed")="" then
		Bs_Passed="False"
	else
		Bs_Passed=Session("Bs_Passed")
	end if
end if
session("Bs_Passed")=Bs_Passed
Bs_Title=Trim(request("Bs_Title"))
Bs_DownID=Request("Bs_DownID")
Bs_BigClassName=Trim(request("Bs_BigClassName"))
Bs_SmallClassName=Trim(request("Bs_SmallClassName"))
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

dim sql
dim rs
sql="select * from Bs_Download where Bs_Passed=" & Bs_Passed
if Bs_Title<>"" then
	sql=sql & " and Bs_Title like '%" & Bs_Title & "%'"
else
	if Bs_BigClassName<>"" then
		sql=sql & " and Bs_BigClassName='" & Bs_BigClassName & "'"
		if Bs_SmallClassName<>"" then
			sql=sql & " and Bs_SmallClassName='" & Bs_SmallClassName & "'"
		end if
	else
'		if SpecialName<>"" then
'			sql=sql & " and SpecialName='" & SpecialName & "'"
'		end if
	end if
end if
sql=sql & " order by Bs_DownID desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.Check.chkAll.checked){
	document.Check.chkAll.checked = document.Check.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
</SCRIPT>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="610" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center"><BR>
      <form name="Bs_Passed" method="Get" action="Bs_Download_Check.asp">
				<div align="center"><strong>����ѡ�</strong> 
				<input name="Bs_Passed" type="radio" value="False" onclick="submit();" <%if Bs_Passed="False" then response.write " checked"%>>
				δ��˵�����&nbsp;&nbsp;&nbsp;&nbsp; 
				<input name="Bs_Passed" type="radio" value="True" onclick="submit();" <%if Bs_Passed="True" then response.write " checked"%>>
				����˵�����
				<input name="Bs_BigClassName" type="hidden" id="Bs_BigClassName" value="<%=Bs_BigClassName%>">
				<input name="Bs_SmallClassName" type="hidden" id="Bs_SmallClassName" value="<%=Bs_SmallClassName%>">
				</div>
			</form>
			<table width="600" border="0" align="center" cellpadding="5" cellspacing="2" class="border">
				<tr class="title"> 
					<td height="25" bgcolor="#999999">|&nbsp; 
<%
dim sqlBigClass,sqlSmallClass,rsBigClass,rsSmallClass,sqlSpecial,rsSpecial
sqlBigClass="select * from Bs_DownBigClass"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1
if rsBigClass.eof then 
	response.Write("��û���κ����������������")
end if
do while not rsBigClass.eof
	if rsBigClass("Bs_BigClassName")=Bs_BigClassName then
		response.Write("<a href='Bs_Download_Check.asp?Bs_BigClassName=" & rsBigClass("Bs_BigClassName") & "'><font color='red'>" & rsBigClass("Bs_BigClassName") & "</font></a> | ")
		if session("purview")=3 then
			strAdmin=rsBigClass("Admin")
			if Instr(strAdmin,"|")>0 then
				arrAdmin=split(strAdmin)
				for i=0 to ubound(arrAdmin)
					if trim(arrAdmin(i))=session("admin") then
						PurviewChecked=True
						exit for
					end if
				next
			else
				if trim(strAdmin)=session("Admin") then
					PurviewChecked=True
				end if
			end if
		end if
	else
		response.Write("<a href='Bs_Download_Check.asp?Bs_BigClassName=" & rsBigClass("Bs_BigClassName") & "'>" & rsBigClass("Bs_BigClassName") & "</a> | ")
	end if
	rsBigClass.movenext
loop
rsBigClass.close
set rsBigClass=nothing
%>
					</td>
				</tr>
<%
if Bs_BigClassName<>"" then
	sqlSmallClass="select * from Bs_DownSmallClass where Bs_BigClassName='" & Bs_BigClassName & "'"
	Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
	rsSmallClass.open sqlSmallClass,conn,1,1
	if not (rsSmallClass.bof and rsSmallClass.eof) then
		response.write "<tr class='tdbg'><td bgcolor='#CCCCCC'>"
		do while not rsSmallClass.eof
			if rsSmallClass("Bs_SmallClassName")=Bs_SmallClassName then
				response.Write("&nbsp;<a href='Bs_Download_Check.asp?Bs_BigClassName=" & rsSmallClass("Bs_BigClassName") & "&Bs_SmallClassName=" & rsSmallClass("Bs_SmallClassName") & "'><font color='red'>" & rsSmallClass("Bs_SmallClassName") & "</font></a>&nbsp;&nbsp;")
				if session("purview")=4 then
					strAdmin=rsSmallClass("Admin")
					if Instr(strAdmin,"|")>0 then
						arrAdmin=split(strAdmin)
						for i=0 to ubound(arrAdmin)
							if trim(arrAdmin(i))=session("admin") then
								PurviewChecked=True
								exit for
							end if
						next
					else
						if trim(strAdmin)=session("Admin") then
							PurviewChecked=True
						end if
					end if
				end if
			else
				response.Write("&nbsp;<a href='Bs_Download_Check.asp?Bs_BigClassName=" & rsSmallClass("Bs_BigClassName") & "&Bs_SmallClassName=" & rsSmallClass("Bs_SmallClassName") & "'>" & rsSmallClass("Bs_SmallClassName") & "</a>&nbsp;&nbsp;")
			end if
			rsSmallClass.movenext
		loop
		response.write "</td></tr>"
	end if
	rsSmallClass.close
	set rsSmallClass=nothing
end if
%>
			</table>
			<form action="Bs_Download_Check_Set.asp" method="Post" name="Check" id="Check">
			<table width="600" height="22" border="0" cellpadding="0" cellspacing="2">
				<tr> 
					<td width="453" height="22" bgcolor="#CCCCCC"><a href="Bs_Download_Check.asp">&nbsp;�������</a>&nbsp;&gt;&gt;&nbsp; 
<%
if Bs_Title="" and Bs_BigClassName="" and Bs_SmallClassName="" then
	if Bs_Passed="False" then
		response.write "����δ��˵�����"
	else
		response.write "��������˵�����"
	end if
else
	if request("Query")<>"" then
		if Bs_Title<>"" then
			response.write "�����к��С�<font color=blue>" & Bs_Title & "</font>������δ��˵�����"
		else
			response.Write("����δ��˵�����")
		end if
 	else
		if Bs_BigClassName<>"" then
			response.write "<a href='Bs_Download_Check.asp?Bs_BigClassName=" & Bs_BigClassName & "'>" & Bs_BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if Bs_SmallClassName<>"" then
				response.write "<a href='Bs_Download_Check.asp?Bs_BigClassName=" & Bs_BigClassName & "&Bs_SmallClassName=" & Bs_SmallClassName & "'>" & Bs_SmallClassName & "</a>"
			else
				response.write "����С��"
			end if
		end if
	end if
end if
%>
					</td>
					<td width="161" bgcolor="#CCCCCC"> &nbsp; 
<%
	if rs.eof and rs.bof then
	response.write "���ҵ� 0 ������</td></tr></table>"
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
	response.Write "���ҵ� " & totalPut & " ������"
%>
					</td>
				</tr>
			</table>
<%		
		if currentPage=1 then
				showContent
				showpage strFileName,totalput,MaxPerPage,true,false," ������"
		else
				if (currentPage-1)*MaxPerPage<totalPut then
						rs.move  (currentPage-1)*MaxPerPage
					dim bookmark
						bookmark=rs.bookmark
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," ������"
				else
					currentPage=1
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," ������"
			end if
	end if
end if
%>
<%  
sub showContent
   	dim i
    i=0
%>
			<table class="border" border="0" cellspacing="2" width="600" cellpadding="0" style="word-break:break-all">
				<tr bgcolor="#C0C0C0" class="title"> 
					<td width="30" height="25" align="center" bgcolor="#C0C0C0"><strong>ѡ��</strong></td>
					<td width="37"  height="20" align="center"><strong>ID</strong></td>
					<td width="317" align="center" ><strong>��������</strong></td>
					<td width="63" align="center" ><strong>����ʱ��</strong></td>
					<td width="57" align="center" ><strong>�����</strong></td>
					<td width="40" align="center" ><strong>�����</strong></td>
					<td width="60" align="center" ><strong>����</strong></td>
				</tr>
<%do while not rs.eof%>
				<tr class="tdbg"> 
					<td width="30" height="22" align="center" bgcolor="#C0C0C0"> 
					<input name='Bs_DownID'type='checkbox'onclick="unselectall()" id="Bs_DownID2" value='<%=cstr(rs("Bs_DownID"))%>'>
					</td>
					<td width="37" align="center" bgcolor="#E3E3E3"><%=rs("Bs_DownID")%></td>
					<td bgcolor="#E3E3E3"> <a href="../Bs_DownloadShow.asp?Bs_DownID=<%=rs("Bs_DownID")%>" 
					title="<%=replace(left(nohtml(rs("Bs_Content")),200),chr(34),"") & "����"%>">&nbsp;<%=rs("Bs_Title")%></a></td>
					<td width="63" align="center" bgcolor="#E3E3E3"><%= FormatDateTime(rs("Bs_UpdateTime"),2) %></td>
					<td width="57" align="center" bgcolor="#E3E3E3"><%=rs("Bs_Hits")%> </td>
					<td width="40" align="center" bgcolor="#E3E3E3">
<%if rs("Bs_Passed")=true then response.write "��" else response.write "��" end if%>
					</td>
					<td width="60" align="center" bgcolor="#E3E3E3"> 
<%
if session("purview")<=2 or PurviewChecked=True then
	If rs("Bs_Passed")=False Then 
%>
<a href="Bs_Download_Check_Set.asp?Action=Check&Bs_DownID=<%=rs("Bs_DownID")%>">ͨ�����</a> 
<% Else %>
<a href="Bs_Download_Check_Set.asp?Action=CancelCheck&Bs_DownID=<%=rs("Bs_DownID")%>">ȡ��ͨ��</a> 
<%
	End If
end if
%>
					</td>
				</tr>
<%
i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
			</table>
			<table width="600" border="0" cellpadding="0" cellspacing="0">
				<tr> 
					<td width="250" height="30">
					<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">ѡ�б�ҳ��ʾ����������</td>
					<td><input name="submit" type='submit'
					value="<%if Bs_Passed="True" then response.write "ȡ��"%>���ѡ��������" 
					<%if session("purview")>=3 and session("purview")<=4 and PurviewChecked=False then response.write "disabled"%>> 
					<input name="Action" type="hidden" id="Action" value="
<%
if Bs_Passed="False" then
response.write "Check"
else
response.write "CancelCheck"
end if%>"></td>
				</tr>
			</table>
<%
end sub 
%>
			</form>
			<table width="600" class="border">
				<tr class="tdbg"> 
					<form name="searchsoft" method="get" action="Bs_Download_Check.asp">
					<td height="30"> <strong>�������أ�</strong> 
					<input name="Bs_Title" type="text" class=smallInput id="Bs_Title3" size="20"> 
					<input name="Query" type="submit" id="Query" value="�� ѯ">
					&nbsp;&nbsp;�������������ơ����Ϊ�գ�������������ء� </td>
					</form>
        </tr>
      </table><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
rs.close
set rs=nothing  
call CloseConn()

sub CheckArticle(id)
    dim sql
    sql="update Bs_Download set Bs_Passed = True where Bs_DownID="&cstr(id)
    conn.execute sql
    if err.Number<>0 then
		err.clear
		'response.write "ɾ �� ʧ �� !<br>"
    else
        'response.write "ɾ���ɹ���<br>"
    end if	
End sub
sub CancelCheckArticle(id)
    dim sql
    sql="update Bs_Download set Bs_Passed = False where Bs_DownID="&cstr(id)
    conn.execute sql
    if err.Number<>0 then
		err.clear
		'response.write "ɾ �� ʧ �� !<br>"
    else
        'response.write "ɾ���ɹ���<br>"
    end if	
End sub
%>
