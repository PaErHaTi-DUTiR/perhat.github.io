<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
'=========================================================
'
'产品名称：科技 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有：www.web300.cn
'程序制作：QQ812256@hotmail.com
'联系 方式：QQ ：812256
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
		<td class="a1" height="25" align="center"><strong>下 载 审 核</strong></td>
	</tr>
	<tr class="a4">
		<td align="center"><BR>
      <form name="Bs_Passed" method="Get" action="Bs_Download_Check.asp">
				<div align="center"><strong>下载选项：</strong> 
				<input name="Bs_Passed" type="radio" value="False" onclick="submit();" <%if Bs_Passed="False" then response.write " checked"%>>
				未审核的下载&nbsp;&nbsp;&nbsp;&nbsp; 
				<input name="Bs_Passed" type="radio" value="True" onclick="submit();" <%if Bs_Passed="True" then response.write " checked"%>>
				已审核的下载
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
	response.Write("还没有任何类别，请首先添加类别。")
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
					<td width="453" height="22" bgcolor="#CCCCCC"><a href="Bs_Download_Check.asp">&nbsp;下载审核</a>&nbsp;&gt;&gt;&nbsp; 
<%
if Bs_Title="" and Bs_BigClassName="" and Bs_SmallClassName="" then
	if Bs_Passed="False" then
		response.write "所有未审核的下载"
	else
		response.write "所有已审核的下载"
	end if
else
	if request("Query")<>"" then
		if Bs_Title<>"" then
			response.write "标题中含有“<font color=blue>" & Bs_Title & "</font>”并且未审核的下载"
		else
			response.Write("所有未审核的下载")
		end if
 	else
		if Bs_BigClassName<>"" then
			response.write "<a href='Bs_Download_Check.asp?Bs_BigClassName=" & Bs_BigClassName & "'>" & Bs_BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if Bs_SmallClassName<>"" then
				response.write "<a href='Bs_Download_Check.asp?Bs_BigClassName=" & Bs_BigClassName & "&Bs_SmallClassName=" & Bs_SmallClassName & "'>" & Bs_SmallClassName & "</a>"
			else
				response.write "所有小类"
			end if
		end if
	end if
end if
%>
					</td>
					<td width="161" bgcolor="#CCCCCC"> &nbsp; 
<%
	if rs.eof and rs.bof then
	response.write "共找到 0 个下载</td></tr></table>"
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
	response.Write "共找到 " & totalPut & " 个下载"
%>
					</td>
				</tr>
			</table>
<%		
		if currentPage=1 then
				showContent
				showpage strFileName,totalput,MaxPerPage,true,false," 个下载"
		else
				if (currentPage-1)*MaxPerPage<totalPut then
						rs.move  (currentPage-1)*MaxPerPage
					dim bookmark
						bookmark=rs.bookmark
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," 个下载"
				else
					currentPage=1
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," 个下载"
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
					<td width="30" height="25" align="center" bgcolor="#C0C0C0"><strong>选择</strong></td>
					<td width="37"  height="20" align="center"><strong>ID</strong></td>
					<td width="317" align="center" ><strong>下载名称</strong></td>
					<td width="63" align="center" ><strong>加入时间</strong></td>
					<td width="57" align="center" ><strong>点击数</strong></td>
					<td width="40" align="center" ><strong>已审核</strong></td>
					<td width="60" align="center" ><strong>操作</strong></td>
				</tr>
<%do while not rs.eof%>
				<tr class="tdbg"> 
					<td width="30" height="22" align="center" bgcolor="#C0C0C0"> 
					<input name='Bs_DownID'type='checkbox'onclick="unselectall()" id="Bs_DownID2" value='<%=cstr(rs("Bs_DownID"))%>'>
					</td>
					<td width="37" align="center" bgcolor="#E3E3E3"><%=rs("Bs_DownID")%></td>
					<td bgcolor="#E3E3E3"> <a href="../Bs_DownloadShow.asp?Bs_DownID=<%=rs("Bs_DownID")%>" 
					title="<%=replace(left(nohtml(rs("Bs_Content")),200),chr(34),"") & "……"%>">&nbsp;<%=rs("Bs_Title")%></a></td>
					<td width="63" align="center" bgcolor="#E3E3E3"><%= FormatDateTime(rs("Bs_UpdateTime"),2) %></td>
					<td width="57" align="center" bgcolor="#E3E3E3"><%=rs("Bs_Hits")%> </td>
					<td width="40" align="center" bgcolor="#E3E3E3">
<%if rs("Bs_Passed")=true then response.write "是" else response.write "否" end if%>
					</td>
					<td width="60" align="center" bgcolor="#E3E3E3"> 
<%
if session("purview")<=2 or PurviewChecked=True then
	If rs("Bs_Passed")=False Then 
%>
<a href="Bs_Download_Check_Set.asp?Action=Check&Bs_DownID=<%=rs("Bs_DownID")%>">通过审核</a> 
<% Else %>
<a href="Bs_Download_Check_Set.asp?Action=CancelCheck&Bs_DownID=<%=rs("Bs_DownID")%>">取消通过</a> 
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
					<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">选中本页显示的所有下载</td>
					<td><input name="submit" type='submit'
					value="<%if Bs_Passed="True" then response.write "取消"%>审核选定的下载" 
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
					<td height="30"> <strong>查找下载：</strong> 
					<input name="Bs_Title" type="text" class=smallInput id="Bs_Title3" size="20"> 
					<input name="Query" type="submit" id="Query" value="查 询">
					&nbsp;&nbsp;请输入下载名称。如果为空，则查找所有下载。 </td>
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
		'response.write "删 除 失 败 !<br>"
    else
        'response.write "删除成功！<br>"
    end if	
End sub
sub CancelCheckArticle(id)
    dim sql
    sql="update Bs_Download set Bs_Passed = False where Bs_DownID="&cstr(id)
    conn.execute sql
    if err.Number<>0 then
		err.clear
		'response.write "删 除 失 败 !<br>"
    else
        'response.write "删除成功！<br>"
    end if	
End sub
%>
