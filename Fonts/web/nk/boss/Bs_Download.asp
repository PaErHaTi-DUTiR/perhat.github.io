<!--#include file="bsconfig.asp"-->
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
dim totalPut,CurrentPage,TotalPages
dim i,j
dim Bs_BigClassName,Bs_SmallClassName
dim Bs_DownID,Bs_Title
dim sql,rs
dim PurviewChecked
dim strAdmin,arrAdmin

PurviewChecked=false
strFileName="Bs_Download.asp?Bs_BigClassName=" & Bs_BigClassName & "&Bs_SmallClassName=" & Bs_SmallClassName

Bs_DownID=Request("Bs_DownID")
Bs_BigClassName=Trim(request("Bs_BigClassName"))
Bs_SmallClassName=Trim(request("Bs_SmallClassName"))
Bs_Title=Trim(request("Bs_Title"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
'============
sql="select * from Bs_Download where Bs_DownID>0"
if session("purview")>4 then
	sql=sql & " and Editor='" & Session("Admin") & "'and Bs_Passed=false"
end if
if Bs_Title<>"" then
	sql=sql & " and Bs_Title like '%" & Bs_Title & "%'"
end if
if Bs_BigClassName<>"" then
	sql=sql & " and Bs_BigClassName='" & Bs_BigClassName & "'"
	if Bs_SmallClassName<>"" then
		sql=sql & " and Bs_SmallClassName='" & Bs_SmallClassName & "'"
	end if
end if
'============
sql=sql & " order by Bs_DownID desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>

<SCRIPT language=javascript>
function unselectall()
{
    if(document.del.chkAll.checked){
	document.del.chkAll.checked = document.del.chkAll.checked&0;
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
function ConfirmDel()
{
   if(confirm("确定要删除选中的下载吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}
</SCRIPT>

<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="660" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>下 载 管 理</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <table width="650" border="0" align="center" cellpadding="5" cellspacing="2" class="border">
        <tr class="title"> 
          <td bgcolor="#C0C0C0" height="25">|  
<% 'Start 主类菜单
dim sqlBigClass,sqlSmallClass,rsBigClass,rsSmallClass,sqlSpecial,rsSpecial
sqlBigClass="select * from Bs_DownBigClass"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1
if rsBigClass.eof then 
	response.Write("还没有任何栏目，请首先添加栏目。")
end if
do while not rsBigClass.eof
	if rsBigClass("Bs_BigClassName")=Bs_BigClassName then
		response.Write("<a href='Bs_Download.asp?Bs_BigClassName=" & rsBigClass("Bs_BigClassName") & "'><font color='red'>" & rsBigClass("Bs_BigClassName") & "</font></a> | ")
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
		response.Write("<a href='Bs_Download.asp?Bs_BigClassName=" & rsBigClass("Bs_BigClassName") & "'>" & rsBigClass("Bs_BigClassName") & "</a> | ")
	end if
	rsBigClass.movenext
loop
rsBigClass.close
set rsBigClass=nothing
'END 主类菜单
%>
          </td>
        </tr>
<% 'Start 小类菜单
if Bs_BigClassName<>"" then
	sqlSmallClass="select * from Bs_DownSmallClass where Bs_BigClassName='" & Bs_BigClassName & "'"
	Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
	rsSmallClass.open sqlSmallClass,conn,1,1
	if not (rsSmallClass.bof and rsSmallClass.eof) then
		response.write "<tr class='tdbg'><td bgcolor='#C0C0C0'>"
		do while not rsSmallClass.eof
			if rsSmallClass("Bs_SmallClassName")=Bs_SmallClassName then
				response.Write("&nbsp;<a href='Bs_Download.asp?Bs_BigClassName=" & rsSmallClass("Bs_BigClassName") & "&Bs_SmallClassName=" & rsSmallClass("Bs_SmallClassName") & "'><font color='red'>" & rsSmallClass("Bs_SmallClassName") & "</font></a>&nbsp;&nbsp;")
				if session("purview")=4 then
					strAdmin=rsSmallClass("Admin")
					if Instr(strAdmin,"|")>0 then
						arrAdmin=split(strAdmin)
						for i=0 to ubound(arrAdmin)
							if trim(arrAdmin(i))=session("Admin") then
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
				response.Write("&nbsp;<a href='Bs_Download.asp?Bs_BigClassName=" & rsSmallClass("Bs_BigClassName") & "&Bs_SmallClassName=" & rsSmallClass("Bs_SmallClassName") & "'>" & rsSmallClass("Bs_SmallClassName") & "</a>&nbsp;&nbsp;")
			end if
			rsSmallClass.movenext
		loop
		response.write "</td></tr>"
	end if
	rsSmallClass.close
	set rsSmallClass=nothing
end if
'END 小类菜单
%>
      </table>
      <form name="del" method="Post" action="Bs_Download_Del.asp" onsubmit="return ConfirmDel();">
        <table width="650" border="0" cellpadding="0" cellspacing="2">
          <tr> 
            <td height="25" bgcolor="#CCCCCC"><a href="Bs_Download.asp">&nbsp;下载管理</a>&gt;&gt;
<% 'Start 菜单树
if request.querystring="" then
	response.write "所有下载"
else
	if request("Query")<>"" then
		if Bs_Title<>"" then
			response.write "名称中含有“<font color=blue>" & Bs_Title & "</font>”的下载"
		else
			response.Write("所有下载")
		end if
 	else
		if Bs_BigClassName<>"" then
			response.write "<a href='Bs_Download.asp?Bs_BigClassName=" & Bs_BigClassName & "'>" & Bs_BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if Bs_SmallClassName<>"" then
				response.write "<a href='Bs_Download.asp?Bs_BigClassName=" & Bs_BigClassName & "&Bs_SmallClassName=" & Bs_SmallClassName & "'>" & Bs_SmallClassName & "</a>"
			else
				response.write "所有小类"
			end if
		end if
	end if
end if
'END 菜单树
%>
            </td>
            <td width="151" bgcolor="#CCCCCC">&nbsp;
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
			showpage strFileName,totalput,MaxPerPage,true,false,"个下载"
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rs.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rs.bookmark
				showContent
				showpage strFileName,totalput,MaxPerPage,true,false,"个下载"
			else
				currentPage=1
				showContent
				showpage strFileName,totalput,MaxPerPage,true,false,"个下载"
			end if
		end if
end if
%>

<%  
sub showContent
   	dim i
    i=0
%>
        <table class="border" border="0" cellspacing="2" width="650" cellpadding="0" style="word-break:break-all">
          <tr bgcolor="#C0C0C0" class="title"> 
            <td width="46" height="25" align="center"><strong>选中</strong></td>
            <td width="30"  height="25" align="center" bgcolor="#C0C0C0"><strong>ID</strong></td>
            <td width="350" align="center" bgcolor="#C0C0C0" ><strong>下载名称</strong></td>
            <td width="72" align="center" ><strong>加入时间</strong></td>
            <td width="72" align="center" ><strong>审核情况</strong></td>
            <td width="80" align="center" ><strong>操作</strong></td>
          </tr>
          <%do while not rs.eof%>
          <tr class="tdbg"> 
            <td width="46" height="22" align="center" bgcolor="#C0C0C0"> <input name='Bs_DownID'type='checkbox'onclick="unselectall()" id="Bs_DownID" value='<%=cstr(rs("Bs_DownID"))%>'> 
            </td>
            <td width="30" align="center" bgcolor="#E3E3E3"><%=rs("Bs_DownID")%></td>
            <td bgcolor="#E3E3E3"><a href="../Chinese/Bs_DownloadShow.asp?Bs_DownID=<%=rs("Bs_DownID")%>" target="_blank" 
						title="<%=replace(left(rs("Bs_Content"),200),"<br>","") & "……"%>">&nbsp;<%=rs("Bs_Title")%></a></td>
            <td width="72" align="center" bgcolor="#E3E3E3"><%= FormatDateTime(rs("Bs_UpDateTime"),2) %></td>
            <td width="72" align="center" bgcolor="#E3E3E3"> 
              <%if rs("Bs_Passed")=true then%>
              已审核 
              <%else%>
              <font color="#FF0000">未审核</font> 
              <%end if%>
            </td>
            <td width="80" align="center" bgcolor="#E3E3E3"> 
              <a href="Bs_Download_edit.asp?Bs_DownID=<%=rs("Bs_DownID")%>">修改</a> 
              <a href="Bs_Download_Del.asp?Bs_DownID=<%=rs("Bs_DownID")%>&Action=Del" onclick="return ConfirmDel();">删除</a> 
            </td>
          </tr>
<%
	i=i+1
		if i>=MaxPerPage then exit do
		rs.movenext
	loop
%>
        </table>
        <table width="650" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有下载</td>
            <td><input name="submit" type='submit'value='删除选定的下载'<%if session("purview")>=3 and session("purview")<=4 and PurviewChecked=False then response.write "disabled"%>> 
              <input name="Action" type="hidden" id="Action" value="Del"></td>
          </tr>
        </table>
        <%
   end sub 
%>
      </form>
      <br> 
      <table width="650" border="0" cellpadding="0" cellspacing="0" class="border">
        <tr class="tdbg">
          <form name="searchsoft" method="get" action="Bs_Download.asp">
            <td height="25"> <strong>查找下载：</strong> 
              <input name="Bs_Title" type="text" class=smallInput id="Title3" size="20" maxlength="50"> 
              <input name="Query" type="submit" id="Query" value="查 询">
              &nbsp;&nbsp;请输入下载名称。如果为空，则查找所有下载。</td>
          </form>
        </tr>
      </table><BR>
		</td>
	</tr>
</table>
<BR>
<%
rs.close
set rs=nothing  
call CloseConn()
%>
<%htmlend%>
