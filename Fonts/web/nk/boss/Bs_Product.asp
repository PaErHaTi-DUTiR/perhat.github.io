<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
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
dim totalPut,CurrentPage,TotalPages
dim i
dim ArticleID
dim Title
dim sql,rs
dim BigClassName,SmallClassName,SpecialName
dim PurviewChecked
dim strAdmin,arrAdmin
const MaxPerPage=20
PurviewChecked=false

Title=Trim(request("Title"))
ArticleID=Request("ArticleID")
BigClassName=Trim(request("BigClassName"))
SmallClassName=Trim(request("SmallClassName"))
SpecialName=trim(request("SpecialName"))
dim strFileName
strFileName="Bs_Product.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "&SpecialName=" & SpecialName

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sql="select * from Bs_Product where ArticleID>0"
if session("purview")>4 then
	sql=sql & " and Editor='" & Session("admin") & "'and Passed=false"
end if
if Title<>"" then
	sql=sql & " and title like '%" & Title & "%'"
end if
if BigClassName<>"" then
	sql=sql & " and BigClassName='" & BigClassName & "'"
	if SmallClassName<>"" then
		sql=sql & " and SmallClassName='" & SmallClassName & "'"
	end if
else
	if SpecialName<>"" then
		sql=sql & " and SpecialName='" & SpecialName & "'"
	end if
end if
sql=sql & " order by ArticleID desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<SCRIPT language=javascript>
function unselectall()
{
	if(document.del.chkAll.checked)
	{
		document.del.chkAll.checked = document.del.chkAll.checked;
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
   if(confirm("ȷ��Ҫɾ��ѡ�еĲ�Ʒ��һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;
}
</SCRIPT>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="610" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� Ʒ �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <table width="600" border="0" align="center" cellpadding="5" cellspacing="2" class="border">
        <tr class="title"> 
          <td bgcolor="#C0C0C0" height="25">|&nbsp; 
            <%
dim sqlBigClass,sqlSmallClass,rsBigClass,rsSmallClass,sqlSpecial,rsSpecial
sqlBigClass="select * from Bs_PrBigClasss"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1
if rsBigClass.eof then 
	response.Write("��û���κ���Ŀ�������������Ŀ��")
end if
do while not rsBigClass.eof
	if rsBigClass("BigClassName")=BigClassName then
		response.Write("<a href='Bs_Product.asp?BigClassName=" & rsBigClass("BigClassName") & "'><font color='red'>" & rsBigClass("BigClassName") & "</font></a> | ")
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
		response.Write("<a href='Bs_Product.asp?BigClassName=" & rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</a> | ")
	end if
	rsBigClass.movenext
loop
rsBigClass.close
set rsBigClass=nothing
%>
          </td>
        </tr>
        <%
if BigClassName<>"" then
	sqlSmallClass="select * from Bs_PrSmallClass where BigClassName='" & BigClassName & "'"
	Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
	rsSmallClass.open sqlSmallClass,conn,1,1
	if not (rsSmallClass.bof and rsSmallClass.eof) then
		response.write "<tr class='tdbg'><td bgcolor='#C0C0C0'>"
		do while not rsSmallClass.eof
			if rsSmallClass("SmallClassName")=SmallClassName then
				response.Write("&nbsp;<a href='Bs_Product.asp?BigClassName=" & rsSmallClass("BigClassName") & "&SmallClassName=" & rsSmallClass("SmallClassName") & "'><font color='red'>" & rsSmallClass("SmallClassName") & "</font></a>&nbsp;&nbsp;")
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
				response.Write("&nbsp;<a href='Bs_Product.asp?BigClassName=" & rsSmallClass("BigClassName") & "&SmallClassName=" & rsSmallClass("SmallClassName") & "'>" & rsSmallClass("SmallClassName") & "</a>&nbsp;&nbsp;")
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
      <form name="del" method="Post" action="Bs_Product_Del.asp" onsubmit="return ConfirmDel();">
        <table width="600" border="0" cellpadding="0" cellspacing="2">
          <tr> 
            <td height="25" bgcolor="#CCCCCC"><a href="Bs_Product.asp">&nbsp;��Ʒ����</a> 
              &gt;&gt; 
              <%
if request.querystring="" then
	response.write "���в�Ʒ"
else
	if request("Query")<>"" then
		if Title<>"" then
			response.write "�����к��С�<font color=blue>" & Title & "</font>���Ĳ�Ʒ"
		else
			response.Write("���в�Ʒ")
		end if
 	else
		if BigClassName<>"" then
			response.write "<a href='Bs_Product.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Bs_Product.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>"
			else
				response.write "����С��"
			end if
		end if
		if SpecialName<>"" then
			response.write "<font color=red>[ר��]</font> " & SpecialName
		end if
	end if
end if
%>
            </td>
            <td width="150" bgcolor="#CCCCCC">&nbsp;
            <%
  	if rs.eof and rs.bof then
		response.write "���ҵ� 0 ����Ʒ</td></tr></table>"
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
		response.Write "���ҵ� " & totalPut & " ����Ʒ"
%>
            </td>
          </tr>
        </table>
        <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"����Ʒ"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"����Ʒ"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"����Ʒ"
	    	end if
		end if
	end if
%>
        <%  
sub showContent()
   	dim i
    i=0
%>
        <table class="border" border="0" cellspacing="2" width="600" cellpadding="0" style="word-break:break-all">
          <tr bgcolor="#C0C0C0" class="title"> 
            <td width="47" height="25" align="center"><strong>ѡ��</strong></td>
            <td width="28"  height="25" align="center" bgcolor="#C0C0C0"><strong>ID</strong></td>
            <td width="82" align="center" bgcolor="#C0C0C0"><strong>��Ʒ���</strong></td>
            <td width="231" align="center" bgcolor="#C0C0C0" ><strong>��Ʒ����</strong></td>
            <td width="68" align="center" ><strong>����ʱ��</strong></td>
            <td width="68" align="center" ><strong>������</strong></td>
            <td width="80" align="center" ><strong>����</strong></td>
          </tr>
          <%do while not rs.eof%>
          <tr class="tdbg"> 
            <td width="47" height="22" align="center" bgcolor="#C0C0C0"> <input name='ArticleID'type='checkbox'onclick="unselectall()" id="ArticleID" value='<%=cstr(rs("articleID"))%>'> 
            </td>
            <td width="28" align="center" bgcolor="#E3E3E3"><%=rs("articleid")%></td>
            <td width="82" align="center" bgcolor="#E3E3E3"><%=rs("Product_Id")%></td>
            <td bgcolor="#E3E3E3"><a href='../Chinese/Bs_ProductShow.asp?ArticleID=<%=rs("articleid")%>'target='_blank'
						title='<%=replace(left(nohtml(rs("Content")),200),chr(34),"") & "����"%>'>&nbsp;<%=rs("title")%></a></td>
            <td width="68" align="center" bgcolor="#E3E3E3"><%= FormatDateTime(rs("UpdateTime"),2) %></td>
            <td width="68" align="center" bgcolor="#E3E3E3"> 
              <%if rs("Passed")=true then%>
              ����� 
              <%else%>
              <font color="#FF0000">δ���</font> 
              <%end if%>
            </td>
            <td width="80" align="center" bgcolor="#E3E3E3"> 
              <a href="Bs_Product_edit.asp?ArticleID=<%=rs("articleid")%>">�޸�</a> 
              <a href="Bs_Product_Del.asp?ArticleID=<%=rs("ArticleID")%>&Action=Del" onclick="return ConfirmDel();">ɾ��</a> 
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
            <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ��ʾ�����в�Ʒ</td>
            <td><input name="submit" type='submit'value='ɾ��ѡ���Ĳ�Ʒ'<%if session("purview")>=3 and session("purview")<=4 and PurviewChecked=False then response.write "disabled"%>> 
              <input name="Action" type="hidden" id="Action" value="Del"></td>
          </tr>
        </table>
<%
   end sub 
%>
      </form>
      <br> 
      <table width="600" border="0" cellpadding="0" cellspacing="0" class="border">
        <tr class="tdbg">
          <form name="searchsoft" method="get" action="Bs_Product.asp">
            <td height="25"> <strong>���Ҳ�Ʒ��</strong> 
              <input name="Title" type="text" class=smallInput id="Title3" size="20" maxlength="50"> 
              <input name="Query" type="submit" id="Query" value="�� ѯ">
              &nbsp;&nbsp;�������Ʒ���ơ����Ϊ�գ���������в�Ʒ��</td>
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
%>
