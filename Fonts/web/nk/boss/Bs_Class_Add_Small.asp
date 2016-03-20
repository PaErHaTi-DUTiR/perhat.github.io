<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->

<%

%>

<%
dim Action,BigClassName,EnBigClassName,SmallClassName,EnSmallClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))
EnSmallClassName=trim(request("EnSmallClassName"))

sqlBig="select * from Bs_PrBigClasss where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
EnBigClassName=rsBig("EnBigClassName")
rsBig.close

if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章大类名不能为空！</li>"
	end if
	if SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章小类名不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		if yzcv<>zcv then response.end
		rs.open "Select * from Bs_PrSmallClass Where BigClassName='" & BigClassName & "'AND SmallClassName='" & SmallClassName & "'",conn,1,3
		if not rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>“" & BigClassName & "”中已经存在文章小类“" & SmallClassName & "”！</li>"
		else
     		rs.addnew
				rs("BigClassName")=BigClassName
				rs("EnBigClassName")=EnBigClassName
    	 	rs("SmallClassName")=SmallClassName
    	 	rs("EnSmallClassName")=EnSmallClassName
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "Bs_Class.asp"  
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<script language="JavaScript" type="text/JavaScript">
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("请先添加大类名称！");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("小类名称不能为空！");
	document.form2.SmallClassName.focus();
	return false;
  }
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="360" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>产 品 类 别 设 置</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<BR>
			<form name="form2" method="post" action="Bs_Class_Add_Small.asp" onSubmit="return checkSmall()">
			<table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
				<tr bgcolor="#999999" class="title"> 
					<td height="25" colspan="2" align="center"><strong>添加小类</strong></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>所属大类：</strong></td>
					<td width="218" bgcolor="#E3E3E3"> 
<select name="BigClassName">
<%
dim rsBigClass
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * from Bs_PrBigClasss",conn,1,1
if rsBigClass.bof and rsBigClass.bof then
	response.write "<option>请先添加文章大类</option>"
else
	do while not rsBigClass.eof
		if rsBigClass("BigClassName")=BigClassName then
			response.write "<option value='"& rsBigClass("BigClassName") & "'selected>" & rsBigClass("BigClassName") & "</option>"
		else
			response.write "<option value='"& rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</option>"
		end if
		rsBigClass.movenext
	loop
end if
rsBigClass.close
set rsBigClass=nothing
%>
</select>
					</td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>小类名称：</strong></td>
					<td bgcolor="#E3E3E3"><input name="SmallClassName" type="text" size="20" maxlength="120"></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>English名称：</strong></td>
					<td bgcolor="#E3E3E3"><input name="EnSmallClassName" type="text" size="20" maxlength="120"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" align="center" bgcolor="#C0C0C0">&nbsp; </td>
					<td height="22" align="center" bgcolor="#E3E3E3"><div align="left"
><input name="Action" type="hidden" id="Action3" value="Add"><input name="Add" type="submit" value=" 添 加 "></div></td>
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
call CloseConn()
%>
