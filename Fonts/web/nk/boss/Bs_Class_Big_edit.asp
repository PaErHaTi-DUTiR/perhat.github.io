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
dim BigClassID,Action,rs,NewBigClassName,EnNewBigClassName,OldBigClassName,EnOldBigClassName,FoundErr,ErrMsg
BigClassID=trim(Request("BigClassID"))
Action=trim(Request("Action"))
NewBigClassName=trim(Request("NewBigClassName"))
EnNewBigClassName=trim(Request("EnNewBigClassName"))
OldBigClassName=trim(Request("OldBigClassName"))
EnOldBigClassName=trim(Request("EnOldBigClassName"))

if BigClassID="" then
  response.Redirect("Bs_Class.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from Bs_PrBigClasss where BigClassID=" & CLng(BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>此产品大类不存在！</li>"
else
	if Action="Modify" then
		if NewBigClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>产品大类名不能为空！</li>"
		end if
		if FoundErr<>True then
			rs("BigClassName")=NewBigClassName
			rs("EnBigClassName")=EnNewBigClassName
			rs("Admin")=Admin
    	rs.update
			rs.Close
	    set rs=Nothing
			if NewBigClassName<>OldBigClassName or EnNewBigClassName<>EnOldBigClassName then
				conn.execute "Update Bs_PrSmallClass set BigClassName='" & NewBigClassName & "'where BigClassName='" & OldBigClassName & "'"
				conn.execute "Update Bs_Product set BigClassName='" & NewBigClassName & "'where BigClassName='" & OldBigClassName & "'"

				conn.execute "Update Bs_PrSmallClass set EnBigClassName='" & EnNewBigClassName & "'where EnBigClassName='" & EnOldBigClassName & "'"
				conn.execute "Update Bs_Product set EnBigClassName='" & EnNewBigClassName & "'where EnBigClassName='" & EnOldBigClassName & "'"
     	end if		
			call CloseConn()
     		Response.Redirect "Bs_Class.asp" 
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.BigClassName.focus();
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
      <br>
			<form name="form1" method="post" action="Bs_Class_Big_edit.asp">
			<table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
				<tr class="title"> 
					<td height="25" colspan="2" align="center" bgcolor="#999999"><strong>修改大类名称</strong></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" bgcolor="#C0C0C0" align=right><strong>大类ID：</strong></td>
					<td width="204" bgcolor="#E3E3E3"> 
					<%=rs("BigClassID")%><input name="BigClassID" type="hidden" id="BigClassID" value="<%=rs("BigClassID")%>">
					<input name="OldBigClassName" type="hidden" id="OldBigClassName" value="<%=rs("BigClassName")%>">
					<input name="EnOldBigClassName" type="hidden" id="EnOldBigClassName" value="<%=rs("EnBigClassName")%>"></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" bgcolor="#C0C0C0" align=right><strong>大类名称：</strong></td>
					<td bgcolor="#E3E3E3">
					<input name="NewBigClassName" type="text" id="NewBigClassName" value="<%=rs("BigClassName")%>" size="20" maxlength="30"></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" bgcolor="#C0C0C0" align=right><strong>English名称：</strong></td>
					<td bgcolor="#E3E3E3">
					<input name="EnNewBigClassName" type="text" id="EnNewBigClassName" value="<%=rs("EnBigClassName")%>" size="20" maxlength="30"></td>
				</tr>
				<tr class="tdbg"> 
					<td align="center" bgcolor="#C0C0C0">&nbsp;</td>
					<td align="center" bgcolor="#E3E3E3"><div align="left">
					<input name="Action" type="hidden" id="Action" value="Modify">
					<input name="Save" type="submit" id="Save" value=" 修 改 ">
					</div></td>
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
end if
rs.close
set rs=nothing
call CloseConn()
%>
