<!--#include file="bsconfig.asp"-->
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
dim Bs_BigClassID,Action,rs,NewBigClassName,OldBigClassName,FoundErr,ErrMsg
Bs_BigClassID=trim(Request("Bs_BigClassID"))
Action=trim(Request("Action"))
NewBigClassName=trim(Request("NewBigClassName"))
OldBigClassName=trim(Request("OldBigClassName"))

if Bs_BigClassID="" then
  response.Redirect("Bs_DownClass.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from Bs_DownBigClass where Bs_BigClassID=" & CLng(Bs_BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�����ش��಻���ڣ�</li>"
else
	if Action="Modify" then
		if NewBigClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>���ش���������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("Bs_BigClassName")=NewBigClassName
			rs("Admin")=Admin
    	 	rs.update
			rs.Close
	     	set rs=Nothing
			if NewBigClassName<>OldBigClassName then
				conn.execute "Update SmallClass set Bs_BigClassName='" & NewBigClassName & "'where Bs_BigClassName='" & OldBigClassName & "'"
				conn.execute "Update Product set Bs_BigClassName='" & NewBigClassName & "'where Bs_BigClassName='" & OldBigClassName & "'"
     		end if		
			call CloseConn()
     		Response.Redirect "Bs_DownClass.asp" 
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.Bs_BigClassName.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.Bs_BigClassName.focus();
    return false;
  }
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="360" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
			<form name="form1" method="post" action="Bs_DownBigClass_edit.asp">
			<table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
				<tr class="title"> 
					<td height="25" colspan="2" align="center" bgcolor="#DCDDDE"><strong>�޸Ĵ�������</strong></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" bgcolor="#DCDDDE" align=right><strong>����ID��</strong></td>
					<td width="204" bgcolor="#E3E3E3"> 
					<%=rs("Bs_BigClassID")%><input name="Bs_BigClassID" type="hidden" id="Bs_BigClassID" value="<%=rs("Bs_BigClassID")%>">
					<input name="OldBigClassName" type="hidden" id="OldBigClassName" value="<%=rs("Bs_BigClassName")%>"></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" bgcolor="#DCDDDE" align=right><strong>�������ƣ�</strong></td>
					<td bgcolor="#E3E3E3">
					<input name="NewBigClassName" type="text" id="NewBigClassName" value="<%=rs("Bs_BigClassName")%>" size="20" maxlength="30"></td>
				</tr>
				<tr class="tdbg"> 
					<td align="center" bgcolor="#DCDDDE">&nbsp;</td>
					<td align="center" bgcolor="#DCDDDE"><div align="left">
					<input name="Action" type="hidden" id="Action" value="Modify">
					<input name="Save" type="submit" id="Save" value=" �� �� ">
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
