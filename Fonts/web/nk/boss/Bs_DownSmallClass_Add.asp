<!--#include file="bsconfig.asp"-->
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
dim Action,Bs_BigClassName,Bs_SmallClassName,sqlBig,rsBig,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
Bs_BigClassName=trim(request("Bs_BigClassName"))
Bs_SmallClassName=trim(request("Bs_SmallClassName"))

sqlBig="select * from Bs_DownBigClass where Bs_BigClassName='" & Bs_BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
rsBig.close

if Action="Add" then
	if Bs_BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>���ش���������Ϊ�գ�</li>"
	end if
	if Bs_SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����С��������Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From Bs_DownSmallClass Where Bs_BigClassName='" & Bs_BigClassName & "'AND Bs_SmallClassName='" & Bs_SmallClassName & "'",conn,1,3
		if not rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��" & Bs_BigClassName & "�����Ѿ���������С�ࡰ" & Bs_SmallClassName & "����</li>"
		else
     		rs.addnew
				rs("Bs_BigClassName")=Bs_BigClassName
    	 	rs("Bs_SmallClassName")=Bs_SmallClassName
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "Bs_DownClass.asp"  
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
  if (document.form2.Bs_BigClassName.value=="")
  {
    alert("������Ӵ������ƣ�");
	document.form1.Bs_BigClassName.focus();
	return false;
  }

  if (document.form2.Bs_SmallClassName.value=="")
  {
    alert("С�����Ʋ���Ϊ�գ�");
	document.form2.Bs_SmallClassName.focus();
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
			<BR>
			<form name="form2" method="post" action="Bs_DownSmallClass_Add.asp" onsubmit="return checkSmall()">
			<table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
				<tr bgcolor="#999999" class="title"> 
					<td height="25" colspan="2" align="center"><strong>���С��</strong></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>�������ࣺ</strong></td>
					<td width="218" bgcolor="#E3E3E3"> 
<select name="Bs_BigClassName">
<%
dim rsBigClass
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * From Bs_DownBigClass",conn,1,1
if yzcv<>zcv then response.end
if rsBigClass.bof and rsBigClass.bof then
	response.write "<option>����������´���</option>"
else
	do while not rsBigClass.eof
		if rsBigClass("Bs_BigClassName")=Bs_BigClassName then
			response.write "<option value='"& rsBigClass("Bs_BigClassName") & "'selected>" & rsBigClass("Bs_BigClassName") & "</option>"
		else
			response.write "<option value='"& rsBigClass("Bs_BigClassName") & "'>" & rsBigClass("Bs_BigClassName") & "</option>"
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
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>С�����ƣ�</strong></td>
					<td bgcolor="#E3E3E3"><input name="Bs_SmallClassName" type="text" size="20" maxlength="30"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" align="center" bgcolor="#C0C0C0">&nbsp; </td>
					<td height="22" align="center" bgcolor="#E3E3E3"><div align="left"
><input name="Action" type="hidden" id="Action3" value="Add"><input name="Add" type="submit" value=" �� �� "></div></td>
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
