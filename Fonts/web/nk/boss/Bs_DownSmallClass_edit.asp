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
dim Bs_SmallClassID,Action,Bs_BigClassName,Bs_SmallClassName,OldSmallClassName,rs,FoundErr,ErrMsg

Bs_SmallClassID=trim(Request("Bs_SmallClassID"))
Action=trim(Request("Action"))
Bs_BigClassName=trim(Request.form("Bs_BigClassName"))
Bs_SmallClassName=trim(Request.form("Bs_SmallClassName"))
OldSmallClassName=trim(request.form("OldSmallClassName"))

sqlBig="select * from Bs_DownBigClass where Bs_BigClassName='" & Bs_BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
rsBig.close

if Bs_SmallClassID="" then
	response.Redirect("Bs_DownClass.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from Bs_DownSmallClass where Bs_SmallClassID="&Bs_SmallClassID&"",conn,1,3
if rs.Bof or rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>������С�಻���ڣ�</li>"
else
	if Action="Modify" then
		if Bs_SmallClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����С��������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("Bs_SmallClassName")=Bs_SmallClassName
     		rs.update
			rs.Close
    	 	set rs=Nothing
			if Bs_SmallClassName<>OldSmallClassName then
				conn.execute "Update Product set Bs_SmallClassName='" & Bs_SmallClassName & "'where Bs_BigClassName='" & Bs_BigClassName &"'and Bs_SmallClassName='" & OldSmallClassName & "'"
	     	end if	
			call CloseConn()
    	 	Response.Redirect "Bs_DownClass.asp"
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="360" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
			<form name="form1" method="post" action="Bs_DownSmallClass_edit.asp">
			<table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
				<tr bgcolor="#999999" class="title"> 
					<td height="25" colspan="2" align="center"><strong>�޸�С������</strong></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>�������ࣺ</strong></td>
					<td width="204" bgcolor="#E3E3E3"><%=rs("Bs_BigClassName")%> 
					<input name="Bs_SmallClassID" type="hidden" id="Bs_SmallClassID" value="<%=rs("Bs_SmallClassID")%>"> 
					<input name="OldSmallClassName" type="hidden" id="OldSmallClassName" value="<%=rs("Bs_SmallClassName")%>"> 
					<input name="Bs_BigClassName" type="hidden" id="Bs_BigClassName" value="<%=rs("Bs_BigClassName")%>"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" bgcolor="#C0C0C0" align=right><strong>С�����ƣ�</strong></td>
					<td bgcolor="#E3E3E3">
					<input name="Bs_SmallClassName" type="text" id="Bs_SmallClassName" value="<%=rs("Bs_SmallClassName")%>" size="20" maxlength="30"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" align="center" bgcolor="#C0C0C0">&nbsp;</td>
					<td align="center" bgcolor="#E3E3E3"> <div align="left">
					<input name="Action" type="hidden" id="Action4" value="Modify">
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
