<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
'=========================================================
'
'��Ʒ���ƣ���������Ƽ� ��˾(��ҵ)��վ����ϵͳ����ƣ�liangjing.net��
'��Ȩ���У�liangjing.net
'����������liangjing�����Ŷ�
'��ϵ ��ʽ��QQ ��2113269
'Copyright 2006-2008 liangjing.net - All Rights Reserved. 
'
'========================================================
%>
<%
dim SmallClassID,Action,BigClassName,EnBigClassName, SmallClassName,EnSmallClassName, OldSmallClassName,EnOldSmallClassName,rs,FoundErr,ErrMsg

SmallClassID=trim(Request("SmallClassID"))
Action=trim(Request("Action"))
BigClassName=trim(Request.form("BigClassName"))
SmallClassName=trim(Request.form("SmallClassName"))
EnSmallClassName=trim(Request.form("EnSmallClassName"))
OldSmallClassName=trim(request.form("OldSmallClassName"))
EnOldSmallClassName=trim(request.form("EnOldSmallClassName"))

sqlBig="select * from Bs_PrBigClasss where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
EnBigClassName=rsBig("EnBigClassName")
rsBig.close

if SmallClassID="" then
	response.Redirect("Bs_Class.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from Bs_PrSmallClass where SmallClassID="&SmallClassID&"",conn,1,3
if rs.Bof or rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>������С�಻���ڣ�</li>"
else
	if Action="Modify" then
		if SmallClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����С��������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("SmallClassName")=SmallClassName
			rs("EnSmallClassName")=EnSmallClassName
     	rs.update
			rs.Close
   	 	set rs=Nothing
			if SmallClassName<>OldSmallClassName or EnSmallClassName<>EnOldSmallClassName then
				conn.execute "Update Product set SmallClassName='" & SmallClassName & "'where BigClassName='" & BigClassName & "'and SmallClassName='" & OldSmallClassName & "'"
				conn.execute "Update Product set EnSmallClassName='" & EnSmallClassName & "'where EnBigClassName='" & EnBigClassName & "'and EnSmallClassName='" & EnOldSmallClassName & "'"
	   	end if	
			call CloseConn()
    	 	Response.Redirect "Bs_Class.asp"
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
		<td class="a1" height="25" align="center"><strong>�� Ʒ �� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
			<form name="form1" method="post" action="Bs_Class_Small_edit.asp">
			<table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
				<tr bgcolor="#999999" class="title"> 
					<td height="25" colspan="2" align="center"><strong>�޸�С������</strong></td>
				</tr>
				<tr class="tdbg"> 
					<td width="126" height="22" bgcolor="#C0C0C0" align=right><strong>�������ࣺ</strong></td>
					<td width="204" bgcolor="#E3E3E3"><%=rs("BigClassName")%> 
					<input name="SmallClassID" type="hidden" id="SmallClassID" value="<%=rs("SmallClassID")%>"> 
					<input name="OldSmallClassName" type="hidden" id="OldSmallClassName" value="<%=rs("SmallClassName")%>"> 
					<input name="EnOldSmallClassName" type="hidden" id="EnOldSmallClassName" value="<%=rs("EnSmallClassName")%>"> 
					<input name="BigClassName" type="hidden" id="BigClassName" value="<%=rs("BigClassName")%>"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" bgcolor="#C0C0C0" align=right><strong>С�����ƣ�</strong></td>
					<td bgcolor="#E3E3E3">
					<input name="SmallClassName" type="text" id="SmallClassName" value="<%=rs("SmallClassName")%>" size="20" maxlength="120"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" bgcolor="#C0C0C0" align=right><strong>English���ƣ�</strong></td>
					<td bgcolor="#E3E3E3">
					<input name="EnSmallClassName" type="text" id="EnSmallClassName" value="<%=rs("EnSmallClassName")%>" size="20" maxlength="120"></td>
				</tr>
				<tr class="tdbg"> 
					<td height="22" align="center" bgcolor="#C0C0C0">&nbsp;</td>
					<td align="center" bgcolor="#E3E3E3"> <div align="left">
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
