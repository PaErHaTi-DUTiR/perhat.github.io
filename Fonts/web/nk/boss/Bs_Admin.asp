<!--#include file="bsconfig.asp"-->

<%
'������������������������������������������������������������
'�������������� COPYRIGHT ������������ ��
'���������ƣ���ҵ��վ����ϵͳMac3.0  (ASBDcorpweb Mac3.0)  �� 
'����Ȩ���У�WWW.ASBD.CN  BBS.ASBD.CN                      ��
'������������amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  �� 
'��Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ��
'������������������������������������������������������������
%>

<%
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_Admin"
rs.open sqltext,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<script language="javascript">
function confirmdel(id){
if (confirm("���Ҫɾ���˹���Ա�ʺ�?"))
window.location.href="Bs_Admin_Del.asp?id="+id+"  "   }
</script>

<SCRIPT language=javascript id=clientEventHandlersJS>
<!--
function form1_onsubmit()
{
if (document.FORM1.pwd1.value!=document.FORM1.pwd2.value)
{
alert ("��ȷ���������롣");
document.FORM1.pwd1.value='';
document.FORM1.pwd2.value='';
document.FORM1.pwd1.focus();
return false;
}
}


//-->
</SCRIPT>
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>��ӹ���Ա</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <FORM language=javascript name=FORM1  onsubmit="return form1_onsubmit()" action="Bs_Admin_Add.asp" method=post>
			<table width="50%" border="0" align="center" cellpadding="0" cellspacing="2">
				<tr> 
					<td height="25" colspan="2">&nbsp;</td>
				</tr>
				<tr> 
					<td width="29%" height="22"> <div align="right">����Ա�ʺţ�</div></td>
					<td width="71%"><input type="text" name="UserName" size="16" maxlength="20"></td>
				</tr>
				<tr> 
					<td height="22"> <div align="right">��  �ݣ�</div></td>
					<td><input type="text" name="realname" size="16" maxlength="20"></td>
				</tr>
				<tr> 
					<td height="22"> <div align="right">����Ա���룺</div></td>
					<td><input type="text" name="pwd1" size="16"></td>
				</tr>
				<tr> 
					<td height="22"> <div align="right">����ȷ�ϣ�</div></td>
					<td><input type="text" name="pwd2" size="16"></td>
				</tr>
				<tr> 
					<td height="22" colspan="2"><div align="center"><INPUT type=submit value='ȷ�����'name=Submit2></div></td>
				</tr>
			</table>
      </form>
		</td>
	</tr>
</table>
<BR>
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>����Ա�ʺŹ���</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<FORM language=javascript action="Bs_Admin_Add.asp" method=post>
			<table width="550" border="0" align="center" cellpadding="0" cellspacing="2">
				<tr> 
					<td width="20%" height="25" bgcolor="#C0C0C0"> <div align="center">����Ա�ʺ�</div></td>
					<td width="20%" bgcolor="#C0C0C0"> <div align="center">��  ��</div></td>
					<td width="20%" bgcolor="#C0C0C0"> <div align="center">����Ա����</div></td>
					<td width="20%" bgcolor="#C0C0C0"> <div align="center">����</div></td>
					<td width="20%" bgcolor="#C0C0C0"> <div align="center">ɾ��</div></td>
				</tr>
<%
if not rs.eof then
do while not rs.eof
%>
				<tr> 
					<td height="22" bgcolor="#E3E3E3"> <div align="center"><%=rs("UserName")%></div></td>
					<td bgcolor="#E3E3E3"><div align="center"><%=rs("realname")%></div></td>
					<td bgcolor="#E3E3E3"><div align="center"><%=decrypt(rs("PassWord"))%></div></td>
					<td bgcolor="#E3E3E3"><div align="center"><%response.write "<a href='Bs_AdminPass_Edit.asp?ID="&rs("Id")&"' >�޸�����</a>"%></div></td>
					<td bgcolor="#E3E3E3"><div align="center"><%response.write "<a href='javascript:confirmdel(" & rs("Id") & ")'>ɾ��</a>"%></div></td>
				</tr>
<%
rs.movenext
loop
end if
%>
<%
rs.close
conn.close
%>
			</table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
