<!--#include file="bsconfig.asp"-->
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
dim sql,rsBigClass,rsSmallClass,ErrMsg
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * from Bs_PrBigClasss",conn,1,3
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.BigClassName.focus();
    return false;
  }
}
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("������Ӵ������ƣ�");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("С�����Ʋ���Ϊ�գ�");
	document.form2.SmallClassName.focus();
	return false;
  }
}
function ConfirmDelBig()
{
   if(confirm("ȷ��Ҫɾ���˴�����ɾ���˴���ͬʱ��ɾ����������С�࣬���Ҳ��ָܻ���"))
     return true;
   else
     return false;
	 
}

function ConfirmDelSmall()
{
   if(confirm("ȷ��Ҫɾ����С����һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;
	 
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="510" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� Ʒ �� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>����ѡ�<a href="Bs_Class_Add_Big.asp">��Ӳ�Ʒ����</a><br>
      <br>
      <table width="500" border="0" cellpadding="0" cellspacing="2" class="border">
        <tr bgcolor="#999999" class="title"> 
          <td width="290" height="25" align="center"><strong>��Ŀ����&nbsp;[English]</strong></td>
          <td width="210" height="25" align="center"><strong>����ѡ��</strong></td>
        </tr>
<%
do while not rsBigClass.eof
%>
        <tr bgcolor="#E3E3E3" class="tdbg"> 
          <td width="290" height="22" bgcolor="#C0C0C0"><img src="../Images/tree_folder4.gif" width="15" height="15"><%=rsBigClass("BigClassName")%>&nbsp;[<%=rsBigClass("EnBigClassName")%>]</td>
          <td align="center" bgcolor="#C0C0C0"><a href="Bs_Class_Add_Small.asp?BigClassName=<%=rsBigClass("BigClassName")%>">�������Ŀ</a> 
            | <a href="Bs_Class_Big_edit.asp?BigClassID=<%=rsBigClass("BigClassID")%>">�޸�</a> 
            | <a href="Bs_Class_Del_Big.asp?BigClassName=<%=rsBigClass("BigClassName")%>" onClick="return ConfirmDelBig();">ɾ��</a></td>
        </tr>
<%
set rsSmallClass=server.CreateObject("adodb.recordset")
rsSmallClass.open "Select * from Bs_PrSmallClass Where BigClassName='" & rsBigClass("BigClassName") & "'",conn,1,3
if not(rsSmallClass.bof and rsSmallClass.eof) then
do while not rsSmallClass.eof
%>
        <tr bgcolor="#E3E3E3" class="tdbg"> 
          <td width="290" height="22">&nbsp;&nbsp;<img src="../Images/tree_folder3.gif" width="15" height="15"><%=rsSmallClass("SmallClassName")%>&nbsp;[<%=rsSmallClass("EnSmallClassName")%>]</td>
          <td align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <a href="Bs_Class_Small_edit.asp?SmallClassID=<%=rsSmallClass("SmallClassID")%>">�޸�</a> 
            | <a href="Bs_Class_Del_Small.asp?SmallClassID=<%=rsSmallClass("SmallClassID")%>" onClick="return ConfirmDelSmall();">ɾ��</a></td>
        </tr>
<%
rsSmallClass.movenext
loop
end if
rsSmallClass.close
set rsSmallClass=nothing	
rsBigClass.movenext
loop
%>
      </table><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>

<%
rsBigClass.close
set rsBigClass=nothing
call CloseConn()
%>
