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
dim Action,BigClassName,EnBigClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
EnBigClassName=trim(request("EnBigClassName"))
if yzcv<>zcv then response.end
if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>���´���������Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * from Bs_PrBigClasss Where BigClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>���´��ࡰ" & BigClassName & "���Ѿ����ڣ�</li>"
		else
    	 	rs.addnew
     		rs("BigClassName")=BigClassName
     		rs("EnBigClassName")=EnBigClassName
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
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.BigClassName.focus();
    return false;
  }
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="360" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� Ʒ �� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
				<form name="form1" method="post" action="Bs_Class_Add_Big.asp" onsubmit="return checkBig()">
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center" align=right><strong>��Ӵ���</strong></td>
          </tr>
          <tr bgcolor="#E3E3E3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#C0C0C0"> 
              <div align="right" align=right><strong>�������ƣ�</strong></div></td>
            <td width="218" bgcolor="#E3E3E3"> 
              <input name="BigClassName" type="text" size="20" maxlength="30">
            </td>
          </tr>
          <tr bgcolor="#E3E3E3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#C0C0C0"> 
              <div align="right" align=right><strong>English���ƣ�</strong></div></td>
            <td width="218" bgcolor="#E3E3E3"> 
              <input name="EnBigClassName" type="text" size="20" maxlength="30">
            </td>
          </tr>
          <tr bgcolor="#C0C0C0" class="tdbg"> 
            <td height="22" align="center" bgcolor="#C0C0C0">&nbsp; </td>
            <td height="22" align="center" bgcolor="#E3E3E3"> 
              <div align="left">
                <input name="Action" type="hidden" id="Action" value="Add">
                <input name="Add" type="submit" value=" �� �� ">
              </div>
						</td>
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
