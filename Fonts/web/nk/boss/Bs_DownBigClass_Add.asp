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
dim Action,Bs_BigClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
Bs_BigClassName=trim(request("Bs_BigClassName"))
if Action="Add" then
	if Bs_BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>���ش���������Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From Bs_DownBigClass Where Bs_BigClassName='" & Bs_BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>���ش��ࡰ" & Bs_BigClassName & "���Ѿ����ڣ�</li>"
		else
    	 	rs.addnew
     		rs("Bs_BigClassName")=Bs_BigClassName
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
				<form name="form1" method="post" action="Bs_DownBigClass_Add.asp" onsubmit="return checkBig()">
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center" align=right><strong>��Ӵ���</strong></td>
          </tr>
          <tr bgcolor="#E3E3E3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#C0C0C0"> 
              <div align="right" align=right><strong>�������ƣ�</strong></div></td>
            <td width="218" bgcolor="#E3E3E3"> 
              <input name="Bs_BigClassName" type="text" size="20" maxlength="30">
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
