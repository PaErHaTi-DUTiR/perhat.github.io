<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
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
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>����û�</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <form name="adduser" method="post" action="Bs_Mail_add_user.asp">
        <table border="0" cellspacing="2" cellpadding="0" width="300">
          <tr> 
            <td height="25" bgcolor="#C0C0C0"><b><font color="#000000">&nbsp;�ֶ�����ʼ���ַ</font></b></td>
          </tr>
          <tr> 
            <td bgcolor="#E3E3E3" align="center" height="95"> 
              <table border="0" cellspacing="0" cellpadding="4">
                <tr> 
                  <td align="right"><font color="#000000">Email��ַ</font></td>
                  <td valign="top"> 
                    <input type="text" name="email" maxlength="50" class="bk">
                  </td>
                </tr>
                <tr> 
                  <td align="center" colspan="2" height="35"><font color=#000000>������д��������</font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="25" align="center" bgcolor="#C0C0C0"> 
              <input type="submit" name="Submit" value="���" class="Tips_bo">
              <input type="hidden" name="action" value="adduser">
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
if request("action")<>"adduser" then
response.end
else
email=request("email")
mailnow=now()
conn.execute "insert into email (dateandtime,email) values ('"&mailnow&"','"&email&"')"
response.write"<SCRIPT language=JavaScript>alert('�ʺ���ӳɹ�������Լ�����ӣ�');"
response.write "</SCRIPT>"
end if
%>
