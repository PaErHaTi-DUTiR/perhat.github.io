<!--#include file="Include/Bs_System.asp"-->
<!--#include file="Include/Bs_StrLeft.asp"-->
<!--#include file="../Inc/md5.asp"-->
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
'����Ķ����������д���
'ShowSmallClassType=ShowSmallClassType_Default
'MaxPerPage=MaxPerPage_Default
'strFileName="Default.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "&SpecialName=" & SpecialName

dim UserName,Password,PwdConfirm,Question,Answer,Sex,Email,HomePage,Comane,Add,Somane,Zip,Phone,Fox
UserName=trim(request("UserName"))
Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
Question=trim(request("Question"))
Answer=trim(request("Answer"))
Sex=trim(Request("Sex"))
Email=trim(request("Email"))
HomePage=trim(request("HomePage"))
Comane=trim(request("Comane"))
Add=trim(request("Add"))
Somane=trim(request("Somane"))
Zip=trim(request("Zip"))
Phone=trim(request("Phone"))
Fox=trim(request("Fox"))
if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>�������û���(���ܴ���14С��4)</li>"
else
  	if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
		errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>����������(���ܴ���12С��6)</li>"
else
	if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>�����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if PwdConfirm="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ȷ������(���ܴ���12С��6)</li>"
else
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>�����ȷ�����벻һ��</li>"
	end if
end if
if Question="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
end if
if Answer="" then
	founderr=true
	errmsg=errmsg & "<br><li>����𰸲���Ϊ��</li>"
end if
if Sex="" then
	founderr=true
	errmsg=errmsg & "<br><li>�Ա���Ϊ��</li>"
else
	sex=cint(sex)
	if Sex<>0 and Sex<>1 then
		Sex=1
	end if
end if
if Email="" then
	founderr=true
	errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
else
	if IsValidEmail(Email)=false then
		errmsg=errmsg & "<br><li>����Email�д���</li>"
   		founderr=true
	end if
end if
if Add="" then
	founderr=true
	errmsg=errmsg & "<br><li>�ջ���ַ����Ϊ��</li>"
end if
if Zip="" then
	founderr=true
	errmsg=errmsg & "<br><li>�������벻��Ϊ��</li>"
end if
if Phone="" then
	founderr=true
	errmsg=errmsg & "<br><li>��ϵ�绰����Ϊ��</li>"
end if

if founderr=false then
	dim sqlReg,rsReg
	sqlReg="select * from [Bs_User] where UserName='" & Username & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>��ע����û��Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
	else
		rsReg.addnew
		rsReg("UserName")=UserName
		rsReg("Password")=md5(Password)
		rsReg("Question")=Question
		rsReg("Answer")=md5(Answer)
		rsReg("Sex")=Sex
		rsReg("Email")=Email
		rsReg("HomePage")=HomePage
		rsReg("Comane")=Comane
		rsReg("Add")=Add
		rsReg("Somane")=Somane
		rsReg("Zip")=Zip
		rsReg("Phone")=Phone
		rsReg("Fox")=Fox
		rsReg.update
		founderr=false
	end if
	rsReg.close
	set rsReg=nothing
end if		
%>
<BODY bgColor=#666666 leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="Include/Bs_Head.asp" -->
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR> 
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Member() %>
			</td>
		</tr>
	</table> 
</TD>

<%if strSkins <= 200 and strSkins >= 1 then%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<%end if%><TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '����� %>

<!-- �û�ע�� -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
��Ա����
&nbsp;</SPAN></TD>
						<TD class='MC_Pt_TD4'><div align="right">
&nbsp;</div></TD>
						<TD class='MC_Pt_TD5'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_ServeNr_Title'>
					<TR> 
						<TD height=1>
<%
if founderr=false then
call RegSuccess()
else
call WriteErrMsg2()
end if
%> 
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<!-- �û�ע�� END -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_bottom_title'>
					<tr> 
						<td>
						</td>
					</tr>
					<tr><td class='MC_bottom_Td2'></td></tr>
				</table>
			</TD>
		</TR>
		<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	</TABLE>
</TD>
<%if strSkins <= 200 and strSkins >= 1 then%>

</TR>
</TABLE>
<!--#include file="Include/Bs_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>
<%
sub WriteErrMsg2()
    response.write "<table align='center' width='300'border='0'cellpadding='2'cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center'height='15'>�������µ�ԭ����ע���û���</td></tr>"
    response.write "<tr class='tdbg'><td align='left'height='100'><br>" & errmsg & "<p align='center'>��<a href='javascript:onclick=history.go(-1)'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<table align='center' width='300'border='0'cellpadding='2'cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center'height='15'>�ɹ�ע���û���</td></tr>"
    response.write "<tr class='tdbg'><td align='left'height='100'><br>��ע����û�����" & UserName & "<p align='center'>��<a href='javascript:onclick=window.close()'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub
%>