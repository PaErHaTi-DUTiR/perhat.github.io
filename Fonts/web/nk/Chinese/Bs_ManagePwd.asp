<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
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
dim Action,UserName
dim FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
UserName=trim(request("UserName"))
if Action="" and session("UserName")="" then
	response.redirect "Bs_Server.asp"
end if
if Action="Modify" and UserName<>"" then
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [Bs_User] where UserName='" & UserName & "'"
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
	else
		dim OldPassword,Password,PwdConfirm
		OldPassword=trim(request("OldPassword"))
		Password=trim(request("Password"))
		PwdConfirm=trim(request("PwdConfirm"))
		if OldPassword="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����������룡</li>"
		else
			if Instr(OldPassword,"=")>0 or Instr(OldPassword,"%")>0 or Instr(OldPassword,chr(32))>0 or Instr(OldPassword,"?")>0 or Instr(OldPassword,"&")>0 or Instr(OldPassword,";")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,"'")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,chr(34))>0 or Instr(OldPassword,chr(9))>0 or Instr(OldPassword,"��")>0 or Instr(OldPassword,"$")>0 then
				errmsg=errmsg+"<br><li>�������к��зǷ��ַ�</li>"
				founderr=true
			else
				if md5(OldPassword)<>rsUser("Password") then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>������ľ����벻�ԣ�û��Ȩ���޸ģ�</li>"
				end if
			end if
		end if
		if strLength(Password)>12 or strLength(Password)<6 then
			founderr=true
			errmsg=errmsg & "<br><li>������������(���ܴ���12С��6)��</li>"
		else
			if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
				errmsg=errmsg+"<br><li>�������к��зǷ��ַ�</li>"
				founderr=true
			end if
		end if
		if PwdConfirm="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>������ȷ�����룡</li>"
		else
			if PwdConfirm<>Password then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>ȷ�������������벻һ�£�</li>"
			end if
		end if
		if FoundErr<>true then
			rsUser("Password")=md5(Password)
			rsUser.update
		end if
	end if
	rsUser.close
	set rsUser=nothing
	if FoundErr=True then
		call WriteErrMsg()
	else
	    response.write"<SCRIPT language=JavaScript>alert('�ɹ��޸����룡');"
        response.write"javascript:history.go(-1)</SCRIPT>"
	end if
else
%>
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
<%end if%>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '����� %>

<!-- �޸��û����� -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
�޸��û�����
&nbsp;</SPAN></TD>
						<TD class='MC_Pt_TD4'><div align="right">
&nbsp;</div></TD>
						<TD class='MC_Pt_TD5'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR> 
			<FORM name="Form1" action="Bs_ManagePwd.asp" method="post">
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 class='MC_ServeNr_Title'>
					<TR><TD height=50><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
					<TR class="tdbg" > 
						<TD width="180" align="right"><b>�� �� ����</b></TD>
						<TD><%=session("UserName")%> <input name="UserName" type="hidden" value="<%=session("UserName")%>"></TD>
					</TR>
					<TR class="tdbg" > 
						<TD width="180" align="right"><B>�� �� �룺</B></TD>
						<TD> <INPUT   type=password maxLength=16 size=30 name=OldPassword> </TD>
					</TR>
					<TR class="tdbg" > 
						<TD width="180" align="right"><B>�� �� �룺</B></TD>
						<TD> <INPUT   type=password maxLength=16 size=30 name=Password> </TD>
					</TR>
					<TR class="tdbg" > 
						<TD width="180" align="right"><strong>ȷ�����룺</strong></TD>
						<TD> <INPUT name=PwdConfirm type=password id="PwdConfirm" size=30 maxLength=16> </TD>
					</TR>
					<TR align="center" class="tdbg" > 
						<TD height="50" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify"> 
						<input name=Submit type=submit id="Submit" value=" �� �� "> 	</TD>
					</TR>
					<TR><TD height=50><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
				</TABLE>
			</TD>
			</form>
		</TR>
<%
end if
%>
<!-- �޸��û����� End -->
		<TR> 
			<TD> 
				<TABLE cellSpacing=0 cellPadding=0 class='MC_bottom_title'>
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
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>