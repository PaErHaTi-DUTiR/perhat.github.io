<!--#include file="bsconfig.asp"-->
<!--#include file="../inc/md5.asp"-->
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
dim UserID,Action,FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
UserID=trim(request("UserID"))
if UserID="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [Bs_User] where UserID=" & Clng(UserID)
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
	else
		if Action="Modify" then
			dim UserName,Password,Question,Answer,Sex,Email,Homepage,LockUser,Comane,Add,Somane,Zip,Phone,Fox
			UserName=trim(request("UserName"))
			Password=trim(request("Password"))
			Question=trim(request("Question"))
			Answer=trim(request("Answer"))
			Sex=trim(Request("Sex"))
			Email=trim(request("Email"))
			Homepage=trim(request("Homepage"))
			Comane=trim(request("Comane"))
			Add=trim(request("Add"))
			Somane=trim(request("Somane"))
			Zip=trim(request("Zip"))
			Phone=trim(request("Phone"))
			Fox=trim(request("Fox"))
			LockUser=trim(request("LockUser"))
			if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
				founderr=true
				errmsg=errmsg & "<br><li>�������û���(���ܴ���14С��4)</li>"
			else
  				if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
					errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
					founderr=true
				else
					dim sqlReg,rsReg
					sqlReg="select * from [Bs_User] where UserName='" & Username & "'and UserID<>" & UserID
					set rsReg=server.createobject("adodb.recordset")
					rsReg.open sqlReg,conn,1,1
					if not(rsReg.bof and rsReg.eof) then
						founderr=true
						errmsg=errmsg & "<br><li>�û����Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
					end if
					rsReg.Close
					set rsReg=nothing
				end if
			end if
			if Password<>"" then
				if strLength(Password)>12 or strLength(Password)<6 then
					founderr=true
					errmsg=errmsg & "<br><li>����������(���ܴ���12С��6)���粻���޸ģ������գ�</li>"
				else
					if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
						errmsg=errmsg+"<br><li>�����к��зǷ��ַ�</li>"
						founderr=true
					end if
				end if
			end if
			if Question="" then
				founderr=true
				errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
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
			if LockUser="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�û�״̬����Ϊ�գ�</li>"
			end if
			if FoundErr<>true then
				rsUser("UserName")=UserName
				if Password<>"" then
					rsUser("Password")=md5(Password)
				end if
				rsUser("Question")=Question
				if Answer<>"" then
					rsUser("Answer")=md5(Answer)
				end if
				rsUser("Sex")=Sex
				rsUser("Email")=Email
				rsUser("HomePage")=HomePage
				rsUser("Comane")=Comane
				rsUser("Add")=Add
				rsUser("Somane")=Somane
				rsUser("Zip")=Zip
				rsUser("Phone")=Phone
				rsUser("Fox")=Fox
				rsUser("LockUser")=LockUser
				rsUser.update
				rsUser.Close
				set rsUser=nothing
				call CloseConn() 
				response.redirect "Bs_User.asp"
			end if
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�޸�ע���û���Ϣ</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<FORM name="Form1" action="Bs_User_edit.asp" method="post">
        <table width=500 border=0 align="center" cellpadding=2 cellspacing=2 class='border'>
          <TR align=center class='title'> 
            <TD height=20 colSpan=2><font class=en><b></b></font></TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><b>�� �� ����</b></TD>
            <TD> <INPUT name=UserName value="<%=rsUser("UserName")%>" size=30   maxLength=14></TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><B>����(����6λ)��</B></TD>
            <TD> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">��������޸ģ�������</font> 
            </TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>�������⣺</strong></TD>
            <TD> <INPUT name="Question"   type=text value="<%=rsUser("Question")%>" size=30> 
            </TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>����𰸣�</strong></TD>
            <TD> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">��������޸ģ�������</font></TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>�Ա�</strong></TD>
            <TD> <INPUT type=radio value="1" name=sex <%if rsUser("Sex")=1 then response.write "CHECKED"%>>
              �� &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if rsUser("Sex")=0 then response.write "CHECKED"%>>
              Ů</TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>Email��ַ��</strong></TD>
            <TD> <INPUT name=Email value="<%=rsUser("Email")%>" size=30   maxLength=50> 
            </TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>��ҳ��</strong></TD>
            <TD> <INPUT   maxLength=100 size=30 name=homepage value="<%=rsUser("HomePage")%>"></TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>��˾���ƣ�</strong></TD>
            <TD> <INPUT name=Comane value="<%=rsUser("Comane")%>" size=30 maxLength=20></TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>�ջ���ַ��</strong></TD>
            <TD> <INPUT name=Add value="<%=rsUser("Add")%>" size=30 maxLength=50></TD>
          </TR>
          <TR class="tdbg" > 
            <TD align="right"><strong>�ջ��ˣ�</strong></TD>
            <TD><INPUT name=Somane value="<%=rsUser("Somane")%>" size=30 maxLength=50></TD>
          </TR>
          <TR class="tdbg" > 
            <TD align="right"><strong>�������룺</strong></TD>
            <TD><INPUT name=Zip value="<%=rsUser("Zip")%>" size=30 maxLength=50></TD>
          </TR>
          <TR class="tdbg" > 
            <TD align="right"><strong>��ϵ�绰��</strong><br></TD>
            <TD><INPUT name=Phone value="<%=rsUser("Phone")%>" size=30 maxLength=50></TD>
          </TR>
          <TR class="tdbg" > 
            <TD align="right"><strong>�� �棺</strong></TD>
            <TD><INPUT name=Fox value="<%=rsUser("Fox")%>" size=30 maxLength=50></TD>
          </TR>
          <TR class="tdbg" > 
            <TD width="120" align="right"><strong>�û�״̬��</strong></TD>
            <TD><input type="radio" name="LockUser" value="False" <%if rsUser("LockUser")=False then response.write "checked"%>>
              ����&nbsp;&nbsp; <input type="radio" name="LockUser" value="True" <%if rsUser("LockUser")=True then response.write "checked"%>>
              ����</TD>
          </TR>
          <TR align="center" class="tdbg" > 
            <TD height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify"> 
              <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"></TD>
          </TR>
        </TABLE>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
	end if
	rsUser.close
	set rsUser=nothing
end if
call CloseConn()
%>
