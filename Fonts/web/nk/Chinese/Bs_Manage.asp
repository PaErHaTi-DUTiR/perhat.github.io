<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
<!--#include file="Include/Bs_StrLeft.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<%
dim Action,UserName
dim FoundErr,ErrMsg
dim rsUser,sqlUser
dim Sex,Email,Homepage,Comane,Add,Somane,Zip,Phone,Fox
Action=trim(request("Action"))
'UserName=trim(request("UserName"))

if UserName="" then
	UserName=session("UserName")
end if
if  UserName="" then
	if Action="" then
		response.redirect "Bs_Server.asp"
	else
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	end if
end if
if FoundErr=true then
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [Bs_User] where UserName='" & UserName & "'"
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		call writeErrMsg()
	else
		if Action="Modify" then
			Sex=trim(Request("Sex"))
			Email=trim(request("Email"))
			Homepage=trim(request("Homepage"))
			Comane=trim(request("Comane"))
			Add=trim(request("Add"))
			Somane=trim(request("Somane"))
			Zip=trim(request("Zip"))
			Phone=trim(request("Phone"))
			Fox=trim(request("Fox"))
			if Sex="" then
				FoundErr=true
				ErrMsg=ErrMsg & "<br><li>性别不能为空</li>"
			else
				sex=cint(sex)
				if Sex<>0 and Sex<>1 then
					Sex=1
				end if
			end if
			if Email="" then
				FoundErr=true
				ErrMsg=ErrMsg & "<br><li>Email不能为空</li>"
			else
				if IsValidEmail(Email)=false then
					ErrMsg=ErrMsg & "<br><li>您的Email有错误</li>"
				  FoundErr=true
				end if
			end if
			if Comane="" then
				FoundErr=true
				ErrMsg=ErrMsg & "<br><li>公司名称不能为空</li>"
			end if
			if Add="" then
				FoundErr=true
				ErrMsg=ErrMsg & "<br><li>收货地址不能为空</li>"
			end if
			if Somane="" then
				FoundErr=true
				ErrMsg=ErrMsg & "<br><li>收货人不能为空</li>"
			end if
			if Phone="" then
				FoundErr=true
				ErrMsg=ErrMsg & "<br><li>联系电话不能为空</li>"
			end if
			if FoundErr<>true then
				rsUser("Sex")=Sex
				rsUser("Email")=Email
				rsUser("HomePage")=HomePage
				rsUser("Comane")=Comane
				rsUser("Add")=Add
				rsUser("Somane")=Somane
				rsUser("Zip")=Zip
				rsUser("Phone")=Phone
				rsUser("Fox")=Fox
				rsUser.update
				response.write"<SCRIPT language=JavaScript>alert('会员资料修改成功！');"
                response.write"javascript:history.go(-1)</SCRIPT>"
				'call WriteSuccessMsg("会员资料修改成功！")
			else
				call WriteErrMsg()
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
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<%end if%>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 修改会员资料 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
修改注册用户信息
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
						<TD><FORM name="Form1" action="Bs_Manage.asp" method="post">
							<table width="90%" border="0" align="center" cellpadding="0" cellspacing="5">
								<TR><TD height=15><IMG height=1 src="img/1x1_pix.gif" width=10></TD></TR>
								<TR class="tdbg" > 
									<TD width="180" align="right"><b>用 户 名：</b></TD>
									<TD> <%=rsUser("UserName")%><INPUT name="UserName" type="hidden" value="<%=rsUser("UserName")%>" size=30 maxLength=30></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="180" align="right"><strong>性别：</strong></TD>
									<TD> <INPUT type=radio value="1" name=sex <%if rsUser("Sex")=1 then response.write "CHECKED"%>>
									男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if rsUser("Sex")=0 then response.write "CHECKED"%>>
									女</TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="180" align="right"><strong>Email地址：</strong></TD>
									<TD> <INPUT name=Email value="<%=rsUser("Email")%>" size=30   maxLength=50></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="180" align="right"><strong>主页：</strong></TD>
									<TD> <INPUT   maxLength=100 size=30 name=homepage value="<%=rsUser("HomePage")%>"></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="180" align="right"><strong>公司名称：</strong></TD>
									<TD> <INPUT name=Comane value="<%=rsUser("Comane")%>" size=30 maxLength=20></TD>
								</TR>
								<TR class="tdbg" > 
									<TD width="180" align="right"><strong>收货地址：</strong></TD>
									<TD> <INPUT name=Add value="<%=rsUser("Add")%>" size=30 maxLength=50></TD>
								</TR>
								<TR class="tdbg" > 
									<TD align="right"><strong>收货人：</strong></TD>
									<TD><INPUT name=Somane value="<%=rsUser("Somane")%>" size=30 maxLength=50></TD>
								</TR>
								<TR class="tdbg" > 
									<TD align="right"><strong>邮政编码：</strong></TD>
									<TD><INPUT name=Zip value="<%=rsUser("Zip")%>" size=30 maxLength=50></TD>
								</TR>
								<TR class="tdbg" > 
									<TD align="right"><strong>联系电话：</strong><br></TD>
									<TD><INPUT name=Phone value="<%=rsUser("Phone")%>" size=30 maxLength=50></TD>
								</TR>
								<TR class="tdbg" > 
									<TD align="right"><strong>传 真：</strong></TD>
									<TD><INPUT name=Fox value="<%=rsUser("Fox")%>" size=30 maxLength=50></TD>
								</TR>
								<TR align="center" class="tdbg" > 
									<TD height="50" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify"
									><input name=Submit type=submit id="Submit" value="保存修改结果"></TD>
								</TR>
							</TABLE>
						</form></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
<%
end if
end if
rsUser.close
set rsUser=nothing
end if
%>
<!-- 修改会员资料 End -->
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
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>