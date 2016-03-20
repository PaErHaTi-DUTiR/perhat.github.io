<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/En_System.asp"-->
<!--#include file="Include/En_StrLeft.asp"-->
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<%If Session("UserName")="" Then
response.redirect"En_Server.asp"
end if%>
<%
dim Id,rst,sql
Id=Session("UserName")
set rst = Server.CreateObject("ADODB.recordset")
sql="select * from Bs_User where UserName='"&Id&"'"
rst.open sql,conn,1,1
%>
<!--#include file="Include/En_Head.asp" -->
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
<%end if%><TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 留言中心 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
Message centre
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
				<TABLE cellSpacing=0 cellPadding=0 class='MC_ServeNr_Title'>
					<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
					<TR> 
						<TD>
<!--  -->			
				<table width="90%" border="0" align="center" cellpadding="0" cellspacing="5">
					<tr> 
						<td><div align="center"><img src="../img/newsico.gif" width="12" height="13"
						>&nbsp;&nbsp;<b><%=Session("UserName")%></b> Message prefecture</div></td>
					</tr>
					<tr> 
						<td>　　<%=Session("UserName")%>，This is your special-purpose private dense visitor book, you can leave<BR>a message for us here , we will is it in reply you this , please don is it is it watch and<BR>reply to come back to forget within 24 hours to do the best！If you is it do shopping to<BR>stand in me already, please reach here register the remittance after remitting money,<BR>fill in your order number , remittance time , which account number is remitted to us,<BR>so that we verify reach the situation of accounts and deliver for you in timing of your fund . </td>
					</tr>
				</table>
				<form method="post" action="En_NetBooKSave.asp">
				<input type=hidden readonly name="Name" value="<%=Session("UserName")%>">
					<table width="90%" border="0" align="center" cellpadding="4" cellspacing="1">
						<tr> 
							<td width="22%" align="right"></td>
							<td width="78%"></td>
						</tr>
						<tr> 
							<td width="22%" align="right">UserName：</td>
							<td width="78%"><font> 
							<input type="text" readonly name="Name2" maxlength="36" value="<%if Session("UserName")="" then response.write"未注册用户" end if%><%=Session("UserName")%>" style="background-color: #FFFFFF; border-style: solid; border-color: #FFFFFF" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">CompanyName：</td>
							<td width="78%"><font> 
							<input type="text" name="Comane" size="30" maxlength="36" value="<%=rst("Comane")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">ContactPerson：</td>
							<td width="78%"><font> 
							<input type="text" name="Somane" size="12" maxlength="30" value="<%=rst("Somane")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">Telephone：</td>
							<td width="78%"><font> 
							<input type="text" name="Phone" size="24" maxlength="36" value="<%=rst("Phone")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">Fax：</td>
							<td width="78%"><font> 
							<input type="text" name="Fox" size="18" maxlength="36" value="<%=rst("Fox")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">E-mail：</td>
							<td width="78%"><font> 
							<input type="text" name="Email" size="18" maxlength="36" value="<%=rst("Email")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td width="22%" align="right">LeaveWordTheme：</td>
							<td width="78%"><font>
							<input type="text" name="Title" size="22" maxlength="36" style="font-size: 14px; width: 320; height: 19" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td width="22%" align="right">LeaveWordContent：</td>
							<td width="78%"><font>
							<textarea rows="6" name="Content" cols="50" style="font-size: 14px" class="smallInput"></textarea> 
							</font></td>
						</tr>
						<tr> 
							<td width="22%" align="right"></td>
							<td width="78%"><input type="submit" value="ReferWord" name="cmdOk"
							> &nbsp; <input type="reset" value="Rewrite" name="cmdReset"> </td>
						</tr>
					</table>
				</form>
<%
dim name,rs
name=Session("UserName")
set rs=server.createobject("adodb.recordset")
sql="select * from Bs_Book where name='"&name&"'or name='管理员'order by id desc"
rs.open sql,conn,1,1

dim MaxPerPageB
MaxPerPageB=10

'假如没有数据时
'If rs.eof and rs.bof then
'call showpages
'response.write "<p align='center'><font color='#ff0000'>您还没任何留言资料</font></p>"
'response.end
'End if

'取得页数,并判断留言输入的是否数字类型的数据，如不是将以第一页显示
dim text,checkpage,CurrentPage,i
text="0123456789"
Rs.PageSize=MaxPerPageB
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if

call showpages
call list

If Rs.recordcount > MaxPerPageB then
call showpages
end if

'显示帖子的子程序
Sub list()

%> 
				<table width="90%" height="19" border="0" align="center" cellpadding="0" cellspacing="0">
					<tr> 
						<td width="22%" height="18" bgcolor="#eeeeee"><img src="../img/arrow.gif" width="16" height="13"><strong>Leave word list</strong></td>
<td width="86%" bgcolor="#eeeeee"> <%
Response.write "<strong><font color='#000000'>-> All- </font>"
Response.write " amount </font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'> bar info </font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'> </font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='En_NetBook.asp?&page="+cstr(1)+"'><font color='#000000'> Home </font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='En_NetBook.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'> Previous </font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'> Previous </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='En_NetBook.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'> Next </font></a><font color='#ffffff'> </font>"
Response.write "<a href='En_NetBook.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'> End </font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'> Next </font></strong>&nbsp;&nbsp;"
End if
'response.write "</td><td align='right'>"
'response.write "<font color='#000000'>转到：</font><input type='text'name='page'size=4 maxlength=4 class=smallInput value="&Currentpage&">&nbsp;"
'response.write "<input class=buttonface type='submit' value='Go' name='cndok'>&nbsp;&nbsp;"
%> </td>
					</tr>
					<tr> 
						<td height="1" colspan="2" bgcolor="#999999"></td>
					</tr>
				</table>
<%
If rs.eof and rs.bof then
call showpages
response.write ""
response.end
End if
%> <br>
				<table width="90%" height="18" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#666666">
					<tr bgcolor="#99CCFF"> 
						<td width="56%" height="25" align="center" bgcolor="#CCCCCC"
						><p align="center"><font color="#000000"><b>Leave word motif</b></font></p></td>
						<td width="22%" height="21" align="center" bgcolor="#CCCCCC"><font color="#000000"><b>Time</b></font></td>
						<td width="22%" height="21" align="center" bgcolor="#CCCCCC"><font color="#000000"><b>Revert</b></font></td>
					</tr>
<%
if not rs.eof then
i=0
do while not rs.eof
%>
					<tr bgcolor="#99CCFF"> 
						<td height="22" bgcolor="#FFFFFF">
<%if rs("name")="管理员"then%>
『Website bulletin』 
<%end if%> 
						<a href="javascript:winopen('En_NetBookRe.asp?id=<%=rs("id")%>&amp;name=<%=rs("name")%>')"><%=rs("title")%></a>
						</td>
						<td height="1" bgcolor="#FFFFFF" align="center">&nbsp; <%=rs("time")%></td>
						<td height="1" bgcolor="#FFFFFF" align="center">&nbsp; <%If rs("rebook")<>"" Then%> 
						<a href="javascript:winopen('En_NetBookRe.asp?id=<%=rs("id")%>&amp;name=<%=rs("name")%>')">Revert</a> 
<%else%>
No 
<%End If%> &nbsp;
						</td>
					</tr>
<%
i=i+1
if i >= MaxPerPageB then exit do
rs.movenext
loop
end if
%>
<%end sub%>
				</table>
<!--  -->
						</TD>
					</TR>
		<TR> 
			<TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
		</TR>
				</table>
			</TD>
		</TR>
<%
rst.close
set rst=nothing
%>
<!-- 留言中心 End -->
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
<!--#include file="Include/En_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-

mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Download() %></TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot4.asp" -->
<%end if%>