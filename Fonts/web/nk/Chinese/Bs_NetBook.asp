<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
<!--#include file="Include/Bs_StrLeft.asp"-->
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
response.redirect"Bs_Server.asp"
end if%>
<%
dim Id,rst,sql
Id=Session("UserName")
set rst = Server.CreateObject("ADODB.recordset")
sql="select * from Bs_User where UserName='"&Id&"'"
rst.open sql,conn,1,1
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

<!-- 留言中心 -->
		<TR>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0  align="center" class='MC_Pt_Title'>
					<TR>
						<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
						<TD class='MC_Pt_TD3'><SPAN class=Glow>
留言中心
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
					<TR><TD height=10><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
					<TR> 
						<TD>
<!--  -->			
				<table width="90%" border="0" align="center" cellpadding="0" cellspacing="5">
					<tr> 
						<td><div align="center"><img src="../img/newsico.gif" width="12" height="13"
						>&nbsp;&nbsp;<b><%=Session("UserName")%></b> 的留言专区</div></td>
					</tr>
					<tr> 
						<td style="line-height: 150%">
						&nbsp;&nbsp;&nbsp;&nbsp;<%=Session("UserName")%>，这里是您的专用私密留言簿，您可在此给我们留言，我们会尽力在24小时内在此答复您，请不要忘记回来察看回复哦！<br>
						&nbsp;&nbsp;&nbsp;&nbsp;如您已经在我站购物，汇款后请到这里进行汇款登记，填写您的订单号码、汇款时间、汇往我们哪个账号，以便我们核查您的款项的到帐情况并及时为您发货。</td>
					</tr>
				</table>
				<form method="post" action="Bs_NetBookSave.asp">
				<input type=hidden readonly name="Name" value="<%=Session("UserName")%>">
					<table width="90%" border="0" align="center" cellpadding="4" cellspacing="1">
						<tr> 
							<td width="22%" align="right">用户名称：</td>
							<td width="78%"><font> 
							<input type="text" readonly name="Name2" maxlength="36" value="<%if Session("UserName")="" then response.write"未注册用户" end if%><%=Session("UserName")%>" style="background-color: #FFFFFF; border-style: solid; border-color: #FFFFFF" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">公司名称：</td>
							<td width="78%"><font> 
							<input type="text" name="Comane" size="30" maxlength="36" value="<%=rst("Comane")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">联系人：</td>
							<td width="78%"><font> 
							<input type="text" name="Somane" size="12" maxlength="30" value="<%=rst("Somane")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">联系电话：</td>
							<td width="78%"><font> 
							<input type="text" name="Phone" size="20" maxlength="36" value="<%=rst("Phone")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">联系传真：</td>
							<td width="78%"><font> 
							<input type="text" name="Fox" size="20" maxlength="36" value="<%=rst("Fox")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td align="right">E-mail：</td>
							<td width="78%"><font> 
							<input type="text" name="Email" size="20" maxlength="36" value="<%=rst("Email")%>" style="font-size: 14px" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td width="22%" align="right">留言标题：</td>
							<td width="78%"><font>
							<input type="text" name="Title" size="22" maxlength="36" style="font-size: 14px; width: 320; height: 19" class="smallInput">
							</font></td>
						</tr>
						<tr> 
							<td width="22%" align="right">留言内容：</td>
							<td width="78%"><font>
							<textarea rows="6" name="Content" cols="50" style="font-size: 14px" class="smallInput"></textarea> 
							</font></td>
						</tr>
						<tr> 
							<td width="22%" align="right"></td>
							<td width="78%"><input type="submit" value="提交你的留言" name="cmdOk"
							> &nbsp; <input type="reset" value="复位" name="cmdReset"> </td>
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
If rs.eof and rs.bof then
	'call showpages
	response.write "<p align='center'><font color='#ff0000'>您还没任何留言资料</font></p>"
	response.end
End if

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
						<td width="14%" height="18" bgcolor="#eeeeee"><img src="../img/arrow.gif" width="16" height="13"><strong>留言列表</strong></td>
<td width="86%" bgcolor="#eeeeee"> <%
Response.write "<strong><font color='#000000'>-> 全部-</font>"
Response.write "共</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>条留言</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>第</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='Bs_NetBook.asp?&page="+cstr(1)+"'><font color='#000000'>首页</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='Bs_NetBook.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'>上一页</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>上一页 </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='Bs_NetBook.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'>下一页</font></a><font color='#ffffff'> </font>"
Response.write "<a href='Bs_NetBook.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'>尾页</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>下一页</font></strong>&nbsp;&nbsp;"
End if
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
						><p align="center"><font color="#000000"><b>留言主题</b></font></p></td>
						<td width="22%" height="21" align="center" bgcolor="#CCCCCC"><font color="#000000"><b>时间</b></font></td>
						<td width="22%" height="21" align="center" bgcolor="#CCCCCC"><font color="#000000"><b>回复</b></font></td>
					</tr>
<%
if not rs.eof then
i=0
do while not rs.eof
%>
					<tr bgcolor="#99CCFF"> 
						<td height="22" bgcolor="#FFFFFF">
<%if rs("name")="管理员"then%>
『网站公告』 
<%end if%> 
						<a href="javascript:winopen('Bs_NetBookRe.asp?id=<%=rs("id")%>&amp;name=<%=rs("name")%>')"><%=rs("title")%></a>
						</td>
						<td height="1" bgcolor="#FFFFFF" align="center">&nbsp; <%=rs("time")%></td>
						<td height="1" bgcolor="#FFFFFF" align="center">&nbsp; <%If rs("rebook")<>"" Then%> 
						<a href="javascript:winopen('Bs_NetBookRe.asp?id=<%=rs("id")%>&amp;name=<%=rs("name")%>')">回复</a> 
<%else%>
无 
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