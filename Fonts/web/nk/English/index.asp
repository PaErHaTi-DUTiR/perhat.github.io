<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/En_System.asp"-->
<!--#include file="../Inc/ubbcode.asp"-->
<!--#include file="Include/En_SysProduct.asp"-->
<!--#include file="Include/En_SysCoSaJo.asp"-->
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
<%  
dim rsCou,sqltext,mylogin
set rsCou = Server.CreateObject("ADODB.Recordset")

if Session("mylogin")<>"1" then
	sqltext="update Bs_SysData set querycount=querycount+1 where code='5062942'"
	conn.execute(sqltext)
	
	rsCou.Open "webcount",conn,3,3
	rsCou.AddNew
	rsCou("ipaddress")=Request.ServerVariables("Remote_HOST")
	rsCou("logindate")=now()
	rsCou.Update
	rsCou.Close
			
	mylogin=Session("mylogin")
	Session("mylogin")="1"
else
	mylogin=Session("mylogin")	
end if
%>

<%
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function
%>
<DIV align="center">
<!--#include file="Include/en_Head.asp" -->
<%if strSkins <= 200 and strSkins >= 1 then%>
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR> 
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Index() %>
			</td>
		</tr>
	</table> 
</TD>

<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 产品展示 -->
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 class='MC_Pt_Title'>
			<TR>
				<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD3'><SPAN class=Glow>
News Product
&nbsp;</SPAN></TD>
				<TD class='HomeMC_Pt_TD5'><a href="En_Product.asp" title="More product exhibition ..."><IMG  src='../img/1x1_pix.gif' width=52 height=23 border=0></a></TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD> 
		<table width="100%" border="0" cellspacing="1" cellpadding="0" class=main_tdbg>
			<tr>
<TD><% call ProductZY() %></TD>
			</tr>
		</table>
	</TD>
</TR>
<!-- 产品展示END -->
<TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<!-- 公司介绍 -->
<TR>
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
			<TR>
<TD width=29><IMG src="../Skin/Img/arrow3.gif" width=29 height=11 border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>Company introduction</SPAN></TD>
<TD width=35 align=right style="padding-top:4px"><a href="En_CoProfile.asp?Action=Profile" title="More company introductions ..."><IMG src="../Skin/Img/more.gif" width=27 height=9 border=0></a>&nbsp;</TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<TR> 
	<TD>
		<table width="100%" border="0" cellspacing="5" cellpadding="0" class=main_tdbg>
			<tr> 
				<td style="line-height: 150%">　　<%=CSJ_Profile%></td>
			</tr>
		</table>
	</TD>
</TR>
<!-- 公司介绍 END -->
<TR><TD height=6><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>

</TABLE> 
</TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot.asp" -->
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TR> 
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Index() %>
			</td>
		</tr>
	</table> 
</TD>
<TD vAlign=top  align="center"> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 产品展示 -->
<TR>
	<TD>
		<TABLE align="center" cellPadding=0 cellSpacing=0 class='MC_Pt_Title'>
			<TR>
				<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD3'><SPAN class=Glow>
News Product
&nbsp;</SPAN></TD>
				<TD class='HomeMC_Pt_TD5'><a href="En_Product.asp" title="More product exhibition ..."><IMG  src='../img/1x1_pix.gif' width=52 height=23 border=0></a></TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD align="center"> 
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class=main_tdbg>
			<tr>
<TD align="center"><% call ProductZY() %></TD>
			</tr>
		</table>
	</TD>
</TR>
<!-- 产品展示END -->
<TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<!-- 公司介绍 -->
<TR>
	<TD>
		<TABLE width="570" align="center" cellPadding=0 cellSpacing=0 class='HomeMC_Pt2_Title'>
			<TR>
<TD width=29><IMG src="../Skin/Img/arrow3.gif" width=29 height=11 border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>Company introduction</SPAN></TD>
<TD width=35 align=right style="padding-top:4px"><a href="En_CoProfile.asp?Action=Profile" title="More company introductions ..."><IMG src="../Skin/Img/more.gif" width=27 height=9 border=0></a>&nbsp;</TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<TR> 
	<TD>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="5" class=main_tdbg>
			<tr> 
				<td style="line-height: 150%">　　<%=CSJ_Profile%></td>
			</tr>
		</table>
	</TD>
</TR>
<!-- 公司介绍 END -->
<TR><TD height=6><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
</TD>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD><TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Index() %></TD>
</TR>
</TABLE>
<!--#include file="Include/en_Foot4.asp" -->
<%end if%>
</DIV>
<%
'--------产品展示--------
Sub ProductZY()
dim rs_ProductZY,sql_ProductZY,row_count,strLb,strli
set rs_ProductZY=server.createobject("adodb.recordset")
sql_ProductZY="select top 12 * from Bs_Product where Passed=True And EnElite=True order by UpdateTime desc"
rs_ProductZY.open sql_ProductZY,conn,1,1
	Response.Write "<table width=100% border=0 cellspacing=2 cellpadding=0 align=center><TR><TR><TD></TD></TR>"

	row_count=1
	 Response.Write "<TR align=""center"">"
	 Do While Not rs_ProductZY.EOF
		strli=""
		strli=strli & " ● &nbsp;&nbsp;Name："& rs_ProductZY("EnTitle") &" "&chr(13)
		strli=strli & " ● &nbsp;&nbsp;Spec："& rs_ProductZY("EnSpec") &" "&chr(13)
		strli=strli & " ● &nbsp;&nbsp;Size："& rs_ProductZY("Size") &" "
			strLb=""
			strLb=strLb & "<TD>"
			strLb=strLb & "<table width=130 border=0 cellpadding=0 cellspacing=0 align=center>"
			strLb=strLb & "<TR><TD height=3></TD></TR>"

			strLb=strLb & "<tr>"
			strLb=strLb & "<td width='98%'><div align=center>"
			strLb=strLb & "<a href='En_ProductShow.asp?ArticleID="& rs_ProductZY("articleid") &"'title='" & strli &"'>"
			strLb=strLb & "<img border=0 src='"& rs_ProductZY("DefaultPicUrl") &"' width=130 height=120 class=bk></a>"
			strLb=strLb & "</div></td>"
			strLb=strLb & "</tr>"

			strLb=strLb & "<TR><TD height=3></TD></TR>"
	'		strLb=strLb & "<tr><td width=98% align=center>产品名称</td></tr>"
			strLb=strLb & "<tr>"
			strLb=strLb & "<td height=20 bgcolor=#F0F0F0><div align=center>"
			strLb=strLb & "<a href='En_ProductShow.asp?ArticleID="& rs_ProductZY("articleid") &"'title='"& strli &"'>"& gotTopic(rs_ProductZY("EnTitle"),18) &"</a>"
			strLb=strLb & "</div></td>"
			strLb=strLb & "</tr>"

			strLb=strLb & "<tr>"
			strLb=strLb & "<td height=3 bgcolor=#DCDDDE><IMG height=1 src='../img/1x1_pix.gif' width=1></td>"
			strLb=strLb & "</tr>"
			strLb=strLb & "</table></TD>"
			Response.Write strLb

		if row_count mod 4 <>0 then
		end if
		if row_count mod 4 =0 then
		Response.Write "</TR><TR>"
		end if

		rs_ProductZY.MoveNext
	row_count=row_count+1
	Loop

rs_ProductZY.close
 Response.Write "</TR><TR><TD></TD></TR></Table>"
set rs_rs_ProductZY=nothing
END Sub
%>
