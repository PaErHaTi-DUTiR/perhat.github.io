<%@ LANGUAGE=VBScript CodePage=936%>
<%option explicit%>
<%Response.Buffer=True%>
<!--#include file="Include/Bs_System.asp"-->
<!--#include file="../Inc/ubbcode.asp"-->
<!--#include file="Include/Bs_SysProduct.asp"-->
<!--#include file="Include/Bs_SysCoSaJo.asp"-->
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
end function%>
<DIV align="center">
<!--#include file="Include/Bs_Head.asp" -->

<%if strSkins <= 200 and strSkins >= 1 then%>

<TABLE cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TD vAlign=top class='M_Left_Td'> 
	<table cellpadding="0" cellspacing="0" align="center" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Index() %>
<% Call Style_Vote() '站内调查 %>
			</td>
		</tr>
	</table> 
</TD>

<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0 align='center' class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>
<!-- 公司介绍 -->
<TR>
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
			<TR>
<TD width=29><IMG src="../Skin/Img/arrow3.gif" width=29 height=11 border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>公司介绍</SPAN></TD>
<TD width=35 align=right style="padding-top:4px"><a href="Bs_CoProfile.asp?Action=Profile" title="更多公司介绍..."><IMG src="../Skin/Img/more.gif" width=27 height=9 border=0></a>&nbsp;</TD>
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

<!-- 产品展示 -->
</TR>
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 class='MC_Pt_Title'>
			<TR>
				<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD3'><SPAN class=Glow>最新产品&nbsp;</SPAN></TD>
				<TD class='HomeMC_Pt_TD5'><a href="Bs_Product.asp" title="更多产品展示..."><IMG  src='../img/1x1_pix.gif' width=52 height=23 border=0></a></TD>
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
<!-- 新闻资讯 -->
<TR>
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 width="100%" border="0">
			<TR>
				<TD align="left" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/Img/arrow3.gif" width=29 height=11 border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>公司新闻</SPAN></TD>
<TD width=35 align=right style="padding-top:4px"><a href="Bs_News.asp?Action=Co" title="更多新闻..."><IMG src="../Skin/Img/more.gif" width=27 height=9 border=0></a>&nbsp;</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD colspan="4">
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top"><% call NewsZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
				<TD width=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD align="right" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/Img/arrow3.gif" width=29 height=11 border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>业内资讯</SPAN></TD>
<TD width=35 align=right style="padding-top:4px"><a href="Bs_News.asp?Action=Ye" title="更多资讯..."><IMG src="../Skin/Img/more.gif" width=27 height=9 border=0></a>&nbsp;</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD colspan="4">
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top"><% call NewsYeZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<!-- 新闻资讯END -->
<TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<!-- 产品动态&下载中心 -->
<TR>
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 width="100%" border="0">
			<TR>
				<TD align="left" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/Img/arrow3.gif" width=29 height=11 border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>产品动态</SPAN></TD>
<TD width=35 align=right style="padding-top:4px"><a href="Bs_News.asp?Action=Pr" title="更多产品动态..."><IMG src="../Skin/Img/more.gif" width=27 height=9 border=0></a>&nbsp;</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD>
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top"><% call NewsProductZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
				<TD width=5><IMG width=2 height=2 src="../img/1x1_pix.gif"></TD>
				<TD align="right" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/Img/arrow3.gif" width=29 height=11 border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>下载中心</SPAN></TD>
<TD width=35 align=right style="padding-top:4px"><a href="Bs_Download.asp" title="更多下载..."><IMG src="../Skin/Img/more.gif" width=27 height=9 border=0></a>&nbsp;</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD>
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top" style="line-height: 150%"><% call DownZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<!-- 产品动态&下载中心END -->
<TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>

</TABLE> 
</TD>
</TR>

</TABLE>
<!--#include file="Include/Bs_Foot.asp" -->

<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<TABLE  align="center" border="0" cellPadding=0 cellSpacing=0 class='Middle_Title'>
<TD vAlign=top class='M_Left_Td'> 
	<table border="0" align="center" cellpadding="0" cellspacing="0" class='M_Left_Title'>
		<tr>
			<td valign="top">
<% Call Left_Index() %>
<% Call Style_Vote() '站内调查 %>
			</td>
		</tr>
	</table> 
</TD>
<TD vAlign=top> 
	<TABLE cellPadding=0 cellSpacing=0  align="center" class='M_Center_Title'>
<% Call Style_Adcolumn() '广告栏 %>

<!-- 新闻资讯 -->
<TR>
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 width="570" border="0" align="center">
			<TR>
				<TD align="left" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/201/arrow3.gif"  border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>公司新闻</SPAN></TD>
<TD width=44 align=right><a href="Bs_News.asp?Action=Co" title="更多新闻..."><IMG src="../Skin/201/more.gif"  border=0></a></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD colspan="4">
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top"><% call NewsZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
<td width="7" align="center" valign="top" background="../Skin/201/7×1.gif"><img src="../Skin/201/7×1.gif" width="7" height="1"></td>				
							<TD align="right" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/201/arrow3.gif"  border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>业内资讯</SPAN></TD>
<TD width=44 align=right ><a href="Bs_News.asp?Action=Ye" title="更多资讯..."><IMG src="../Skin/201/more.gif"  border=0></a></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD colspan="4">
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top"><% call NewsYeZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<!-- 新闻资讯END -->
<TR><TD >	    <table width="100%" border="0" cellspacing="1" cellpadding="0" class=main_tdbg>
			<tr>
<TD align="center"><img border="0" src="../skin/201/mid-line.gif" width="562" height="6"></TD>
			</tr>
		</table></TD></TR>

<!-- 产品展示 -->
<TR>
	<TD>
		<TABLE cellSpacing=0 cellPadding=0 align="center" class='MC_Pt_Title'>
			<TR>
				<TD class='MC_Pt_TD1'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD2'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD>
				<TD class='MC_Pt_TD3'><SPAN class=Glow>最新产品&nbsp;</SPAN></TD>
				<TD class='HomeMC_Pt_TD5'><a href="Bs_Product.asp" title="更多产品展示..."><IMG  src='../img/1x1_pix.gif' width=52 height=23 border=0></a></TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD align="center"> 
		<table width="100%" border="0" cellspacing="1" cellpadding="0" class=main_tdbg>
			<tr>
<TD align="center"><% call ProductZY() %></TD>
			</tr>
		</table>
	</TD>
</TR>
<!-- 产品展示END -->
<TR><TD >	
	    <table width="570" border="0" cellspacing="1" cellpadding="0" class=main_tdbg>
			<tr>
<TD align="center" width="570"><img border="0" src="../skin/201/mid-line.gif" width="562" height="6"></TD>
			</tr>
		</table></TD></TR>

<!-- 产品动态&下载中心 -->
<TR>
	<TD>
		<TABLE width="570" border="0" align="center" cellPadding=0 cellSpacing=0>
			<TR>
				<TD align="left" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/201/arrow3.gif"  border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>产品动态</SPAN></TD>
<TD width=44 align=right ><a href="Bs_News.asp?Action=Pr" title="更多产品动态..."><IMG src="../Skin/201/more.gif"  border=0></a></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD>
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top"><% call NewsProductZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
<td width="7" align="center" valign="top" background="../Skin/201/7×1.gif"><img src="../Skin/201/7×1.gif" width="7" height="1"></td>				
				<TD align="right" class=HomeMC_2TD>
					<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
						<TR>
							<TD>
								<TABLE cellSpacing=0 cellPadding=0 class='HomeMC_Pt2_Title'>
									<TR>
<TD width=29><IMG src="../Skin/201/arrow3.gif"  border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>下载中心</SPAN></TD>
<TD width=44 align=right ><a href="Bs_Download.asp" title="更多下载..."><IMG src="../Skin/201/more.gif"  border=0></a></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
						<TR> 
							<TD>
								<TABLE width="100%" height="200" cellSpacing=3 cellPadding=0 border=0 class=main_tdbg>
									<TR>
<TD width="100%" valign="top" style="line-height: 150%"><% call DownZY() %></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<!-- 产品动态&下载中心END -->
<TR><TD>	    <table width="100%" border="0" cellspacing="1" cellpadding="0" class=main_tdbg>
			<tr>
<TD align="center"><img border="0" src="../skin/201/mid-line.gif" width="562" height="6"></TD>
			</tr>
		</table></TD></TR>
<!-- 公司介绍 -->
<!-- 
<TR>
	<TD>
		<TABLE width="570" align="center" cellPadding=0 cellSpacing=0 class='HomeMC_Pt2_Title'>
			<TR>
<TD width=29><IMG src="../Skin/201/arrow3.gif"  border=0></TD>
<TD style="padding-top:2px" align=left><SPAN class=Glow>公司介绍</SPAN></TD>
<TD width=44 align=right ><a href="Bs_CoProfile.asp?Action=Profile" title="更多公司介绍..."><IMG src="../Skin/201/more.gif" border=0></a></TD>
			</TR>
		</TABLE>
	</TD>
</TR>
<TR><TD height=1><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<TR> 
	<TD>
		<table width="570" border="0" align="center" cellpadding="0" cellspacing="5" class=main_tdbg>
			<tr> 
				<td style="line-height: 150%"><%=CSJ_Profile%></td>
			</tr>
		</table>
	</TD>
</TR>
-->
<!-- 公司介绍 END -->
<TR><TD height=6><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
</TD>
<TD class='M_Jgx_TD'><IMG src='../img/1x1_pix.gif' width=1 height=2></TD>
<TD vAlign=top background="../Skin/201/line-mid-glay.gif" width="6"><img border="0" src="../Skin/201/line-shadow-glay-mid.gif" width="6" height="146"></TD>
<TD vAlign=top class='Mr_TitlePt'><% Call Right_Index() %></TD>
</TR>
</TABLE>
<!--#include file="Include/Bs_Foot4.asp" -->
<%end if%>
</DIV>

<%
'-------- --------
Sub ProductZY()
dim rs_ProductZY,sql_ProductZY,row_count,strLb,strli
set rs_ProductZY=server.createobject("adodb.recordset")
sql_ProductZY="select top 12 * from Bs_Product where Passed=True And Elite=True order by UpdateTime desc"
rs_ProductZY.open sql_ProductZY,conn,1,1
	Response.Write "<table width=100% border=0 cellspacing=0 cellpadding=0 align=center><TR><TD></TD></TR>"

	row_count=1
	 Response.Write "<TR align=""center"">"
	 Do While Not rs_ProductZY.EOF
		strli=""
		strli=strli & "├┈&nbsp;产品名称："& rs_ProductZY("Title") &" "&chr(13)
		strli=strli & "├┈&nbsp;产品规格："& rs_ProductZY("Spec") &" "&chr(13)
		strli=strli & "├┈&nbsp;产品尺寸："& rs_ProductZY("Size") &" "
			strLb=""
			strLb=strLb & "<TD>"
			strLb=strLb & "<table width=130 align=center border=0 cellpadding=0 cellspacing=0>"
			strLb=strLb & "<TR><TD height=3></TD></TR>"
			strLb=strLb & "<tr>"
			strLb=strLb & "<td width='98%'><div align=center>"
			strLb=strLb & "<a href='Bs_ProductShow.asp?ArticleID="& rs_ProductZY("articleid") &"'title='" & strli &"'>"
			strLb=strLb & "<img border=0 src='"& rs_ProductZY("DefaultPicUrl") &"' width=130 height=120 class=bk></a>"
			strLb=strLb & "</div></td>"
			strLb=strLb & "</tr>"
			strLb=strLb & "<TR><TD height=3></TD></TR>"
	'		strLb=strLb & "<tr><td width=98% align=center>产品名称</td></tr>"
			strLb=strLb & "<tr>"
			strLb=strLb & "<td height=20 bgcolor=#f0f0f0><div align=center>"
			strLb=strLb & "<a href='Bs_ProductShow.asp?ArticleID="& rs_ProductZY("articleid") &"'title='"& strli &"'>"& gotTopic(rs_ProductZY("Title"),16) &"</a>"
			strLb=strLb & "</div></td>"
			strLb=strLb & "</tr>"
			strLb=strLb & "<tr>"
			strLb=strLb & "<td height=3 bgcolor=#DCDDDE><IMG height=1 src='../img/1x1_pix.gif' width=1></td>"
			strLb=strLb & "</tr>"
			strLb=strLb & "</table><br></TD>"
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

'-------- --------
Sub NewsZY()
dim rs_NewsZY,sql_NewsZY,i,Title,StrTemp
set rs_NewsZY=server.createobject("adodb.recordset")
sql_NewsZY="select top 10 * from Bs_NewsCo order by id desc"
rs_NewsZY.open sql_NewsZY,conn,1,1
i=0
	Do While Not rs_NewsZY.EOF
		Title=""
		if rs_NewsZY("Title")<>"" then
			StrTemp=replace(rs_NewsZY("Title"),"<b>","")
			StrTemp=replace(StrTemp,"&nbsp;","")
			Title=Title&StrTemp
		else
			Title=Title&"此栏目没有任何内容"
		end if

		Response.Write "<IMG src=""../Skin/Img/article_common.gif"" width=9 height=15 border=0> "
		Response.Write "<a href='Bs_NewsInfo.asp?Action=Co&ID=" & rs_NewsZY("ID") & "'title='"& rs_NewsZY("Title") &"'>"& gotTopic(Title,36) & "</a><BR>"
	rs_NewsZY.MoveNext
	i=i+1
	if i=10 then exit do
	Loop

rs_NewsZY.close
set rs_NewsZY=nothing
END Sub

'-------- --------
Sub NewsYeZY()
dim rs_NewsYeZY,sql_NewsYeZY,i,Title,StrTemp
set rs_NewsYeZY=server.createobject("adodb.recordset")
sql_NewsYeZY="select top 10 * from Bs_NewsYe order by id desc"
rs_NewsYeZY.open sql_NewsYeZY,conn,1,1
i=0
	Do While Not rs_NewsYeZY.EOF
		Title=""
		if rs_NewsYeZY("Title")<>"" then
			StrTemp=replace(rs_NewsYeZY("Title"),"<b>","")
			StrTemp=replace(StrTemp,"&nbsp;","")
			Title=Title&StrTemp
		else
			Title=Title&"此栏目没有任何内容"
		end if

		Response.Write "<IMG src=""../Skin/Img/article_common.gif"" width=9 height=15 border=0> "
		Response.Write "<a href='Bs_NewsInfo.asp?Action=Ye&ID=" & rs_NewsYeZY("ID") & "'title='"& rs_NewsYeZY("Title") &"'>"&  gotTopic(Title,36) & "</a><BR>"
	rs_NewsYeZY.MoveNext
	i=i+1
	if i=10 then exit do
	Loop

rs_NewsYeZY.close
set rs_NewsYeZY=nothing
END Sub

'-------- --------
Sub NewsProductZY()
dim rs_NewsProductZY,sql_NewsProductZY,i,Title,StrTemp
set rs_NewsProductZY=server.createobject("adodb.recordset")
sql_NewsProductZY="select top 10 * from Bs_NewsPr order by ID desc"
rs_NewsProductZY.open sql_NewsProductZY,conn,1,1
i=0
	Do While Not rs_NewsProductZY.EOF
		Title=""
		if rs_NewsProductZY("Title")<>"" then
			StrTemp=replace(rs_NewsProductZY("Title"),"<b>","")
			StrTemp=replace(StrTemp,"&nbsp;","")
			Title=Title&StrTemp
		else
			Title=Title&"此栏目没有任何内容"
		end if

		Response.Write "<IMG src=""../Skin/Img/book.gif"" width=10 height=10 border=0> "
		Response.Write "<a href='Bs_NewsInfo.asp?Action=Pr&ID=" & rs_NewsProductZY("ID") & "'title='"& rs_NewsProductZY("Title") &"'>"& gotTopic(Title,36) & "</a><BR>"
	rs_NewsProductZY.MoveNext
	i=i+1
	if i=10 then exit do
	Loop

rs_NewsProductZY.close
set rs_NewsProductZY=nothing
END Sub
'-------- --------
Sub DownZY()
dim rs_DownZY,sql_DownZY,i,Bs_Title,tempstr
set rs_DownZY=server.createobject("adodb.recordset")
sql_DownZY="select top 10 * from Bs_Download where Bs_Passed=True order by Bs_UpDateTime desc"
rs_DownZY.open sql_DownZY,conn,1,1
i=0
	Do While Not rs_DownZY.EOF
		Bs_Title=""
		if rs_DownZY("Bs_Title")<>"" then
			tempstr=replace(rs_DownZY("Bs_Title"),"<b>","")
			tempstr=replace(tempstr,"&nbsp;","")
			Bs_Title=Bs_Title&tempstr
		else
			Bs_Title=Bs_Title&"此栏目没有任何内容"
		end if

		Response.Write "<IMG src=""../Skin/Img/article_elite.gif"" width=9 height=15 border=0> "
		Response.Write "<a href='bs_DownloadShow.asp?Bs_DownID=" & rs_DownZY("Bs_DownID") & "'title='"& rs_DownZY("Bs_Title") &"'>"& gotTopic(Bs_Title,36) & "</a><BR>"
	rs_DownZY.MoveNext
	i=i+1
	if i=10 then exit do
	Loop

rs_DownZY.close
set rs_DownZY=nothing
END Sub
%>

<%
'--------站内调查--------
Sub Style_Vote()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_tdptbg'><SPAN class=GlowLeft>站内调查</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD vAlign=top>
			<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr>
					<TD><% call ShowVote() %></TD>
				</tr>
			</table>
		</TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub
%>