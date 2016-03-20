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
Sub Style_Bulletin()
%>
<% if strSkins <= 200 and strSkins >= 1 then%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>网 站 公 告</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
 	<TR> 
		<TD vAlign=top><DIV align="center">
			<table width="168" height="120" border="0" cellPadding="0" cellSpacing="0" class="unnamed1">
				<tr>
				  <td align="center">
<marquee scrollamount='1'scrolldelay='170'direction= 'up' width='170'height='118'id="xiaoqing" onmouseover="xiaoqing.stop()" onmouseout="xiaoqing.start()">
            <%=CSJ_Bulletin%>
          </marquee></td>
				</tr>
			</table>
		</DIV></TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE></DIV>
<%
elseif strSkins <= 250 and strSkins >= 201 then
%>
<DIV align="center">
<table border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img border="0" src="../Skin/201/window/zs-glay-top.gif" width="206" height="48" /></td>
    </tr>
    <tr>
      <td background="../Skin/201/window/glay-mid.gif">
        <DIV align="center">
			<table width="168" height="120" border="0" cellPadding="0" cellSpacing="0" class="unnamed1">
				<tr>
				  <td align="center">
<marquee scrollamount='1'scrolldelay='170'direction= 'up' width='170'height='118'id="xiaoqing" onmouseover="xiaoqing.stop()" onmouseout="xiaoqing.start()"><%=CSJ_Bulletin%></marquee>
                  </td>
				</tr>
			</table>
		</DIV>
	  </td>
    </tr>
    <tr>
      <td><img border="0" src="../Skin/201/window/glay-buttom.gif" width="206" height="15" /></td>
    </tr>
</table>
<table cellpadding="0" cellspacing="0" border="0">
    <tr>
      <td height="8"><img src='../img/1x1_pix.gif' width="3" height="1" /></td>
    </tr>
</table>
</DIV>
<%end if%>
<%
End Sub
Sub Style_Login()
if strSkins <= 100 and strSkins >= 1 then
%>
<%
elseif strSkins <= 200 and strSkins >= 101 then
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>会 员 登 陆</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD vAlign=top>
			<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
					<td><% call ShowUserLogin() %></td>
				</tr>
			</table>
		</TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
elseif strSkins <= 250 and strSkins >= 201 then
%>
<DIV align="center">
  <% call ShowUserLogin4() %>
  <TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>

</DIV>
<%
end if
End Sub

Sub Style_CoProfile()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>公 司 简 介</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Profile">公 司 介 绍</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Ceo">总 裁 介 绍</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Culture">公 司 文 化</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Organize">组 织 机 构</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Principle">领 导 关 怀</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Contact">联 系 我 们</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_News()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>新 闻 中 心</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_News.asp?Action=Co">公 司 新 闻</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_News.asp?Action=Ye">业 内 资 讯</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_News.asp?Action=Pr">产 品 动 态</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Product()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>产 品 列 表</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD vAlign=top><% call ShowSmallClass_Tree() %></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Search()
%>
<DIV align="center">
<%if strSkins <= 200 and strSkins >= 1 then%>
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>产 品 搜 索</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD vAlign=top><% call ShowSearch(1) %></TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<table border="0" cellspacing="0" cellpadding="0">
  
  <tr>
    <td><img border="0" src="../Skin/201/window/search-glay-top.gif" width="206" height="48" /></td>
  </tr>
  <tr>
    <td background="../Skin/201/window/glay-mid.gif"><div align="center">
        <table width="92%" border="0" align="center" cellpadding="2" cellspacing="0">
          <tr>
            <td align="left"><form action="../chinese/Bs_Search.asp" method="post">
                <font color="#666666">请输入搜索关键字：<br />
                <input name="keyword" type="text" class="form" id="keyword" size="19" />
                <input type="submit" value="搜 索 " name="sub" class="form" />
                </font>
                </p>
            </form></td>
          </tr>
      </table></td>
  </tr>
  <tr>
    <td><img border="0" src="../Skin/201/window/glay-buttom.gif" width="206" height="15" /></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" border="0">
  <tr></tr>
  <tr>
    <td height="8"><img src='../img/1x1_pix.gif' width="3" height="1" /></td>
  </tr>
</table>
<% end if%>
</DIV>
<%
End Sub

Sub Style_Honor()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>公 司 荣 誉</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Honor.asp?Action=Honor">公 司 荣 誉</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Honor.asp?Action=Img">公 司 形 象</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Sale()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>营 销 网 络</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Sale.asp?Action=Sale">国 内 市 场</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Sale.asp?Action=Salea">国 际 市 场</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Job()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>人 才 招 聘</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Job.asp">人 才 招 聘</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Jobs.asp">人 才 策 略</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Download()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>下 载 列 表</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD vAlign=top><% call Show_DownSmallClass_Tree() %></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Member()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>会 员 中 心</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Faq.asp">常见问题解答</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Went.asp">用户信息反馈</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Manage.asp">修改会员资料</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_ManagePwd.asp">修改会员密码</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_E_shop.asp">购物订单查询</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_NetBook.asp">站内留言中心</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_UserLogout.asp">退出会员中心</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Mail()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>邮 件 列 表</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<form method="POST" action="Mail_chk.asp">
			<TD vAlign=top>
				<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
					<tr> 
						<td height="25"> <div align="center"><input name='email'type='text'value="@" size='15' class=form></div></td>
					</tr>
					<tr> 
						<td height="25"><div align="center"> 
							<input type="submit" value="订阅" name="action">
							&nbsp; 
							<input type="submit" value="退订" name="action">
						</div></td>
					</tr>
				</table>
			</TD>
		</form>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Fellow()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>友 情 连 接</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<%
dim rs_links,sqltext4,i
set rs_links=server.createobject("adodb.recordset")
sqltext4="select top 10 * from Bs_Links order by id desc"
rs_links.open sqltext4,conn,1,1

i=0
do while not rs_links.eof
%>
				<TR vAlign=center>
					<TD align='center' class='ML_Tdbg'>
					<a href="<%=rs_links("link")%>" title="<%=rs_links("note")%>"target="_blank"><%=rs_links("name")%></a></TD>
				</TR>
				<TR><TD class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<%
rs_links.movenext
i=i+1
if i=10 then exit do
loop
rs_links.close 
%>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub
Sub Style_Adcolumn()
%>
<SCRIPT LANGUAGE="JavaScript" src="../Skin/<%=strSkins%>/BsImgLogo.js"></SCRIPT>
<%
End Sub
'首页
Sub Left_Index() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
Call Style_Bulletin() '网站公告
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_CoProfile() '公司简介
Call Style_Product() '产品列表
Call Style_Mail() '邮件列表
Call Style_Fellow() '友情连接
End Sub

'首页右
Sub Right_Index() 
Call Style_Login() '会员登陆
Call Style_Bulletin() '网站公告
Call Style_Search() '产品搜索
End Sub

'下载中心
Sub Right_Download() 
Call Style_Login() '会员登陆
Call Style_Search() '产品搜索
End Sub

'公司简介
Sub Left_CoProfile() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_CoProfile() '公司简介
Call Style_Mail() '邮件列表
Call Style_Fellow() '友情连接
End Sub

'新闻资讯
Sub Left_News() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_News() '新闻中心
Call Style_Mail() '邮件列表
Call Style_Fellow() '友情连接
End Sub

'产品展示
Sub Left_Product() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Product() '产品列表
Call Style_Fellow() '友情连接
End Sub

'公司荣誉
Sub Left_Honor() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Honor() '公司荣誉
Call Style_Mail() '邮件列表
Call Style_Fellow() '友情连接
End Sub

'营销网络
Sub Left_Sale() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Sale() '营销网络
Call Style_Mail() '邮件列表
Call Style_Fellow() '友情连接
End Sub

'人才招聘
Sub Left_Job() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Job() '人才招聘
Call Style_Mail() '邮件列表
Call Style_Fellow() '友情连接
End Sub

'下载中心
Sub Left_Download() '下载列表
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Download() '产品列表
Call Style_Fellow() '友情连接
End Sub

'会员中心
Sub Left_Member() 
Call Style_Member() '会员中心
Call Style_Mail() '邮件列表
Call Style_Fellow() '友情连接
End Sub

'call Style_Bulletin()
%>

