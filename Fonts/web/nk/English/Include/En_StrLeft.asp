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
<DIV align="center">
<%if strSkins <= 200 and strSkins >= 1 then%>
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>Website bulletin</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
 	<TR> 
		<TD vAlign=top><DIV align="center">
			<table width="168" height="120" border="0" cellPadding="0" cellSpacing="0" class="unnamed1">
				<tr>
					<td align="center">
<marquee scrollamount='1'scrolldelay='160'direction= 'up' width='168'height='118'id=xiaoqing onmouseover=xiaoqing.stop() onmouseout=xiaoqing.start()>
<div>　　<%=CSJ_Bulletin%></div></marquee>
					</td>
				</tr>
			</table>
		</DIV></TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<table border="0" width="1"  cellspacing="0" cellpadding="0">
  <tr>
    <td><img border="0" src="../Skin/201/window/enzs-glay-top.gif" width="206" height="48" /></td>
  </tr>
  <tr>
    <td background="../Skin/201/window/glay-mid.gif"><table width="170" border="0" align="center" cellpadding="2" cellspacing="0" id="table17">
      <tr>
        <td align="left"><script language="JavaScript" src="js.asp?Nclassid=4" type="text/javascript"></script>        </td>
      </tr>
      <tr>
        <td align="center"><marquee scrollamount='1'scrolldelay='170'direction= 'up' width='170'height='118'id="xiaoqing" onmouseover="xiaoqing.stop()" onmouseout="xiaoqing.start()">
          <%=CSJ_Bulletin%>
        </marquee></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><img border="0" src="../Skin/201/window/glay-buttom.gif" width="206" height="15" /></td>
  </tr>
</table>
<%end if%>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub
Sub Style_Login()%>
<DIV align="center">
<%if strSkins <= 100 and strSkins >= 1 then%>
<%
elseif strSkins <= 200 and strSkins >= 101 then
%>
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>>Member login</SPAN></TD>
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
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<% call ShowUserLogin4() %>
<%end if%>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub
Sub Style_CoProfile()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>Co,Profile</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_CoProfile.asp?Action=Profile">Co,introduction</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_CoProfile.asp?Action=Ceo">President intro</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_CoProfile.asp?Action=Culture">Co,culture</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_CoProfile.asp?Action=Organize">Organization</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_CoProfile.asp?Action=Principle">Concern from leaders</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_CoProfile.asp?Action=Contact">Contact us</a></TD>
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
		<TD class='ML_TdPtbg'><SPAN class=Glow>Product List</SPAN></TD>
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
	<TR><TD height=15><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>Product Search</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD vAlign=top><% call ShowSearch(1) %></TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<%
elseif strSkins <= 250 and strSkins >= 201 then '新风格==============================================================
%>
<table border="0" width="1" cellspacing="0" cellpadding="0">
  <tr>
    <td><img border="0" src="../Skin/201/window/ensearch-glay-top.gif" width="206" height="48" /></td>
  </tr>
  <tr>
    <td background="../Skin/201/window/glay-mid.gif"><div align="center">
        <table width="92%" border="0" align="center" cellpadding="2" cellspacing="0">
          <tr>
            <td align="left"><form action="../English/en_Search.asp" method="post">
                <font color="#666666">Search for the key word：<br />
                <input name="keyword" type="text" class="form" id="keyword" size="15" />
                <input type="submit" value="Search" name="sub" class="form" />
                </font>
                </p>
                <font color="#666666">
            </form></td>
          </tr>
      </table></td>
  </tr>
  <tr>
    <td><img border="0" src="../Skin/201/window/glay-buttom.gif" width="206" height="15" /></td>
  </tr>
</table>
<%end if%>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Honor()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>Company honor</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_Honor.asp?Action=Honor">Company honor</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_Honor.asp?Action=Img">Corporate image</a></TD>
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
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>Marketing network</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_Sale.asp?Action=Sale">China market</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_Sale.asp?Action=Salea">International market</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
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
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>Member centre</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_Manage.asp">Revise member materials</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_ManagePwd.asp">Revise member password</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_E_shop.asp">Shopping order inquiry</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_NetBook.asp">Website leave word</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="En_UserLogout.asp">Close member centre</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
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
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_CoProfile() '公司简介
Call Style_Product() '产品列表
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
Call Style_Product() '产品列表
End Sub

'产品展示
Sub Left_Product() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Product() '产品列表
if strSkins <= 200 and strSkins >= 1 then
Call Style_Search() '产品搜索
elseif strSkins <= 250 and strSkins >= 201 then 
end if
End Sub

'公司荣誉
Sub Left_Honor() 
if strSkins <= 200 and strSkins >= 1 then

Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Honor() '公司荣誉
Call Style_CoProfile() '公司简介
End Sub

'营销网络
Sub Left_Sale() 
if strSkins <= 200 and strSkins >= 1 then

Call Style_Login() '会员登陆
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Sale() '营销网络
Call Style_Product() '产品列表
End Sub

'会员中心
Sub Left_Member() 

Call Style_Member() '会员中心
End Sub
%>