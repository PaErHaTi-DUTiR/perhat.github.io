
<%
'=========================================================
'
'产品名称： 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有: www.web300.cn 
'程序制作：www.web300.cn开发团队
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<html><head>
<meta http-equiv=Content-Type content=text/html; charset=gb2312 charset=gb2312>
<link href="Inc/bs.css" rel="stylesheet" type="text/css">
<style type="text/css">
BODY {
	BACKGROUND:799ae1;
	MARGIN: 0px;
	background-color: #666666;
}

.sec_menu {
	BORDER-RIGHT: #9f9f9f 1px solid; BACKGROUND: #cccccc; OVERFLOW: hidden; BORDER-LEFT: #9f9f9f 1px solid; BORDER-BOTTOM: #9f9f9f 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #333333; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #666666; POSITION: relative; TOP: 2px
}
a:link {
	color: #333333;
}
a:visited {
	color: #333333;
}
a:hover {
	color: #333333;
}
a:active {
	color: #333333;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<BODY>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td valign="bottom" height="42">
		<img height="38" src="images/title.gif" width="158" border="0"></td>
	</tr>
	<tr>
		<td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="images/title_bg_quit.gif" height="25">
		<span><b class="menu_title2">
		<a target="_top" href="../../../../../../../../../../../../../index.asp"><font color="333333">回到首页</font></a></b> 
		| <a target="_top" href="Loginout.asp"><font color="333333"><b>退出
		</font></a></b></span></td>
	</tr>
	<tr>
		<td align="center">
		<font face="Webdings" color="#FFFFFF">5</font> </td>
	</tr>
</table>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td id="imgmenu1" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(1);loadThreadFollow()" onMouseOut="this.className='menu_title';" style="cursor:hand" background=images/menudown.gif height="25">
		<span>系统管理 </span></td>
	</tr>

	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_SysData.asp"><font color="000000">网站资料设置</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Admin.asp"><font color="000000">管理员管理</font></a></td></tr>
				<tr><td><a target="main" href="QQsetup.asp"><font color="000000">在线咨询设置</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Bulletin.asp"><font color="000000">网站公告</font></a></td></tr>
				<tr><td><a target="main" href="http://bbs.asbd.cn"><font color="000000">系统帮助</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Upload_File.asp"><font color="000000">上传文件管理</font></a></td></tr>
				<tr><td><a target="main" href="sysadmin_count_list.asp"><font color="000000">网站访问情况</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu2" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(2)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>公司信息 </span></td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Profile"><font color="000000">公司简介</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Ceo"><font color="000000">总裁介绍</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Culture"><font color="000000">公司文化</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Organize"><font color="000000">组织机构</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Principle"><font color="000000">领导关怀</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Contact"><font color="000000">联系我们</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu3" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(3)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>产品管理 </span></td>
	</tr>
	<tr>
		<td id="submenu3" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Class.asp"><font color="000000">产品类别设置</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Product.asp"><font color="000000">产品管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Product_Add.asp"><font color="000000">添加产品</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Product_Check.asp"><font color="000000">审核产品</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu4" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(4)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>下载管理 </span></td>
	</tr>
	<tr>
		<td id="submenu4" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_DownClass.asp"><font color="000000">下载类别设置</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Download.asp"><font color="000000">下载管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Download_Add.asp"><font color="000000">添加下载</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Download_Check.asp"><font color="000000">审核下载</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu5" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(5)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>订单管理 </span></td>
	</tr>
	<tr>
		<td id="submenu5" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Eshop.asp"><font color="000000">订单管理</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu6" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(6)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>会员管理 </span></td>
	</tr>
	<tr>
		<td id="submenu6" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_User.asp"><font color="000000">注册会员管理</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu7" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(7)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>新闻管理 </span></td>
	</tr>
	<tr>
		<td id="submenu7" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Co"><font color="000000">公司新闻管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Co"><font color="000000">添加公司新闻</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Ye"><font color="000000">业内资讯管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Ye"><font color="000000">添加业内资讯</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Pr"><font color="000000">产品动态管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Pr"><font color="000000">添加产品动态</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Faq"><font color="000000">问题解答管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Faq"><font color="000000">添加问题解答</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>
	
	<tr>
		<td id="imgmenu8" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(8)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>留言管理 </span></td>
	</tr>
	<tr>
		<td id="submenu8" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Book.asp"><font color="000000">留言管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Book_Add.asp"><font color="000000">管理员公告</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu9" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(9)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>荣誉管理 </span></td>
	</tr>
	<tr>
		<td id="submenu9" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Honor.asp"><font color="000000">公司荣誉管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Honor_add.asp"><font color="000000">添加公司荣誉</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Img.asp"><font color="000000">公司形象管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Img_add.asp"><font color="000000">添加公司形象</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu10" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(10)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>营销网络 </span></td>
	</tr>
	<tr>
		<td id="submenu10" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Sale"><font color="000000">国内市场</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Salea"><font color="000000">国际市场</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu11" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(11)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>人才管理 </span></td>
	</tr>
	<tr>
		<td id="submenu11" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Job.asp"><font color="000000">招聘管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Job_Add.asp"><font color="000000">发布招聘</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Job_Book.asp"><font color="000000">应聘管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Jobs.asp"><font color="000000">人才策略</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu12" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(12)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>调查管理 </span></td>
	</tr>
	<tr>
		<td id="submenu12" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Vote.asp"><font color="000000">调查设置</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Vote_Add.asp"><font color="000000">添加新调查</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu13" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(13)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>邮件列表 </span></td>
	</tr>
	<tr>
		<td id="submenu13" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Mail_default.asp"><font color="000000">邮件列表设置</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Mail_Send.asp"><font color="000000">发送邮件</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Mail_View_user.asp"><font color="000000">用户管理</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Mail_Add_user.asp"><font color="000000">添加用户</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu14" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(14)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>友情链接 </span></td>
	</tr>
	<tr>
		<td id="submenu14" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Link.asp"><font color="000000">友情链接管理</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu15" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(15)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>数据库管理</span></td>
	</tr>
	<tr>
		<td id="submenu15" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Database.asp?Action=Backup"><font color="000000">备份数据库</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Database.asp?Action=Restore"><font color="000000">恢复数据库</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Database.asp?Action=Compact"><font color="000000">压缩数据库</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Database.asp?Action=SpaceSize"><font color="000000">系统空间占用情况</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>


</table>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="images/title_bg_quit.gif" height="25"> 
      <span>版权信息</span> </td>
	</tr>
	<tr>
		<td>
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
			<table cellspacing="4" cellpadding="3">
				<tr>
					
              <td height="20" style="line-height: 150%;"> 
				软件名称：企业网站管理系统Mac3.0<br>
				技术技持：<a href="http://bbs.asbd.cn" target=_blank>BBS.ASBD.CN</a></td>
				</tr>
			</table>
			</div>
		</div>
		</td>
	</tr>
</table>
</div>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td align="center" valign="bottom">
		<font face="Webdings" color="#FFFFFF">6</font> </td>
	</tr>
</table>
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}

function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function expands(el) {
whichEl1=eval(el+"Child");
if (whichEl1.style.display=="none"){
initIt();
whichEl1.style.display="block";
}else{whichEl1.style.display="none";}
}
var tree= 0;
function loadThreadFollow(){
if (tree==0){
document.frames["hiddenframe"].location.replace("LeftTree.asp");
tree=1
}
}

function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="images/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="images/menudown.gif";
}
}

function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
top.document.title=" 公司(企业)网站管理系统"; 
</script>
</BODY></HTML>
