
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
		<a target="_top" href="../../../../../../../../../../../../../index.asp"><font color="333333">�ص���ҳ</font></a></b> 
		| <a target="_top" href="Loginout.asp"><font color="333333"><b>�˳�
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
		<span>ϵͳ���� </span></td>
	</tr>

	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_SysData.asp"><font color="000000">��վ��������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Admin.asp"><font color="000000">����Ա����</font></a></td></tr>
				<tr><td><a target="main" href="QQsetup.asp"><font color="000000">������ѯ����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Bulletin.asp"><font color="000000">��վ����</font></a></td></tr>
				<tr><td><a target="main" href="http://bbs.asbd.cn"><font color="000000">ϵͳ����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Upload_File.asp"><font color="000000">�ϴ��ļ�����</font></a></td></tr>
				<tr><td><a target="main" href="sysadmin_count_list.asp"><font color="000000">��վ�������</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu2" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(2)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>��˾��Ϣ </span></td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Profile"><font color="000000">��˾���</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Ceo"><font color="000000">�ܲý���</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Culture"><font color="000000">��˾�Ļ�</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Organize"><font color="000000">��֯����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Principle"><font color="000000">�쵼�ػ�</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Contact"><font color="000000">��ϵ����</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu3" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(3)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>��Ʒ���� </span></td>
	</tr>
	<tr>
		<td id="submenu3" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Class.asp"><font color="000000">��Ʒ�������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Product.asp"><font color="000000">��Ʒ����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Product_Add.asp"><font color="000000">��Ӳ�Ʒ</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Product_Check.asp"><font color="000000">��˲�Ʒ</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu4" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(4)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>���ع��� </span></td>
	</tr>
	<tr>
		<td id="submenu4" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_DownClass.asp"><font color="000000">�����������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Download.asp"><font color="000000">���ع���</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Download_Add.asp"><font color="000000">�������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Download_Check.asp"><font color="000000">�������</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu5" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(5)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>�������� </span></td>
	</tr>
	<tr>
		<td id="submenu5" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Eshop.asp"><font color="000000">��������</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu6" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(6)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>��Ա���� </span></td>
	</tr>
	<tr>
		<td id="submenu6" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_User.asp"><font color="000000">ע���Ա����</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu7" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(7)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>���Ź��� </span></td>
	</tr>
	<tr>
		<td id="submenu7" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Co"><font color="000000">��˾���Ź���</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Co"><font color="000000">��ӹ�˾����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Ye"><font color="000000">ҵ����Ѷ����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Ye"><font color="000000">���ҵ����Ѷ</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Pr"><font color="000000">��Ʒ��̬����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Pr"><font color="000000">��Ӳ�Ʒ��̬</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News.asp?UrlName=Faq"><font color="000000">���������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_News_Add.asp?UrlName=Faq"><font color="000000">���������</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>
	
	<tr>
		<td id="imgmenu8" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(8)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>���Թ��� </span></td>
	</tr>
	<tr>
		<td id="submenu8" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Book.asp"><font color="000000">���Թ���</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Book_Add.asp"><font color="000000">����Ա����</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu9" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(9)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>�������� </span></td>
	</tr>
	<tr>
		<td id="submenu9" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Honor.asp"><font color="000000">��˾��������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Honor_add.asp"><font color="000000">��ӹ�˾����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Img.asp"><font color="000000">��˾�������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Img_add.asp"><font color="000000">��ӹ�˾����</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu10" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(10)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>Ӫ������ </span></td>
	</tr>
	<tr>
		<td id="submenu10" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Sale"><font color="000000">�����г�</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Company.asp?UrlName=Salea"><font color="000000">�����г�</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu11" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(11)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>�˲Ź��� </span></td>
	</tr>
	<tr>
		<td id="submenu11" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Job.asp"><font color="000000">��Ƹ����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Job_Add.asp"><font color="000000">������Ƹ</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Job_Book.asp"><font color="000000">ӦƸ����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Jobs.asp"><font color="000000">�˲Ų���</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu12" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(12)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>������� </span></td>
	</tr>
	<tr>
		<td id="submenu12" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Vote.asp"><font color="000000">��������</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Vote_Add.asp"><font color="000000">����µ���</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu13" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(13)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>�ʼ��б� </span></td>
	</tr>
	<tr>
		<td id="submenu13" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Mail_default.asp"><font color="000000">�ʼ��б�����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Mail_Send.asp"><font color="000000">�����ʼ�</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Mail_View_user.asp"><font color="000000">�û�����</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Mail_Add_user.asp"><font color="000000">����û�</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu14" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(14)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>�������� </span></td>
	</tr>
	<tr>
		<td id="submenu14" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Link.asp"><font color="000000">�������ӹ���</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		<td id="imgmenu15" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(15)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25">
		<span>���ݿ����</span></td>
	</tr>
	<tr>
		<td id="submenu15" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr><td><a target="main" href="Bs_Database.asp?Action=Backup"><font color="000000">�������ݿ�</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Database.asp?Action=Restore"><font color="000000">�ָ����ݿ�</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Database.asp?Action=Compact"><font color="000000">ѹ�����ݿ�</font></a></td></tr>
				<tr><td><a target="main" href="Bs_Database.asp?Action=SpaceSize"><font color="000000">ϵͳ�ռ�ռ�����</font></a></td></tr>
			</table>
		</div>
		<br>
		</td>
	</tr>


</table>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="images/title_bg_quit.gif" height="25"> 
      <span>��Ȩ��Ϣ</span> </td>
	</tr>
	<tr>
		<td>
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
			<table cellspacing="4" cellpadding="3">
				<tr>
					
              <td height="20" style="line-height: 150%;"> 
				������ƣ���ҵ��վ����ϵͳMac3.0<br>
				�������֣�<a href="http://bbs.asbd.cn" target=_blank>BBS.ASBD.CN</a></td>
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
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//����100Ϊ�����ٶ�
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
top.document.title=" ��˾(��ҵ)��վ����ϵͳ"; 
</script>
</BODY></HTML>
