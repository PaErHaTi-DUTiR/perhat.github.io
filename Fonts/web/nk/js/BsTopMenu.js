//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.

//菜单
var menuOffX=0	//菜单距连接文字最左端距离
var menuOffY=20	//菜单距连接文字顶端距离

var ie4=document.all&&navigator.userAgent.indexOf("Opera")==-1
var ns6=document.getElementById&&!document.all
function ShowMenu(e,vmenu,mod){
	which=vmenu
	menuobj=document.getElementById("popmenu")
	menuobj.thestyle=menuobj.style
	menuobj.innerHTML=which
	menuobj.contentwidth=menuobj.offsetWidth
	eventX=e.clientX
	eventY=e.clientY
	var rightedge=document.body.clientWidth-eventX
	var bottomedge=document.body.clientHeight-eventY

		if (rightedge<menuobj.contentwidth)
			menuobj.thestyle.left=document.body.scrollLeft+eventX-menuobj.contentwidth+menuOffX
		else
			menuobj.thestyle.left=ie4? ie_x(event.srcElement)+menuOffX : ns6? window.pageXOffset+eventX : eventX
		
		if (bottomedge<menuobj.contentheight&&mod!=0)
			menuobj.thestyle.top=document.body.scrollTop+eventY-menuobj.contentheight-event.offsetY+menuOffY-23
		else
			menuobj.thestyle.top=ie4? ie_y(event.srcElement)+menuOffY : ns6? window.pageYOffset+eventY+10 : eventY

	menuobj.thestyle.visibility="visible"
}


function ie_y(e){  
	var t=e.offsetTop;  
	while(e=e.offsetParent){  
		t+=e.offsetTop;
	}  
	return t;  
}  
function ie_x(e){  
	var l=e.offsetLeft;  
	while(e=e.offsetParent){  
		l+=e.offsetLeft;  
	}  
	return l;  
}

function highlightmenu(e,state){
	if (document.all)
		source_el=event.srcElement
		while(source_el.id!="popmenu"){
			source_el=document.getElementById? source_el.parentNode : source_el.parentElement
			if (source_el.className=="menuitems"){
				source_el.id=(state=="on")? "mouseoverstyle" : ""
		}
	}
}
function hidemenu(){if (window.menuobj)menuobj.thestyle.visibility="hidden"}
function dynamichide(e){if ((ie4||ns6)&&!menuobj.contains(e.toElement))hidemenu()}
document.onclick=hidemenu
document.write("<div class=menuskin id=popmenu onmouseover=highlightmenu(event,'on') onmouseout=highlightmenu(event,'off');dynamichide(event)></div>");
//菜单END

//cn
var index ='<div class=menuitems><a href=\"index.asp\">网 站 首 页</a></div>'
		index+='<div class=menuitems><a href=\"../Default.asp\">公司形象页</a></div>'
var CoProfile ='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Profile\">公 司 介 绍</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Ceo\">总 裁 致 辞</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Culture\">公 司 文 化</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Organize\">组 织 机 构</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Principle\">精 神 理 念</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Contact\">联 系 我 们</a></div>'
var News ='<div class=menuitems><a href=\"Bs_News.asp?Action=Co\">公 司 新 闻</a></div>'
		News+='<div class=menuitems><a href=\"Bs_News.asp?Action=Ye\">业 内 资 讯</a></div>'
		News+='<div class=menuitems><a href=\"Bs_News.asp?Action=Pr\">产 品 动 态</a></div>'
var Product ='<div class=menuitems><a href=\"Bs_Product.asp\">产 品 展 示</a></div>'
		Product+='<div class=menuitems><a href=\"Bs_Products.asp\">产 品 分 类</a></div>'
		Product+='<div class=menuitems><a href=\"Bs_Search.asp\">产 品 搜 索</a></div>'
var Honor ='<div class=menuitems><a href=\"Bs_Honor.asp?Action=Honor\">公 司 荣 誉</a></div>'
		Honor+='<div class=menuitems><a href=\"Bs_Honor.asp?Action=Img\">公 司 形 象</a></div>'
var Sale ='<div class=menuitems><a href=\"Bs_Sale.asp?Action=Sale\">国 内 市 场</a></div>'
		Sale+='<div class=menuitems><a href=\"Bs_Sale.asp?Action=Salea\">国 际 市 场</a></div>'
var Job ='<div class=menuitems><a href=\"Bs_Job.asp\">人 才 招 聘</a></div>'
		Job+='<div class=menuitems><a href=\"Bs_Jobs.asp\">人 才 策 略</a></div>'
//var Download='<div class=menuitems><a href=\"Job.asp\">下 载 中 心</a></div>'
var Server ='<div class=menuitems><a href=\"Bs_Faq.asp\">问 题 解 答</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_Server.asp\">会 员 中 心</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_Went.asp\">信 息 反 馈</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_NetBook.asp\">留 言 中 心</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_E_shop.asp\">订 单 查 询</a></div>'
var Skins ='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=1\">橘色风格</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=2\">水藍风格</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=3\">红色风格</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=4\">绿色风格</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=5\">紫色风格</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=6\">灰色风格</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=7\">风格 007</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=8\">中国民俗大红</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=101\">风格 101</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=151\">风格 151</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=152\">风格 152</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=161\">风格 161</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=162\">风格 162</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=201\">风格 201</a></div>'
		

//En
var En_index ='<div class=menuitems><a href=\"index.asp\">HomePage</a></div>'
		En_index+='<div class=menuitems><a href=\"../Default.asp\">Home image</a></div>'
var En_CoProfile ='<div class=menuitems><a href=\"En_CoProfile.asp?Action=Profile\">Co,introduction</a></div>'
		En_CoProfile+='<div class=menuitems><a href=\"En_CoProfile.asp?Action=Ceo\">President oration</a></div>'
		En_CoProfile+='<div class=menuitems><a href=\"En_CoProfile.asp?Action=Culture\">Co,culture</a></div>'
		En_CoProfile+='<div class=menuitems><a href=\"En_CoProfile.asp?Action=Organize\">Organization</a></div>'
		En_CoProfile+='<div class=menuitems><a href=\"En_CoProfile.asp?Action=Principle\">Spiritual idea</a></div>'
		En_CoProfile+='<div class=menuitems><a href=\"En_CoProfile.asp?Action=Contact\">Contact us</a></div>'
var En_News ='<div class=menuitems><a href=\"#\">Co,news</a></div>'
		En_News+='<div class=menuitems><a href=\"#\">Trade info</a></div>'
var En_Product ='<div class=menuitems><a href=\"En_Product.asp\">Product show</a></div>'
		En_Product+='<div class=menuitems><a href=\"En_Products.asp\">Product class</a></div>'
		En_Product+='<div class=menuitems><a href=\"En_search.asp\">Product Search</a></div>'
var En_Honor ='<div class=menuitems><a href=\"En_Honor.asp?Action=Honor\">Company honor</a></div>'
		En_Honor+='<div class=menuitems><a href=\"En_Honor.asp?Action=Img\">Corporate image</a></div>'
var En_Sale ='<div class=menuitems><a href=\"En_Sale.asp?Action=Sale\">China market</a></div>'
		En_Sale+='<div class=menuitems><a href=\"En_Sale.asp?Action=Salea\">International market</a></div>'
var En_Job ='<div class=menuitems><a href=\"En_Job.asp\">Job recruitment</a></div>'
		En_Job+='<div class=menuitems><a href=\"En_Jobs.asp\">Job tactics</a></div>'
var En_Server ='<div class=menuitems><a href=\"En_Server.asp\">Member centre</a></div>'
		En_Server+='<div class=menuitems><a href=\"En_NetBook.asp\">Leave word</a></div>'
		En_Server+='<div class=menuitems><a href=\"En_E_shop.asp\">Order inquiry</a></div>'
var En_Skins ='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=1\">Orange style</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=2\">Blue style</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=3\">Red style</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=4\">Green style</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=5\">Purple style</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=6\">Gray style</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=101\">style 101</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=151\">style 151</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=152\">style 152</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=161\">style 161</a></div>'
		En_Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=201\">style 201</a></div>'
		

var IMG1="<IMG height=20 src=../Skin/"
var IMG2="/bgmenu_m.gif width=1 align=absMiddle>"
//顶边的中文菜单
function CnTopMenu(JsSkins){
document.write ('<DIV align="center">');
document.write ('<TABLE cellPadding=0 cellSpacing=0 class="TopMenu">');
document.write ('<TR><TD class="TopMenu_Td1"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR><TD class="TopMenu_Td2"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR>');
document.write ('<TD class="TopMenu_Td3">');
document.write ('<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>');
document.write ('<TR>'); 
document.write ('<TD width="98%" valign="bottom" align="center">　');
document.write ('<a class="Top" href="index.asp" onMouseOver=\'ShowMenu(event,index)\'>首 页</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_CoProfile.asp?Action=Profile" onMouseOver=\'ShowMenu(event,CoProfile)\'>公司简介</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_News.asp?Action=Co" onMouseOver=\'ShowMenu(event,News)\'>新闻资讯</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_Product.asp" onMouseOver=\'ShowMenu(event,Product)\'>产品展示</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_Honor.asp?Action=Honor" onMouseOver=\'ShowMenu(event,Honor)\'>公司荣誉</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_Sale.asp?Action=Sale" onMouseOver=\'ShowMenu(event,Sale)\'>营销网络</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_Job.asp" onMouseOver=\'ShowMenu(event,Job)\'>人才招聘</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_Download.asp">下载中心</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="Bs_Faq.asp" onMouseOver=\'ShowMenu(event,Server)\'>服务支持</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href="cookies.asp?menu=BsSkins&no=1" onMouseOver=\'ShowMenu(event,Skins)\'>风 格</a>');
document.write ('&nbsp;&nbsp;</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</DIV>');
}
//顶边的英文菜单
function EnTopMenu(JsSkins){
document.write ('<DIV align="center">');
document.write ('<TABLE cellPadding=0 cellSpacing=0 class="TopMenu">');
document.write ('<TR><TD class="TopMenu_Td1"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR><TD class="TopMenu_Td2"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR>');
document.write ('<TD class="TopMenu_Td3">');
document.write ('<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>');
document.write ('<TR>'); 
document.write ('<TD width="98%" valign="bottom" align="center">　');
document.write ('<a class="Top" href=\"index.asp\" onMouseOver=\'ShowMenu(event,En_index)\'>Home</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href=\"En_CoProfile.asp?Action=Profile\" onMouseOver=\'ShowMenu(event,En_CoProfile)\'>Co,Profile</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href=\"En_Product.asp\" onMouseOver=\'ShowMenu(event,En_Product)\'>Products</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href=\"En_Honor.asp?Action=Honor\" onMouseOver=\'ShowMenu(event,En_Honor)\'>Co,Honor</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href=\"En_Sale.asp?Action=Sale\" onMouseOver=\'ShowMenu(event,En_Sale)\'>Market</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href=\"En_Went.asp\">Feedback</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href=\"En_Server.asp\" onMouseOver=\'ShowMenu(event,En_Server)\'>Services</a>　'+IMG1+JsSkins+IMG2+'　');
document.write ('<a class="Top" href=\"cookies.asp?menu=BsSkins&no=1\" onMouseOver=\'ShowMenu(event,En_Skins)\'>Style</a>');
document.write ('&nbsp;&nbsp;</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</DIV>');
}
