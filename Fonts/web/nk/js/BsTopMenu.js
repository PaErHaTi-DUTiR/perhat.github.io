//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.

//�˵�
var menuOffX=0	//�˵���������������˾���
var menuOffY=20	//�˵����������ֶ��˾���

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
//�˵�END

//cn
var index ='<div class=menuitems><a href=\"index.asp\">�� վ �� ҳ</a></div>'
		index+='<div class=menuitems><a href=\"../Default.asp\">��˾����ҳ</a></div>'
var CoProfile ='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Profile\">�� ˾ �� ��</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Ceo\">�� �� �� ��</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Culture\">�� ˾ �� ��</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Organize\">�� ֯ �� ��</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Principle\">�� �� �� ��</a></div>'
		CoProfile+='<div class=menuitems><a href=\"Bs_CoProfile.asp?Action=Contact\">�� ϵ �� ��</a></div>'
var News ='<div class=menuitems><a href=\"Bs_News.asp?Action=Co\">�� ˾ �� ��</a></div>'
		News+='<div class=menuitems><a href=\"Bs_News.asp?Action=Ye\">ҵ �� �� Ѷ</a></div>'
		News+='<div class=menuitems><a href=\"Bs_News.asp?Action=Pr\">�� Ʒ �� ̬</a></div>'
var Product ='<div class=menuitems><a href=\"Bs_Product.asp\">�� Ʒ չ ʾ</a></div>'
		Product+='<div class=menuitems><a href=\"Bs_Products.asp\">�� Ʒ �� ��</a></div>'
		Product+='<div class=menuitems><a href=\"Bs_Search.asp\">�� Ʒ �� ��</a></div>'
var Honor ='<div class=menuitems><a href=\"Bs_Honor.asp?Action=Honor\">�� ˾ �� ��</a></div>'
		Honor+='<div class=menuitems><a href=\"Bs_Honor.asp?Action=Img\">�� ˾ �� ��</a></div>'
var Sale ='<div class=menuitems><a href=\"Bs_Sale.asp?Action=Sale\">�� �� �� ��</a></div>'
		Sale+='<div class=menuitems><a href=\"Bs_Sale.asp?Action=Salea\">�� �� �� ��</a></div>'
var Job ='<div class=menuitems><a href=\"Bs_Job.asp\">�� �� �� Ƹ</a></div>'
		Job+='<div class=menuitems><a href=\"Bs_Jobs.asp\">�� �� �� ��</a></div>'
//var Download='<div class=menuitems><a href=\"Job.asp\">�� �� �� ��</a></div>'
var Server ='<div class=menuitems><a href=\"Bs_Faq.asp\">�� �� �� ��</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_Server.asp\">�� Ա �� ��</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_Went.asp\">�� Ϣ �� ��</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_NetBook.asp\">�� �� �� ��</a></div>'
		Server+='<div class=menuitems><a href=\"Bs_E_shop.asp\">�� �� �� ѯ</a></div>'
var Skins ='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=1\">��ɫ���</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=2\">ˮ�{���</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=3\">��ɫ���</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=4\">��ɫ���</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=5\">��ɫ���</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=6\">��ɫ���</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=7\">��� 007</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=8\">�й����״��</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=101\">��� 101</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=151\">��� 151</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=152\">��� 152</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=161\">��� 161</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=162\">��� 162</a></div>'
		Skins+='<div class=menuitems><a href=\"cookies.asp?menu=BsSkins&no=201\">��� 201</a></div>'
		

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
//���ߵ����Ĳ˵�
function CnTopMenu(JsSkins){
document.write ('<DIV align="center">');
document.write ('<TABLE cellPadding=0 cellSpacing=0 class="TopMenu">');
document.write ('<TR><TD class="TopMenu_Td1"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR><TD class="TopMenu_Td2"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR>');
document.write ('<TD class="TopMenu_Td3">');
document.write ('<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>');
document.write ('<TR>'); 
document.write ('<TD width="98%" valign="bottom" align="center">��');
document.write ('<a class="Top" href="index.asp" onMouseOver=\'ShowMenu(event,index)\'>�� ҳ</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_CoProfile.asp?Action=Profile" onMouseOver=\'ShowMenu(event,CoProfile)\'>��˾���</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_News.asp?Action=Co" onMouseOver=\'ShowMenu(event,News)\'>������Ѷ</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_Product.asp" onMouseOver=\'ShowMenu(event,Product)\'>��Ʒչʾ</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_Honor.asp?Action=Honor" onMouseOver=\'ShowMenu(event,Honor)\'>��˾����</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_Sale.asp?Action=Sale" onMouseOver=\'ShowMenu(event,Sale)\'>Ӫ������</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_Job.asp" onMouseOver=\'ShowMenu(event,Job)\'>�˲���Ƹ</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_Download.asp">��������</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="Bs_Faq.asp" onMouseOver=\'ShowMenu(event,Server)\'>����֧��</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href="cookies.asp?menu=BsSkins&no=1" onMouseOver=\'ShowMenu(event,Skins)\'>�� ��</a>');
document.write ('&nbsp;&nbsp;</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</DIV>');
}
//���ߵ�Ӣ�Ĳ˵�
function EnTopMenu(JsSkins){
document.write ('<DIV align="center">');
document.write ('<TABLE cellPadding=0 cellSpacing=0 class="TopMenu">');
document.write ('<TR><TD class="TopMenu_Td1"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR><TD class="TopMenu_Td2"><IMG src="../img/1x1_pix.gif" width=3 height=1></TD></TR>');
document.write ('<TR>');
document.write ('<TD class="TopMenu_Td3">');
document.write ('<TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>');
document.write ('<TR>'); 
document.write ('<TD width="98%" valign="bottom" align="center">��');
document.write ('<a class="Top" href=\"index.asp\" onMouseOver=\'ShowMenu(event,En_index)\'>Home</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href=\"En_CoProfile.asp?Action=Profile\" onMouseOver=\'ShowMenu(event,En_CoProfile)\'>Co,Profile</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href=\"En_Product.asp\" onMouseOver=\'ShowMenu(event,En_Product)\'>Products</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href=\"En_Honor.asp?Action=Honor\" onMouseOver=\'ShowMenu(event,En_Honor)\'>Co,Honor</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href=\"En_Sale.asp?Action=Sale\" onMouseOver=\'ShowMenu(event,En_Sale)\'>Market</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href=\"En_Went.asp\">Feedback</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href=\"En_Server.asp\" onMouseOver=\'ShowMenu(event,En_Server)\'>Services</a>��'+IMG1+JsSkins+IMG2+'��');
document.write ('<a class="Top" href=\"cookies.asp?menu=BsSkins&no=1\" onMouseOver=\'ShowMenu(event,En_Skins)\'>Style</a>');
document.write ('&nbsp;&nbsp;</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</DIV>');
}
