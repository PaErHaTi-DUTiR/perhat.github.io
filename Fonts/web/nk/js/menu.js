<!--
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
var index='<div class=menuitems><a href=\"index.asp\">�� վ �� ҳ</a></div><div class=menuitems><a href=\"Default.asp\">��˾����ҳ</a></div>'
var CoProfile='<div class=menuitems><a href=\"CoProfile.asp\">�� ˾ �� ��</a></div><div class=menuitems><a href=\"CoCulture.asp\">�� ˾ �� ��</a></div><div class=menuitems><a href=\"CoOrganize.asp\">�� ֯ �� ��</a></div><div class=menuitems><a href=\"CoPrinciple.asp\">�� �� �� ��</a></div><div class=menuitems><a href=\"CoContact.asp\">�� ϵ �� ��</a></div>'
var News='<div class=menuitems><a href=\"News.asp\">�� ˾ �� ��</a></div><div class=menuitems><a href=\"yeNews.asp\">ҵ �� �� Ѷ</a></div>'
var Product= '<div class=menuitems><a href=\"Product.asp\">�� Ʒ չ ʾ</a></div><div class=menuitems><a href=\"Products.asp\">�� Ʒ �� ��</a></div><div class=menuitems><a href=\"search.asp\">�� Ʒ �� ��</a></div>'
var Honor='<div class=menuitems><a href=\"Honor.asp\">�� ˾ �� ��</a></div><div class=menuitems><a href=\"Img.asp\">�� ˾ �� ��</a></div>'
var Sale='<div class=menuitems><a href=\"Sale.asp\">�� �� �� ��</a></div><div class=menuitems><a href=\"Salea.asp\">�� �� �� ��</a></div>'
var Job='<div class=menuitems><a href=\"Job.asp\">�� �� �� Ƹ</a></div><div class=menuitems><a href=\"Jobs.asp\">�� �� �� ��</a></div>'
var Server='<div class=menuitems><a href=\"Server.asp\">�� Ա �� ��</a></div><div class=menuitems><a href=\"NetBook.asp\">�� �� �� ��</a></div><div class=menuitems><a href=\"E_shop.asp\">�� �� �� ѯ</a></div>'
//En
var En_index='<div class=menuitems><a href=\"En_index.asp\">HomePage</a></div><div class=menuitems><a href=\"Default.asp\">Home image</a></div>'

var En_CoProfile='<div class=menuitems><a href=\"En_CoProfile.asp\">Co,introduction</a></div><div class=menuitems><a href=\"En_CoCulture.asp\">Co,culture</a></div><div class=menuitems><a href=\"En_CoOrganize.asp\">Organization</a></div><div class=menuitems><a href=\"En_CoPrinciple.asp\">Spiritual idea</a></div><div class=menuitems><a href=\"En_CoContact.asp\">Contact us</a></div>'//Spiritual idea �� �� �� ��Principle

var En_News='<div class=menuitems><a href=\"#\">Co,news</a></div><div class=menuitems><a href=\"#\">Trade info</a></div>'

var En_Product='<div class=menuitems><a href=\"En_Product.asp\">Product show</a></div><div class=menuitems><a href=\"En_Products.asp\">Product class</a></div><div class=menuitems><a href=\"En_search.asp\">Product Search</a></div>'

var En_Honor='<div class=menuitems><a href=\"En_Honor.asp\">Company honor</a></div><div class=menuitems><a href=\"En_Img.asp\">Corporate image</a></div>'

var En_Sale='<div class=menuitems><a href=\"En_Sale.asp\">China market</a></div><div class=menuitems><a href=\"En_Salea.asp\">International market</a></div>'

var En_Job='<div class=menuitems><a href=\"En_Job.asp\">Job recruitment</a></div><div class=menuitems><a href=\"En_Jobs.asp\">Job tactics</a></div>'

var En_Server='<div class=menuitems><a href=\"En_Server.asp\">Member centre</a></div><div class=menuitems><a href=\"En_NetBook.asp\">Leave word</a></div><div class=menuitems><a href=\"En_E_shop.asp\">Order inquiry</a></div>'
