<!--
//=========================================================
//
// 产品名称： 公司(企业)网站管理系统（简称：www.web300.cn）
// 版权所有：www.web300.cn
// 程序制作：QQ812256@hotmail.com
// Copyright 2004-2005 www.web300.cn - All Rights Reserved. 
//========================================================

function setImgLogo(id)
{
	var str;
	switch (id)
	{
		case 1:  //首 页
			str='<div align="center">';
			str+='</div>';
			break;
		case 2:  //公 司 介 绍
			str='<div align="center">';
			str+='<IMG src="../Skin/Img/banner-7.jpg" width=556 height=89>';
			str+='</div>';
			break;
		case 3: //公 司 新 闻
			str='<div align="center">';
			str+='</div>';
			break;
		case 4: //产 品 展 示
			str='<div align="center">';
			str+='</div>';
			break;
		case 5: //公 司 荣 誉
			str='<div align="center">';
			str+='<IMG src="../Skin/Img/banner-4.jpg" width=556 height=89>';
			str+='</div>';
			break;
		case 6: //国 内 市 场
			str='<div align="center">';
			str+='<IMG src="../Skin/Img/banner-8.jpg" width=556 height=89>';
			str+='</div>';
			break;
		case 7: //人 才 招 聘
			str='<div align="center">';
			str+='<IMG src="../Skin/Img/banner-1.jpg" width=556 height=89>';
			str+='</div>';
			break;
		case 8: //下 载 中 心
			str='<div align="center">';
			str+='</div>';
			break;
		case 9: //服 务 支 持
			str='<div align="center">';
			str+='<IMG src="../Skin/Img/banner-5.jpg" width=556 height=89>';
			str+='</div>';
			break;
		case 10: //风 格
			str='<div align="center">';
			str+='</div>';
			break;
		case 11: //
			str='<div align="center">';
			str+='</div>';
			break;
	}	
	window.ImgLogo.innerHTML=str;
}

/////////////////////////////////显示当前页面的Logo/////////////////////////////////////////////
window.onload=function()
{
	var curPage=escape(location.href);//当前页文件名
	if((curPage.indexOf('index.asp')>-1))
	{
		setImgLogo(1);
	}
	else if((curPage.indexOf('Bs_CoProfile.asp')>-1)||(curPage.indexOf('En_CoProfile.asp')>-1))
	{
		setImgLogo(2);
	}
	else if((curPage.indexOf('Bs_News.asp')>-1)||(curPage.indexOf('Bs_NewsInfo.asp')>-1))
	{
		setImgLogo(3);
	}
	else if((curPage.indexOf('Bs_Product.asp')>-1)||(curPage.indexOf('Bs_Products.asp')>-1)||(curPage.indexOf('Bs_Search.asp')>-1)||(curPage.indexOf('En_Product.asp')>-1)||(curPage.indexOf('En_Products.asp')>-1)||(curPage.indexOf('En_Search.asp')>-1))
	{
		setImgLogo(4);
	}
	else if((curPage.indexOf('Bs_Honor.asp')>-1)||(curPage.indexOf('En_Honor.asp')>-1))
	{
		setImgLogo(5);
	}
	else if((curPage.indexOf('Bs_Sale.asp')>-1)||(curPage.indexOf('En_Sale.asp')>-1))
	{
		setImgLogo(6);
	}
	else if((curPage.indexOf('Bs_Job.asp')>-1)||(curPage.indexOf('Bs_Jobs.asp')>-1))
	{
		setImgLogo(7);
	}
	else if((curPage.indexOf('Bs_Download.asp')>-1)||(curPage.indexOf('Bs_DownloadShow.asp')>-1))
	{
		setImgLogo(8);
	}
	else if((curPage.indexOf('Bs_Server.asp')>-1)||(curPage.indexOf('Bs_Faq.asp')>-1)||(curPage.indexOf('Bs_Went.asp')>-1)||(curPage.indexOf('Bs_NetBook.asp')>-1)||(curPage.indexOf('Bs_E_shop.asp')>-1)||(curPage.indexOf('En_Server.asp')>-1)||(curPage.indexOf('En_Faq.asp')>-1)||(curPage.indexOf('En_Went.asp')>-1)||(curPage.indexOf('En_NetBook.asp')>-1)||(curPage.indexOf('En_E_shop.asp')>-1))
	{
		setImgLogo(9);
	}
	else if((curPage.indexOf('cookies.asp')>-1))
	{
		setImgLogo(10);
	}
	else
	{
		setImgLogo(11);
	}
}
document.write ('<TR><TD>');
document.write ('<DIV align="center">');
document.write ('<TABLE cellSpacing=0 cellPadding=0 class="MC_ImgLogo">');
document.write ('<TR>');
document.write ('<TD id=ImgLogo class="MC_ImgLogo_Td1"></TD>');
document.write ('</TR></TABLE>');
document.write ('</DIV>');
document.write ('</TD></TR>');
//-->