<!--
//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.

document.write ('<DIV align="center">');
document.write ('<TABLE cellSpacing=0 cellPadding=0 class="TopMenu2">');
document.write ('<TR><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 class="TopMenu2A">');
document.write ('<TR><TD align="center"');
document.write ('><A onmouseover="javascript:id1.src=\'../Img/Menu/EnMenu_01.gif\';setmenu(1);" onmouseout="id1.src=\'../Img/Menu/EnMenu_01b.gif\'" href="index.asp"');
document.write ('><IMG id=id1 src="../Img/Menu/EnMenu_01b.gif" width=71 height=59 border=0></A');
document.write ('>&nbsp;&nbsp;<IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

document.write ('>&nbsp;&nbsp;<A onmouseover="javascript:id2.src=\'../Img/Menu/EnMenu_02.gif\';setmenu(2);" onmouseout="id2.src=\'../Img/Menu/EnMenu_02b.gif\'" href="En_CoProfile.asp?Action=Profile"');
document.write ('><IMG id=id2 src="../Img/Menu/EnMenu_02b.gif" width=71 height=59 border=0></A');
document.write ('>&nbsp;&nbsp;<IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

//document.write ('><A onmouseover="javascript:id3.src=\'../Img/Menu/EnMenu_03.gif\';setmenu(3);" onmouseout="id3.src=\'../Img/Menu/EnMenu_03b.gif\'" href="Bs_News.asp?Action=Co"');
//document.write ('><IMG id=id3 src="../Img/Menu/EnMenu_03b.gif" width=71 height=59 border=0></A');
//document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

document.write ('>&nbsp;&nbsp;<A onmouseover="javascript:id4.src=\'../Img/Menu/EnMenu_04.gif\';setmenu(4);" onmouseout="id4.src=\'../Img/Menu/EnMenu_04b.gif\'" href="En_Product.asp"');
document.write ('><IMG id=id4 src="../Img/Menu/EnMenu_04b.gif" width=71 height=59 border=0></A');
document.write ('>&nbsp;&nbsp;<IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

document.write ('>&nbsp;&nbsp;<A onmouseover="javascript:id5.src=\'../Img/Menu/EnMenu_05.gif\';setmenu(5);" onmouseout="id5.src=\'../Img/Menu/EnMenu_05b.gif\'" href="En_Honor.asp?Action=Honor"');
document.write ('><IMG id=id5 src="../Img/Menu/EnMenu_05b.gif" width=71 height=59 border=0></A');
document.write ('>&nbsp;&nbsp;<IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

document.write ('> <A onmouseover="javascript:id6.src=\'../Img/Menu/EnMenu_06.gif\';setmenu(6);" onmouseout="id6.src=\'../Img/Menu/EnMenu_06b.gif\'" href="En_Sale.asp?Action=Sale"');
document.write ('><IMG id=id6 src="../Img/Menu/EnMenu_06b.gif" width=71 height=59 border=0></A');
document.write ('> <IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

//document.write ('><A onmouseover="javascript:id7.src=\'../Img/Menu/EnMenu_07.gif\';setmenu(7);" onmouseout="id7.src=\'../Img/Menu/EnMenu_07b.gif\'" href="Bs_Job.asp"');
//document.write ('><IMG id=id7 src="../Img/Menu/EnMenu_07b.gif" width=71 height=59 border=0></A');
//document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

document.write ('>&nbsp;&nbsp;<A onmouseover="javascript:id8.src=\'../Img/Menu/EnMenu_11.gif\';setmenu(8);" onmouseout="id8.src=\'../Img/Menu/EnMenu_11b.gif\'" href="En_Went.asp"');
document.write ('><IMG id=id8 src="../Img/Menu/EnMenu_11b.gif" width=71 height=59 border=0></A');
document.write ('>&nbsp;&nbsp;<IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

document.write ('>&nbsp;&nbsp;<A onmouseover="javascript:id9.src=\'../Img/Menu/EnMenu_09.gif\';setmenu(9);" onmouseout="id9.src=\'../Img/Menu/EnMenu_09b.gif\'" href="En_Server.asp"');
document.write ('><IMG id=id9 src="../Img/Menu/EnMenu_09b.gif" width=71 height=59 border=0></A');
document.write ('>&nbsp;&nbsp;<IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');

document.write ('>&nbsp;&nbsp;<A onmouseover="javascript:id10.src=\'../Img/Menu/EnMenu_10.gif\';setmenu(10);" onmouseout="id10.src=\'../Img/Menu/EnMenu_10b.gif\'" href="cookies.asp?menu=BsSkins&no=152"');
document.write ('><IMG id=id10 src="../Img/Menu/EnMenu_10b.gif" width=71 height=59 border=0></A');
document.write ('></TD></TR></TABLE>');
document.write ('</TD></TR>');
document.write ('');

function setmenu(id)
{
	var str;
	switch (id)
	{
		case 1:  //首 页
			str='<div align="left">';
			str+='&nbsp;&nbsp;';
			str+='<a class="Top" href=\"index.asp\">HomePage</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"../Default.asp\">Home image</a>&nbsp;&nbsp;';
			str+='</div>';
			break;
		case 2:  //公 司 介 绍
			str='<div align="left">';
			str+='&nbsp;&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Profile\">Co,introduction</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Ceo\">President oration</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Culture\">Co,culture</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Organize\">Organization</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Principle\">Spiritual idea</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Contact\">Contact us</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 3: //公 司 新 闻
			str='<div align="left">';
			str+='<IMG src="../img/1x1_pix.gif" width=66 height=1>';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Co\">公 司 新 闻</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Ye\">业 内 资 讯</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Pr\">产 品 动 态</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 4: //产 品 展 示
			str='<div align="left">';
			str+='<IMG src="../img/1x1_pix.gif" width=142 height=1>';
			str+='<a class="Top" href=\"En_Product.asp\">Product show</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_Products.asp\">Product class</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_search.asp\">Product Search</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 5: //公 司 荣 誉
			str='<div align="left">';
			str+='<IMG src="../img/1x1_pix.gif" width=223 height=1>';
			str+='<a class="Top" href=\"En_Honor.asp?Action=Honor\">Company honor</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_Honor.asp?Action=Img\">Corporate image</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 6: //国 内 市 场
			str='<div align="center">';
			str+='<a class="Top" href=\"En_Sale.asp?Action=Sale\">China market</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_Sale.asp?Action=Salea\">International market</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 7: //人 才 招 聘
			str='<div align="right">';
			str+='<a class="Top" href=\"Bs_Job.asp\">人 才 招 聘</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Jobs.asp\">人 才 策 略</a>&nbsp;|&nbsp;';
			str+='<IMG src="../img/1x1_pix.gif" width=212 height=1>';
			str+='</div>';
			break;
		case 8: //信 息 反 馈
			str='<div align="right">';
//			str+='<a class="Top" href=\"Bs_Download.asp\">下 载 中 心</a>&nbsp;|&nbsp;'
			str+='</div>';
			break;
		case 9: //服 务 支 持
			str='<div align="right">';
			str+='<a class="Top" href=\"En_Server.asp\">Member centre</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_NetBook.asp\">Leave word</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_E_shop.asp\">Order inquiry</a>&nbsp;&nbsp;';
			str+='</div>';
			break;
		case 10: //风 格
			str='<div align="right">';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=2">Orange style</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=2\">Blue style</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=4\">Green style</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=6\">Gray style</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=101\">style 101</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=151\">style 151</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=152\">style 152</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=161\">style 161</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=162\">style 162</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=201\">style 201</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 11: //
			str='<div align="right">';
			str+='</div>';
			break;
	}	
	window.menu.innerHTML=str;
}

/////////////////////////////////显示当前页面的菜单/////////////////////////////////////////////
window.onload=function()
{
	var curPage=escape(location.href);//当前页文件名
	setmenu(11);
}
document.write ('<TR><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 class="TopMenu2B">');
document.write ('<TR>');
document.write ('<TD id=menu class="TopMenu2B_Td1"></TD>');
document.write ('</TR></TABLE>');
document.write ('</TD></TR></TABLE>');
document.write ('</DIV>');
//-->
