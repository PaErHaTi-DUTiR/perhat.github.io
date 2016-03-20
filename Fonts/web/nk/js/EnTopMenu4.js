<!--
//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.
//英文版页面置顶菜单设置
document.write ('<DIV align="center">');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu">');
document.write ('<TR><TD class="TopMenu1"> </TD><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu2A">');
document.write ('<TR><TD ');

document.write ('><A onmouseover="javascript:id1.src=\'../Skin/201/Menu/EnMenu_01-1.gif\';setmenu(1);"onmouseout="id1.src=\'../Skin/201/Menu/EnMenu_01.gif\'" href="index.asp"');
document.write ('><IMG id=id1 src="../Skin/201/Menu/EnMenu_01.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id2.src=\'../Skin/201/Menu/EnMenu_02-1.gif\';setmenu(2);" onmouseout="id2.src=\'../Skin/201/Menu/EnMenu_02.gif\'" href="En_CoProfile.asp?Action=Profile"');
document.write ('><IMG id=id2 src="../Skin/201/Menu/EnMenu_02.gif"  border=0></A');
/////新加项目
document.write ('><A onmouseover="javascript:id3.src=\'../Skin/201/Menu/EnMenu_10-1.gif\';setmenu(3);" onmouseout="id3.src=\'../Skin/201/Menu/EnMenu_10.gif\'"href="En_CoProfile.asp?Action=Ceo\"');
document.write ('><IMG id=id3 src="../Skin/201/Menu/EnMenu_10.gif"  border=0></A');
/////
document.write ('><A onmouseover="javascript:id4.src=\'../Skin/201/Menu/EnMenu_03-1.gif\';setmenu(4);" onmouseout="id4.src=\'../Skin/201/Menu/EnMenu_03.gif\'" href="En_Product.asp"');
document.write ('><IMG id=id4 src="../Skin/201/Menu/EnMenu_03.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id5.src=\'../Skin/201/Menu/EnMenu_04-1.gif\';setmenu(5);" onmouseout="id5.src=\'../Skin/201/Menu/EnMenu_04.gif\'" href="En_Honor.asp?Action=Honor"');
document.write ('><IMG id=id5 src="../Skin/201/Menu/EnMenu_04.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id6.src=\'../Skin/201/Menu/EnMenu_05-1.gif\';setmenu(6);" onmouseout="id6.src=\'../Skin/201/Menu/EnMenu_05.gif\'" href="En_Sale.asp?Action=Sale"');
document.write ('><IMG id=id6 src="../Skin/201/Menu/EnMenu_05.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id8.src=\'../Skin/201/Menu/EnMenu_06-1.gif\';setmenu(8);" onmouseout="id8.src=\'../Skin/201/Menu/EnMenu_06.gif\'" href="En_Went.asp"');
document.write ('><IMG id=id8 src="../Skin/201/Menu/EnMenu_06.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id9.src=\'../Skin/201/Menu/EnMenu_07-1.gif\';setmenu(9);" onmouseout="id9.src=\'../Skin/201/Menu/EnMenu_07.gif\'" href="En_Server.asp"');
document.write ('><IMG id=id9 src="../Skin/201/Menu/EnMenu_07.gif"  border=0></A');

//document.write ('><A onmouseover="javascript:id10.src=\'../Skin/201/Menu/EnMenu_08.gif\';setmenu(10);" onmouseout="id10.src=\'../Skin/201/Menu/EnMenu_08.gif\'" href="cookies.asp?menu=BsSkins&no=152"');
//document.write ('><IMG id=id10 src="../Skin/201/Menu/EnMenu_08.gif"  border=0></A');

//document.write ('><IMG src="../Skin/201/Menu/EnMenu_09.gif"  border=0 ');

//document.write ('><IMG src="../Skin/201/Menu/EnMenu_10.gif"  border=0 ');

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
			str+='';
			str+='<a class="Top" href=\"index.asp\">HomePage</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"../index.htm\">Home image</a>';
			str+='</div>';
			break;
		case 2:  //公 司 介 绍
			str='<div align="left">';
			str+='';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Profile\">Co,introduction</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Ceo\">President intro</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Culture\">Co,culture</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Organize\">Organization</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Principle\">Concern from leaders</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_CoProfile.asp?Action=Contact\">Contact us</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 3: //CEO
			str='<div align="left">';
            str+='<a class="Top" href=\"En_CoProfile.asp?Action=Ceo\">President intro</a>&nbsp;|&nbsp;';
			str+='<a class="Top" >Mail to CEO</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 4: //产 品 展 示
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=142 height=1>';
			str+='<a class="Top" href=\"En_Product.asp\">Product show</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_Products.asp\">Product class</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"En_search.asp\">Product Search</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 5: //公 司 荣 誉
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=223 height=1>';
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
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=212 height=1>';
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
			str+='<a class="Top" href=\"En_E_shop.asp\">Order inquiry</a>';
			str+='</div>';
			break;
//		case 10: //风 格
//			str='<div align="right">';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=2">Orange style</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=2\">Blue style</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=4\">Green style</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=6\">Gray style</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=101\">style 101</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=151\">style 151</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=152\">style 152</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=161\">style 161</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=162\">style 162</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=201\">style 201</a>&nbsp;|&nbsp;';
//			str+='</div>';
//			break;
//		case 11: //
//			str='<div align="right">';
//			str+='</div>';
//			break;
	}	
	window.menu.innerHTML=str;
}

/////////////////////////////////显示当前页面的菜单/////////////////////////////////////////////
window.onload=function()
{
	var curPage=escape(location.href);//当前页文件名
	setmenu(11);
}
document.write ('<TR><TD class="TopMenu2"> </TD><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu2B">');
document.write ('<TR>');
document.write ('<TD id=menu class="TopMenu2B_Td1"></TD>');
document.write ('</TR></TABLE>');
document.write ('</TD></TR></TABLE>');
document.write ('</DIV>');
//-->
