<!--
//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.
//中文置顶菜单设置
document.write ('<DIV align="center">');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu">');
document.write ('<TR><TD class="TopMenu1"> </TD><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu2A">');
document.write ('<TR><TD ');

document.write ('><A onmouseover="javascript:id1.src=\'../Skin/201/Menu/BsMenu_01-1.gif\';setmenu(1);" onmouseout="id1.src=\'../Skin/201/Menu/BsMenu_01.gif\'" href="index.asp"');
document.write ('><IMG id=id1 src="../Skin/201/Menu/BsMenu_01.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id2.src=\'../Skin/201/Menu/BsMenu_02-1.gif\';setmenu(2);" onmouseout="id2.src=\'../Skin/201/Menu/BsMenu_02.gif\'" href="Bs_CoProfile.asp?Action=Profile"');
document.write ('><IMG id=id2 src="../Skin/201/Menu/BsMenu_02.gif"  border=0></A');
///新加项目
//document.write ('><A onmouseover="javascript:id10.src=\'../Skin/201/Menu/BsMenu_10-1.gif\';setmenu(10);" onmouseout="id10.src=\'../Skin/201/Menu/BsMenu_10.gif\'" href="Bs_CoProfile.asp?Action=Ceo\"');
//document.write ('><IMG id=id10 src="../Skin/201/Menu/BsMenu_10.gif"  border=0></A');
///
document.write ('><A onmouseover="javascript:id3.src=\'../Skin/201/Menu/BsMenu_03-1.gif\';setmenu(3);" onmouseout="id3.src=\'../Skin/201/Menu/BsMenu_03.gif\'" href="Bs_News.asp?Action=Co"');
document.write ('><IMG id=id3 src="../Skin/201/Menu/BsMenu_03.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id4.src=\'../Skin/201/Menu/BsMenu_04-1.gif\';setmenu(4);" onmouseout="id4.src=\'../Skin/201/Menu/BsMenu_04.gif\'" href="Bs_Product.asp"');
document.write ('><IMG id=id4 src="../Skin/201/Menu/BsMenu_04.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id5.src=\'../Skin/201/Menu/BsMenu_05-1.gif\';setmenu(5);" onmouseout="id5.src=\'../Skin/201/Menu/BsMenu_05.gif\'" href="Bs_Honor.asp?Action=Honor"');
document.write ('><IMG id=id5 src="../Skin/201/Menu/BsMenu_05.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id6.src=\'../Skin/201/Menu/BsMenu_06-1.gif\';setmenu(6);" onmouseout="id6.src=\'../Skin/201/Menu/BsMenu_06.gif\'" href="Bs_Sale.asp?Action=Sale"');
document.write ('><IMG id=id6 src="../Skin/201/Menu/BsMenu_06.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id7.src=\'../Skin/201/Menu/BsMenu_07-1.gif\';setmenu(7);" onmouseout="id7.src=\'../Skin/201/Menu/BsMenu_07.gif\'" href="Bs_Job.asp"');
document.write ('><IMG id=id7 src="../Skin/201/Menu/BsMenu_07.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id8.src=\'../Skin/201/Menu/BsMenu_08-1.gif\';setmenu(8);" onmouseout="id8.src=\'../Skin/201/Menu/BsMenu_08.gif\'" href="Bs_Download.asp"');
document.write ('><IMG id=id8 src="../Skin/201/Menu/BsMenu_08.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id9.src=\'../Skin/201/Menu/BsMenu_09-1.gif\';setmenu(9);" onmouseout="id9.src=\'../Skin/201/Menu/BsMenu_09.gif\'" href="Bs_Faq.asp"');
document.write ('><IMG id=id9 src="../Skin/201/Menu/BsMenu_09.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id10.src=\'../Skin/201/Menu/enp.gif\';setmenu(10);" onmouseout="id10.src=\'../Skin/201/Menu/enp.gif\'" ###"');
document.write ('><IMG id=id10 src="../Skin/201/Menu/enp.gif"  border=0></A');

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
			str+='<a class="Top" href=\"index.asp\">网 站 首 页</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"../index.htm\">公司形象页</a>&nbsp;&nbsp;';
			str+='</div>';
			break;
		case 2:  //公 司 介 绍
			str='<div align="left">';
			str+='&nbsp;&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Profile\">公 司 介 绍</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Ceo\">总 裁 介 绍</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Culture\">公 司 文 化</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Organize\">组 织 机 构</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Principle\">领 导 关 怀</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Contact\">联 系 我 们</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 3: //公 司 新 闻
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=66 height=1>';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Co\">公 司 新 闻</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Ye\">业 内 资 讯</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Pr\">产 品 动 态</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 4: //产 品 展 示
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=142 height=1>';
			str+='<a class="Top" href=\"Bs_Product.asp\">产 品 展 示</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Products.asp\">产 品 分 类</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Search.asp\">产 品 搜 索</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 5: //公 司 荣 誉
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=223 height=1>';
			str+='<a class="Top" href=\"Bs_Honor.asp?Action=Honor\">公 司 荣 誉</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Honor.asp?Action=Img\">公 司 形 象</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 6: //国 内 市 场
			str='<div align="center">';
			str+='<a class="Top" href=\"Bs_Sale.asp?Action=Sale\">国 内 市 场</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Sale.asp?Action=Salea\">国 际 市 场</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 7: //人 才 招 聘
			str='<div align="right">';
			str+='<a class="Top" href=\"Bs_Job.asp\">人 才 招 聘</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Jobs.asp\">人 才 策 略</a>&nbsp;|&nbsp;';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=212 height=1>';
			str+='</div>';
			break;
		case 8: //下 载 中 心
			str='<div align="right">';
//			str+='<a class="Top" href=\"Bs_Download.asp\"  target="_blank">下 载 中 心</a>&nbsp;|&nbsp;'
			str+='</div>';
			break;
		case 9: //服 务 支 持
			str='<div align="right">';
			str+='<a class="Top" href=\"Bs_Faq.asp\">问 题 解 答</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Server.asp\">会 员 中 心</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Went.asp\">信 息 反 馈</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_NetBook.asp\">留 言 中 心</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_E_shop.asp\">订 单 查 询</a>&nbsp;&nbsp;';
			str+='</div>';
			break;
		case 10: //论坛
			str='<div align="right">';
			str+='</div>';
			break;
			
//		case 10: //风 格
//			str='<div align="right">';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=1\">橘色风格</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=2\">水藍风格</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=4\">绿色风格</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=6\">灰色风格</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=101\">风格 101</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=151\">风格 151</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=152\">风格 152</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=161\">风格 161</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=162\">风格 162</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=201\">风格 201</a>&nbsp;|&nbsp;';
//			str+='</div>';
//			break;
			
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
document.write ('<TR><TD class="TopMenu2"> </TD><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu2B">');
document.write ('<TR>');
document.write ('<TD id=menu class="TopMenu2B_Td1"></TD>');
document.write ('</TR></TABLE>');
document.write ('</TD></TR></TABLE>');
document.write ('</DIV>');
//-->
