<!--
//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.
//�����ö��˵�����
document.write ('<DIV align="center">');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu">');
document.write ('<TR><TD class="TopMenu1"> </TD><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 border="0" class="TopMenu2A">');
document.write ('<TR><TD ');

document.write ('><A onmouseover="javascript:id1.src=\'../Skin/201/Menu/BsMenu_01-1.gif\';setmenu(1);" onmouseout="id1.src=\'../Skin/201/Menu/BsMenu_01.gif\'" href="index.asp"');
document.write ('><IMG id=id1 src="../Skin/201/Menu/BsMenu_01.gif"  border=0></A');

document.write ('><A onmouseover="javascript:id2.src=\'../Skin/201/Menu/BsMenu_02-1.gif\';setmenu(2);" onmouseout="id2.src=\'../Skin/201/Menu/BsMenu_02.gif\'" href="Bs_CoProfile.asp?Action=Profile"');
document.write ('><IMG id=id2 src="../Skin/201/Menu/BsMenu_02.gif"  border=0></A');
///�¼���Ŀ
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
		case 1:  //�� ҳ
			str='<div align="left">';
			str+='&nbsp;&nbsp;';
			str+='<a class="Top" href=\"index.asp\">�� վ �� ҳ</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"../index.htm\">��˾����ҳ</a>&nbsp;&nbsp;';
			str+='</div>';
			break;
		case 2:  //�� ˾ �� ��
			str='<div align="left">';
			str+='&nbsp;&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Profile\">�� ˾ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Ceo\">�� �� �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Culture\">�� ˾ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Organize\">�� ֯ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Principle\">�� �� �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_CoProfile.asp?Action=Contact\">�� ϵ �� ��</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 3: //�� ˾ �� ��
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=66 height=1>';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Co\">�� ˾ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Ye\">ҵ �� �� Ѷ</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Pr\">�� Ʒ �� ̬</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 4: //�� Ʒ չ ʾ
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=142 height=1>';
			str+='<a class="Top" href=\"Bs_Product.asp\">�� Ʒ չ ʾ</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Products.asp\">�� Ʒ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Search.asp\">�� Ʒ �� ��</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 5: //�� ˾ �� ��
			str='<div align="left">';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=223 height=1>';
			str+='<a class="Top" href=\"Bs_Honor.asp?Action=Honor\">�� ˾ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Honor.asp?Action=Img\">�� ˾ �� ��</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 6: //�� �� �� ��
			str='<div align="center">';
			str+='<a class="Top" href=\"Bs_Sale.asp?Action=Sale\">�� �� �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Sale.asp?Action=Salea\">�� �� �� ��</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 7: //�� �� �� Ƹ
			str='<div align="right">';
			str+='<a class="Top" href=\"Bs_Job.asp\">�� �� �� Ƹ</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Jobs.asp\">�� �� �� ��</a>&nbsp;|&nbsp;';
			str+='<IMG src="../Skin/201/1x1_pix.gif" width=212 height=1>';
			str+='</div>';
			break;
		case 8: //�� �� �� ��
			str='<div align="right">';
//			str+='<a class="Top" href=\"Bs_Download.asp\"  target="_blank">�� �� �� ��</a>&nbsp;|&nbsp;'
			str+='</div>';
			break;
		case 9: //�� �� ֧ ��
			str='<div align="right">';
			str+='<a class="Top" href=\"Bs_Faq.asp\">�� �� �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Server.asp\">�� Ա �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Went.asp\">�� Ϣ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_NetBook.asp\">�� �� �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_E_shop.asp\">�� �� �� ѯ</a>&nbsp;&nbsp;';
			str+='</div>';
			break;
		case 10: //��̳
			str='<div align="right">';
			str+='</div>';
			break;
			
//		case 10: //�� ��
//			str='<div align="right">';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=1\">��ɫ���</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=2\">ˮ�{���</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=4\">��ɫ���</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=6\">��ɫ���</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=101\">��� 101</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=151\">��� 151</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=152\">��� 152</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=161\">��� 161</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=162\">��� 162</a>&nbsp;|&nbsp;';
//			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=201\">��� 201</a>&nbsp;|&nbsp;';
//			str+='</div>';
//			break;
			
		case 11: //
			str='<div align="right">';
			str+='</div>';
			break;
	}	
	window.menu.innerHTML=str;
}

/////////////////////////////////��ʾ��ǰҳ��Ĳ˵�/////////////////////////////////////////////
window.onload=function()
{
	var curPage=escape(location.href);//��ǰҳ�ļ���
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
