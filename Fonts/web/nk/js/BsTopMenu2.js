<!--
//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.
document.write ('<DIV align="center">');
document.write ('<TABLE cellSpacing=0 cellPadding=0 class="TopMenu2">');
document.write ('<TR><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 class="TopMenu2A">');
document.write ('<TR><TD align="center"');
document.write ('><A onmouseover="javascript:id1.src=\'../Img/Menu/BsMenu_01.gif\';setmenu(1);" onmouseout="id1.src=\'../Img/Menu/BsMenu_01b.gif\'" href="index.asp"');
document.write ('><IMG id=id1 src="../Img/Menu/BsMenu_01b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id2.src=\'../Img/Menu/BsMenu_02.gif\';setmenu(2);" onmouseout="id2.src=\'../Img/Menu/BsMenu_02b.gif\'" href="Bs_CoProfile.asp?Action=Profile"');
document.write ('><IMG id=id2 src="../Img/Menu/BsMenu_02b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id3.src=\'../Img/Menu/BsMenu_03.gif\';setmenu(3);" onmouseout="id3.src=\'../Img/Menu/BsMenu_03b.gif\'" href="Bs_News.asp?Action=Co"');
document.write ('><IMG id=id3 src="../Img/Menu/BsMenu_03b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id4.src=\'../Img/Menu/BsMenu_04.gif\';setmenu(4);" onmouseout="id4.src=\'../Img/Menu/BsMenu_04b.gif\'" href="Bs_Product.asp"');
document.write ('><IMG id=id4 src="../Img/Menu/BsMenu_04b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id5.src=\'../Img/Menu/BsMenu_05.gif\';setmenu(5);" onmouseout="id5.src=\'../Img/Menu/BsMenu_05b.gif\'" href="Bs_Honor.asp?Action=Honor"');
document.write ('><IMG id=id5 src="../Img/Menu/BsMenu_05b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id6.src=\'../Img/Menu/BsMenu_06.gif\';setmenu(6);" onmouseout="id6.src=\'../Img/Menu/BsMenu_06b.gif\'" href="Bs_Sale.asp?Action=Sale"');
document.write ('><IMG id=id6 src="../Img/Menu/BsMenu_06b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id7.src=\'../Img/Menu/BsMenu_07.gif\';setmenu(7);" onmouseout="id7.src=\'../Img/Menu/BsMenu_07b.gif\'" href="Bs_Job.asp"');
document.write ('><IMG id=id7 src="../Img/Menu/BsMenu_07b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id8.src=\'../Img/Menu/BsMenu_08.gif\';setmenu(8);" onmouseout="id8.src=\'../Img/Menu/BsMenu_08b.gif\'" href="Bs_Download.asp"');
document.write ('><IMG id=id8 src="../Img/Menu/BsMenu_08b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id9.src=\'../Img/Menu/BsMenu_09.gif\';setmenu(9);" onmouseout="id9.src=\'../Img/Menu/BsMenu_09b.gif\'" href="Bs_Faq.asp"');
document.write ('><IMG id=id9 src="../Img/Menu/BsMenu_09b.gif" width=71 height=59 border=0></A');
document.write ('><IMG src="../Img/Menu/BsMenu_Space.gif" width=5 height=59 border=0');
document.write ('><A onmouseover="javascript:id10.src=\'../Img/Menu/BsMenu_10.gif\';setmenu(10);" onmouseout="id10.src=\'../Img/Menu/BsMenu_10b.gif\'" href="cookies.asp?menu=BsSkins&no=152"');
document.write ('><IMG id=id10 src="../Img/Menu/BsMenu_10b.gif" width=71 height=59 border=0></A');
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
			str+='<IMG src="../img/1x1_pix.gif" width=66 height=1>';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Co\">�� ˾ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Ye\">ҵ �� �� Ѷ</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_News.asp?Action=Pr\">�� Ʒ �� ̬</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 4: //�� Ʒ չ ʾ
			str='<div align="left">';
			str+='<IMG src="../img/1x1_pix.gif" width=142 height=1>';
			str+='<a class="Top" href=\"Bs_Product.asp\">�� Ʒ չ ʾ</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Products.asp\">�� Ʒ �� ��</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"Bs_Search.asp\">�� Ʒ �� ��</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
		case 5: //�� ˾ �� ��
			str='<div align="left">';
			str+='<IMG src="../img/1x1_pix.gif" width=223 height=1>';
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
			str+='<IMG src="../img/1x1_pix.gif" width=212 height=1>';
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
		case 10: //�� ��
			str='<div align="right">';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=1\">��ɫ���</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=2\">ˮ�{���</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=4\">��ɫ���</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=6\">��ɫ���</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=101\">��� 101</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=151\">��� 151</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=152\">��� 152</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=161\">��� 161</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=162\">��� 162</a>&nbsp;|&nbsp;';
			str+='<a class="Top" href=\"cookies.asp?menu=BsSkins&no=201\">��� 201</a>&nbsp;|&nbsp;';
			str+='</div>';
			break;
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
document.write ('<TR><TD>');
document.write ('<TABLE cellSpacing=0 cellPadding=0 class="TopMenu2B">');
document.write ('<TR>');
document.write ('<TD id=menu class="TopMenu2B_Td1"></TD>');
document.write ('</TR></TABLE>');
document.write ('</TD></TR></TABLE>');
document.write ('</DIV>');
//-->
