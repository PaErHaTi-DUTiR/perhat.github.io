//Copyright (c)2006-2008 BBS.ASBD.CN All rigths reserved.

//底部的中文菜单
function CnFootMenu(){
document.write ('<DIV align="center">');
document.write ('<TABLE cellPadding=0 cellSpacing=0 class="FootMenu">');
document.write ('<TR>');
document.write ('<TD class="FootMenu_Td">');
document.write ('<DIV align=center');
document.write ('> | <a href="index.asp" class="bottom">首 页</a');
document.write ('> | <a href="Bs_CoProfile.asp?Action=Profile" class="bottom">企业介绍</a');
document.write ('> | <a href="Bs_Product.asp" class="bottom">产品介绍</a');
document.write ('> | <a href="Bs_Server.asp" class="bottom">服务支持</a');
document.write ('> | <a href="#" class="bottom" LANGUAGE="javascript" onclick="Copyright()">版权声明</a');
document.write ('> | <a href="../boss/" target="_blank" class="bottom">管理进入</a');

document.write ('> | </DIV>');
document.write ('</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</DIV>');
}
//底部的英文菜单
function EnFootMenu(){
document.write ('<DIV align="center">');
document.write ('<TABLE cellPadding=0 cellSpacing=0 class="FootMenu">');
document.write ('<TR>');
document.write ('<TD class="FootMenu_Td">');
document.write ('<DIV align=center');
document.write ('> | <a href="index.asp" class="bottom">Home</a');
document.write ('> | <a href="En_CoProfile.asp?Action=Profile" class="bottom">Enterprise introduction</a');
document.write ('> | <a href="En_Product.asp" class="bottom">Product introduction</a');
document.write ('> | <a href="En_Server.asp" class="bottom">Service support</a');
//document.write ('> | <a href="#" target="_blank" class="bottom">Copyright statement</a');
document.write ('> | <a href="../boss/" target="_blank" class="bottom">Website manage</a');

document.write ('> | </DIV>');
document.write ('</TD>');
document.write ('</TR>');
document.write ('</TABLE>');
document.write ('</DIV>');
}

function Copyright()
{
  var arr = showModalDialog("Copyright.asp", "", "dialogWidth:680px; dialogHeight:500px; status:0");
}
