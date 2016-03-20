$(function(){
        //当滚动条的位置处于距顶部100像素以下时，跳转链接出现，否则消失
        $(function () {
            $(window).scroll(function(){
                if ($(window).scrollTop()>100){
                    $("#back-to-top").fadeIn(1500);
                }
                else
                {
                    $("#back-to-top").fadeOut(1500);
                }
            });

            //当点击跳转链接后，回到页面顶部位置

            $("#back-to-top").click(function(){
                $('body,html').animate({scrollTop:0},1000);
                return false;
            });
        });
    });
////********生成桌面快捷方式***********
function toDesktop(sUrl,sName){
try
{
var WshShell = new ActiveXObject("WScript.Shell");
var oUrlLink = WshShell.CreateShortcut(WshShell.SpecialFolders("Desktop") + "\\" + sName + ".url");
oUrlLink.TargetPath = sUrl;
oUrlLink.Save();
}
catch(e)
{
alert("当前IE安全级别不允许操作！");
}
}