<%
if Request.QueryString("no")="eshop" then

title=request("title")
content=htmlencode2(Request("content"))

If title="" Then
response.write "SORRY <br>"
response.write "请输入更新内容的标题!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from "& strDataName &" "
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("content")=content
rs("counter")=0
rs.update
rs.close
response.redirect "Bs_News.asp?UrlName="& UrlName &""
end if

Str_M1 = "<SCRIPT LANGUAGE='JScript'>StrM1('$S01','$S02','$S03')</SCRIPT>"
%>
<SCRIPT LANGUAGE='JScript'>
function StrM1(S01,S02){
document.write("<BR> <table cellpadding=2 cellspacing=1 border=0 width=560 align=center class='a2'> <tr> <td class='a1'height=25 align=center><strong>"+S01+"</strong></td> </tr><tr class='a4'> <td align=center><br> <table width=550 border=0 cellspacing=0 cellpadding=0> <tr> <form method='post'action='Bs_News_Add.asp?UrlName="+S02+"&no=eshop'> <td><div align=center> <table width='100%'border=0 cellspacing=2 cellpadding=3> <tr> <td width='8%'height=25 bgcolor=#C0C0C0> <div align=center>标　题</div></td> <td width='62%'bgcolor=#C0C0C0><input type='text'name='title' class='inputtext'maxlength=80 size=50></td> </tr><tr> <td height=22 bgcolor=#E3E3E3> <div align=center>UBB代码</div></td> <td bgcolor=#E3E3E3>")}
function StrM2(){
document.write("</td></tr> <tr><td height=25 bgcolor=#C0C0C0><div align=center>内　容</div></td><td height=25 bgcolor=#C0C0C0><textarea name='content'cols='57'rows='12'wrap='PHYSICAL'></textarea></td></tr> <tr><td height=25 colspan=2 bgcolor=#E3E3E3><div align=center><input type='submit'value='确定'>　<input type='reset'value='取消'></div></td></tr> <tr><td colspan=2>图片上传</td></tr> <tr><td colspan=2><div align=center><iframe name='ad'frameborder=0 width='100%'height=48 scrolling=no src='../WebEdit/Upload_File.asp'></iframe></div></td> </tr> </table></div></td></form></tr></table><br></td></tr></table><BR>")}
</SCRIPT>
<!-- #include file="../Inc/Head.asp" -->
<%
Str_M1=replace(replace(Str_M1,"$S01",strTitleName),"$S02",UrlName)
Response.Write Str_M1
%>
<script src='Inc/eshopcode.js'></script> 
<!--#include file="../Inc/Ubb.inc" -->
<SCRIPT LANGUAGE='JScript'>StrM2()</SCRIPT>
<%htmlend%>
