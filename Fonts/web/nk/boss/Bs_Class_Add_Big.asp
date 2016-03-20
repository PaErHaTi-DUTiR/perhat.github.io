<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
'=========================================================
'
'产品名称：科技 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有：www.web300.cn
'程序制作：QQ812256@hotmail.com
'联系 方式：QQ ：812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>

<%
dim Action,BigClassName,EnBigClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
EnBigClassName=trim(request("EnBigClassName"))
if yzcv<>zcv then response.end
if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章大类名不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * from Bs_PrBigClasss Where BigClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>文章大类“" & BigClassName & "”已经存在！</li>"
		else
    	 	rs.addnew
     		rs("BigClassName")=BigClassName
     		rs("EnBigClassName")=EnBigClassName
    	 	rs.update
     		rs.Close
	     	set rs=Nothing
    	 	call CloseConn()
			Response.Redirect "Bs_Class.asp"  
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.BigClassName.focus();
    return false;
  }
}
</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="360" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>产 品 类 别 设 置</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
				<form name="form1" method="post" action="Bs_Class_Add_Big.asp" onsubmit="return checkBig()">
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center" align=right><strong>添加大类</strong></td>
          </tr>
          <tr bgcolor="#E3E3E3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#C0C0C0"> 
              <div align="right" align=right><strong>大类名称：</strong></div></td>
            <td width="218" bgcolor="#E3E3E3"> 
              <input name="BigClassName" type="text" size="20" maxlength="30">
            </td>
          </tr>
          <tr bgcolor="#E3E3E3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#C0C0C0"> 
              <div align="right" align=right><strong>English名称：</strong></div></td>
            <td width="218" bgcolor="#E3E3E3"> 
              <input name="EnBigClassName" type="text" size="20" maxlength="30">
            </td>
          </tr>
          <tr bgcolor="#C0C0C0" class="tdbg"> 
            <td height="22" align="center" bgcolor="#C0C0C0">&nbsp; </td>
            <td height="22" align="center" bgcolor="#E3E3E3"> 
              <div align="left">
                <input name="Action" type="hidden" id="Action" value="Add">
                <input name="Add" type="submit" value=" 添 加 ">
              </div>
						</td>
         </tr>
		   </table>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
end if
call CloseConn()
%>
