<!--#include file="bsconfig.asp"-->

<%
'=========================================================
'
'产品名称： 公司(企业)网站管理系统（简称：www.web300.cn）
'版权所有: www.web300.cn 
'程序制作：www.web300.cn开发团队
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>

<!-- #include file="Inc/Head.asp" -->
<script language="javascript">
function confirmdel(id,page){
if (confirm("真的要删除这个订单?"))
window.location.href="Bs_Eshop_Del.asp?id="+id+"&page="+page
}
</script>
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="570" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>订单管理</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
		  <%
flag="尚未处理"
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_OrderList  order by RegTime desc"
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=20

'假如没有数据时
'If rs.eof and rs.bof then
'call showpages
'response.write "<p align='center'><font color='#ff0000'>还没任何用户订单</font></p>"
'response.end
'End if

'取得页数,并判断用户输入的是否数字类型的数据，如不是将以第一页显示
dim text,checkpage
text="0123456789"
Rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if

call showpages
call list

If Rs.recordcount > MaxPerPage then
call showpages
end if

'显示帖子的子程序
Sub list()%>
		  <table width="100%" border="0" cellspacing="2" cellpadding="3">
              <tr> 
                <td width="18%" height="25" bgcolor="#C0C0C0"> <div align="center">订单编号</div></td>
                <td width="20%" bgcolor="#C0C0C0"> <div align="center">客户姓名</div></td>
                <td width="22%" bgcolor="#C0C0C0"> <div align="center">联系电话</div></td>
                <td width="15%" bgcolor="#C0C0C0"> <div align="center">处理情况</div></td>
                <td width="15%" bgcolor="#C0C0C0"> <div align="center">订单详情</div></td>
                <td width="10%" bgcolor="#C0C0C0"> <div align="center">操作</div></td>
              </tr>
              <%
if not rs.eof then
i=0
do while not rs.eof
%>
              <tr> 
                <td bgcolor="#E3E3E3"> <div align="center"><%=rs("Form_Id")%></div></td>
                <td bgcolor="#E3E3E3"> <div align="center"><%=rs("UserName")%></div></td>
                <td bgcolor="#E3E3E3"> <div align="center"><%=rs("Phone")%></div></td>
                <td bgcolor="#E3E3E3"> <div align="center"> 
                    <%If rs("Flag")="尚未处理" Then%>
                    尚未处理 
                    <%else%>
                    <b><font color="#FF0000">已经发货</font></b> 
                    <%End If%>
                  </div></td>
                <td bgcolor="#E3E3E3"> <div align="center"> 
                    <%response.write "<a href='Bs_Eshop_detail.asp?ID="&rs("Form_Id")&"&page="&CurrentPage&"' >详细资料</a>"
%>
                  </div></td>
                <td bgcolor="#E3E3E3"> <div align="center"> 
                    <%response.write "<a href='javascript:confirmdel(" & rs("Form_Id") & ","& CurrentPage&")'>删除</a>"
%>
                  </div></td>
              </tr>
			  <%
i=i+1
if i >= MaxPerpage then exit do
rs.movenext
loop
end if
%>
              <tr bgcolor="#C0C0C0"> 
                <td height="25" colspan="6">&nbsp;&nbsp;
                  <%
Response.write "<strong><font color='#000000'>-> 全部-</font>"
Response.write "共</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>件商品</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>第</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='Bs_Eshop.asp?&page="+cstr(1)+"'><font color='#000000'>首页</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='Bs_Eshop.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'>上一页</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>上一页 </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='Bs_Eshop.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'>下一页</font></a><font color='#ffffff'> </font>"
Response.write "<a href='Bs_Eshop.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'>尾页</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>下一页</font></strong>&nbsp;&nbsp;"
End if
'response.write "</td><td align='right'>"
'response.write "<font color='#000000'>转到：</font><input type='text'name='page'size=4 maxlength=4 class=smallInput value="&Currentpage&">&nbsp;"
'response.write "<input class=buttonface type='submit' value='Go' name='cndok'>&nbsp;&nbsp;"
%>
                </td>
              </tr>
              
            </table>
				 <%
End sub
rs.close
conn.close
%> 
				 </td>
        </tr>
      </table>
      
    </td>
	</tr>
</table>
<BR>
<%htmlend%>
