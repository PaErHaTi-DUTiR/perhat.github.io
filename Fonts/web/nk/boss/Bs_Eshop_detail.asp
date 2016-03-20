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
<%
id=request("id")
page=request("page")
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_OrderList where Form_Id=" & id
rs.open sqltext,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="570" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>订单管理</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<table width="562" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="24"> <BR>
            <div align="center"><strong><%=rs("Form_Id")%>号订单管理<br>
              <br>
              </strong></div></td>
        </tr>
        <tr> 
            
          <form method='POST'action="OrderList_Save.asp?Form_Id=<%=rs("Form_Id")%>"><td> 
            <TABLE cellSpacing=1 cellPadding=4 width=562 bgColor=#006699 height="436">
                
                  <TR> 
                    <TD width="548" height="10"  colSpan=2 valign="top" bgcolor="#DBDBDB"></TD>
                  </TR>
                  <TR bgColor=#eeeeee> 
                    <TD width="548" height="32"  colSpan=2><font color="#000000">客户订货单详细资料</font></TD>
                  </TR>
                  <TR> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">订货单号：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Form_Id")%></TD>
                  </TR>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right">公司名称<font color="#000000">：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Comane")%></TD>
                  </TR>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">收货人姓名：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Somane")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">收货地址：</font></TD>
                    <TD  width=410 height=-2 bgcolor="#eeeeee">&nbsp; <%=rs("Add")%></TD>
                  </tr>
                  <tr> 
                    <TD bgColor=#DBDBDB height=25 align="right"><font color="#000000">邮政编码：</font></TD>
                    <TD  width=410 height=0 bgcolor="#eeeeee"> &nbsp; <%=rs("Zip")%></TD>
                  </tr>
                  <tr> 
                    <TD bgColor=#DBDBDB height=25 align="right"><font color="#000000">联系电话：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Phone")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right">联系传真<font color="#000000">：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Fox")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">电子信箱：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Email")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">所选汇款账号：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Pays")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB><font color="#000000">备注：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee">&nbsp; <%=rs("Remark")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=24 align="right"><font color="#000000">订货日期：</font></TD>
                    <TD  width=410 height=24 bgcolor="#eeeeee">&nbsp; <%=rs("RegTime")%></TD>
                  </tr>
                  <tr> 
                    <TD  width=126 bgColor=#DBDBDB height=25 align="right"><font color="#000000">订单是否已经处理：</font></TD>
                    <TD  width=410 height=25 bgcolor="#DCDDDE"> 
                      <%If rs("Flag")="尚未处理" Then%>
                      尚未处理 
                      <%else%>
                      已经发货 
                      <%End If%>                    </TD>
                  </tr>
                  <TR> 
                    <TD width="548" height="31"  colSpan=2 valign="top" bgcolor="#eeeeee"> 
                      <p align="center">订货商品细目</p></TD>
                  </TR>
                  <%
set rs2=server.createobject("adodb.recordset")
sqltext2="select * from Bs_ShopList where Form_Id=" & id
rs2.open sqltext2,conn,1,1
%>
                  <TR> 
                    <TD width="548" height="15"  colSpan=2 valign="top" bgcolor="#eeeeee"> 
                      <div align="center"> 
                        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#006699" bordercolordark="#eeeeee"  height="67">
                          <tr> 
                            <td width="17%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">产品编号</font></td>
                            <td width="41%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">产品名称</font></td>
                            <td width="28%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">产品类别</font></td>
                            <td width="14%" height="21" align="center" bgcolor="#DBDBDB"><font color="#000000">产品数量</font></td>
                          </tr>
                          <%Sum=0
While Not rs2.EOF%>
                          <tr> 
                            <td width="17%" align="center" height="22"><%=rs2("Product_Id")%></td>
                            <td width="41%" align="center" height="22"><%=rs2("Product_Name")%></td>
                            <td width="28%" align="center" height="22"><%=rs2("BigClassName")%> 
                              => <%=rs2("SmallClassName")%></td>
                            <td height="22" align="center"><%=rs2("Number")%></td>
                            <%Sum=Sum+rs2("P_NewPrice")*rs2("Number")%>
                          </tr>
                          <%
rs2.MoveNext
Wend
%>
                          <tr> 
                            <td colspan="4" height="22"> <p align="right">&nbsp;</p></td>
                          </tr>
                        </table>
                      </div></TD>
                  </TR>
                <center>
                  <TR> 
                    <TD width="548" height="25"  colSpan=2 bgcolor="#DCDDDE"> 
                      <p align="center"> 
                        <%If rs("Flag")="尚未处理" Then%>
                        <input type="submit" value="订单处理" name="B1">
                        <%
rs.close
rs2.close
conn.close
End If
%>
                        <input type="button" value="返回" name="B4" onClick="javascript:window.history.go(-1)">
                    </TD>
                  </TR>
                  <TR bgColor=#DBDBDB> 
                    <TD width="548" height="3"  colSpan=2></TD>
                  </TR>
                </center>
              </TABLE>
          </td></form>
        </tr>
      </table><BR>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
