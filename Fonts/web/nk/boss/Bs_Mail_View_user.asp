<!--#include file="bsconfig.asp"-->
<%
'=========================================================
'
'��Ʒ���ƣ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ����: www.web300.cn 
'����������www.web300.cn�����Ŷ�
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>
<% if request("del")<>"" then 
   conn.Execute("delete * from email where id="&request("del"))
   end if %>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
   if (isNaN(go2to.page.value))
		alert("����ȷ��дת��ҳ����");
   else if (go2to.page.value=="") 
	     {
		alert("������ת��ҳ����");
		 }
   else
		go2to.submit();
}
//-->
</SCRIPT>
<% 
     set rs=server.createobject("adodb.recordset") 
     if not isempty(request("page")) then   
     pagecount=cint(request("page"))   
     else   
     pagecount=1   
     end if

     if key="" then
     sql="select * from email order by id desc" 
     else
     sql="select * from email where dateandtime like '%"&key&"%'order by id desc" 
     end if

	 rs.open sql,conn,1,1   
     if rs.eof and rs.bof then   
     response.write"<BR>" 
	 response.write"=== �ʼ��б�Ϊ�գ�����ǰ̨��������һ�������ʼ���ַ���ſ��Թ����� ==="
     response.write"<BR><BR>"

	 response.end  
	 end if
     rs.pagesize=20
     if pagecount>rs.pagecount or pagecount<=0 then              
     pagecount=1              
     end if              
     rs.AbsolutePage=pagecount %>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�û�����</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <br>
      <table border="0" cellpadding="0" cellspacing="0" width="550" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" class=ver>
        <tr> 
          <% if rs.pagecount=1 then %>
          <td height="25" colspan="4"  class="ver"> <font class=ver>&nbsp;����[<font color="#ff0000"><%=rs.recordcount%></font>]Email��ַ 
            ������[<font color="red">1��<%=rs.recordcount%></font>]��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="Bs_Mail_Add_user.asp">�����ʼ���ַ</a></font></td>
        </tr>
        <tr> 
          <td height="7" colspan="4" ></td>
        </tr>
        <%else%>
        <tr> 
          <td height="25" colspan="4" ><font class=ver> 
            <% 
            page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
            ����[<font color="#ff0000"><%=rs.recordcount%></font>]Email��ַ ������[<font color="red"><%=page_start%>��<%=page_end%></font>]</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="Bs_Mail_Add_user.asp">�����ʼ���ַ</a></td>
        </tr>
        <tr> 
          <td height="25" colspan="4" > 
            <% response.write"<form name=go2to form method=Post action='Bs_Mail_view_user.asp?key="+key+"'>"
		     response.write "<font color=ffffff>��&nbsp;</font>"                                              
		     if pagecount=1 then                                                        
			 response.write "<font class=ver>��ҳ ��һҳ</font>&nbsp;"
			 else                                                        
             response.write "<a href=Bs_Mail_view_user.asp?page=1&key="+key+"><font class=ver>��ҳ</font></a>&nbsp;" 
	         response.write "<a href=Bs_Mail_view_user.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>��һҳ</font></a>&nbsp;"
			 end if                                             
             if rs.PageCount-pagecount<1 then                                                        
             response.write "<font class=ver>��һҳ βҳ</font>"                                                    
			 else                                                        
             response.write "<a href=Bs_Mail_view_user.asp?page="+cstr(pagecount+1)+"&key="+key+"><font class=ver>��һҳ</font></a>&nbsp;"
			 response.write "<a href=Bs_Mail_view_user.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font class=ver>βҳ</font></a>"           
             end if 
			 response.write "<font class=ver>&nbsp;ҳ��:<font class=ver>"&pagecount&"</font>/"&rs.pagecount&"ҳ</font>" 
			response.write "<font class=ver> ת����<input type='text'name='page'size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18'value="&PageCount&">ҳ</font>&nbsp;&nbsp;&nbsp;&nbsp;"                               
			response.write "<input class=button1 type='button'value='ȷ ��'onclick=check() style='font-family: ����; font-size: 9pt; color: #000073; position: relative; height: 19'>" %>
          </td>
        </tr>
        <%end if%>
      </table>
      <table border="0" cellspacing="2" cellpadding="0" width="550">
        <tr bgcolor="#C0C0C0"> 
          <td width="9%" height="25" class=ver> <p align="center"><font color="#000000">ID</font></p></td>
          <td width="31%"> <p align="center"><font color="#000000">�ʼ���ַ</font></td>
          <td colspan="3"  class=ver> <p align="center"><font color="#000000">�û���</font></td>
          <td  width="16%"> <p align="center"><font color="#000000">- ɾ�� - </font> 
          </td>
        </tr>
        <% do while not rs.eof %>
        <tr bgcolor="#E3E3E3"> 
          <td  height="25" class=ver width="9%"> <p align="center"><font color=000000><%=rs("id")%></font></p></td>
          <td  class=ver> <p align="left"><font color="000000"> &nbsp;&nbsp;&nbsp;&nbsp;<a href="Bs_Mail_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></font></p></td>
          <td colspan="3"  class=ver> <p align="left"><font color="000000"> &nbsp;&nbsp;&nbsp;&nbsp;<%=rs("dateandtime")%></font></p></td>
          <td width="16%" > <p align="center"><font color="#000000">[<a href="Bs_Mail_view_user.asp?del=<%=rs("id")%>"><font class=ver>ɾ��</font></a>]</font> 
            </p></td>
        </tr>
        <% i=i+1                                                                                                  
          rs.movenext                                                                                                  
          if i>=rs.PageSize then exit do 
		  loop                                                                    
          rs.close                                                                                                
          set rs=nothing                                                                                                
          conn.close                                                                                                
          set conn=nothing 
'end if
%>
      </table>
      <br>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>