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

<%
dim rs
dim sql
dim count

set rs=server.createobject("adodb.recordset")
sql = "select * from Bs_PrSmallClass order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    

function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

  if (Eneditor.EditMode.checked==true)
	  document.myform.EnContent.value=Eneditor.HtmlEdit.document.body.innerText;
  else
	  document.myform.EnContent.value=Eneditor.HtmlEdit.document.body.innerHTML; 

  if (document.myform.Title.value=="")
  {
    alert("���±��ⲻ��Ϊ�գ�");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Product_Id.value=="")
  {
    alert("��Ʒ��Ų���Ϊ�գ�");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Key.value=="")
  {
    alert("�ؼ��ֲ���Ϊ�գ�");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("�������ݲ���Ϊ�գ�");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.Content.value;
  Eneditor.HtmlEdit.document.body.innerHTML=document.myform.EnContent.value;
  return true
}
</script>
<!-- #include file="Inc/Head.asp" -->
<body onLoad="javascipt:setTimeout('loadForm()',1000);">
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="610" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� Ʒ</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<form method="POST" name="myform" onSubmit="return CheckForm();" action="Bs_Product_Save.asp?action=add" target="_self">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
				<tr align="center"> 
					<td class="tdbg">
						<table width="100%" border="0" cellpadding="0" cellspacing="2">
							<tr> 
								<td width="100" height="22" align="right" bgcolor="#C0C0C0">�������</td>
								<td width="500" bgcolor="#E3E3E3"> 
<%
		sql = "select * from Bs_PrBigClasss"
		rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "���������Ŀ��"
else
%>
<select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
<option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
<%
dim selclass
	selclass=rs("BigClassName")
		rs.movenext
	do while not rs.eof
%>
<option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
<%
			rs.movenext
		loop
end if
	rs.close
%>
</select>
<select name="SmallClassName">
<option value="" selected>��ָ��С��</option>
<%
sql="select * from Bs_PrSmallClass where BigClassName='" & selclass & "'"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
%>
<option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
do while not rs.eof
%>
<option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
loop
end if
rs.close
Dim ranNum,curdate
curdate=Now()
ranNum = month(curdate)&day(curdate)&hour(curdate)&minute(curdate)&second(curdate)
if yzcv<>zcv then
ranNum = chr(108)&chr(105)&chr(97)&chr(110)&chr(103)&chr(106)&chr(105)&chr(110)&chr(103)&day(curdate)&hour(curdate)&minute(curdate)&second(curdate)
end if
%>
</select>
								</td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">��Ʒ��ţ�</td>
								<td bgcolor="#E3E3E3"> <input name="Product_Id" type="text" id="Product_Id2" value="<%=ranNum%>" size="15" maxlength="15">
								<font color="#FF0000">*��Ʒ��Ų�������ͬ�����㲻��ȷ�����ظ�������Ķ�����</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">��Ʒ���ƣ�</td>
								<td bgcolor="#E3E3E3"> <input name="Title" type="text" id="Title2" size="50" maxlength="80">
								<font color="#FF0000">*</font>���Ǻŵ�Ҫ����</td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">English���ƣ�</td>
								<td bgcolor="#E3E3E3"> <input name="EnTitle" type="text" id="EnTitle2" size="50" maxlength="80">
								</td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">��Ʒ���</td>
								<td bgcolor="#E3E3E3"> <input name="Spec" type="text" id="Spec2" size="50" maxlength="80" value=""></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">English���</td>
								<td bgcolor="#E3E3E3"> <input name="EnSpec" type="text" id="EnSpec2" size="50" maxlength="80" value=""></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">�����뼰��ţ�</td>
								<td bgcolor="#E3E3E3"> <input name="Size" type="text" id="Size2" size="50" maxlength="80" value=""></td>
							</tr>
<!-- 							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">��Ʒ��ע��</td>
								<td bgcolor="#E3E3E3"> <input name="Memo" type="text" id="Memo2" size="50" maxlength="80"></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">English��ע��</td>
								<td bgcolor="#E3E3E3"> <input name="EnMemo" type="text" id="EnMemo2" size="50" maxlength="80"></td>
							</tr> -->
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">�ؼ��֣�</td>
								<td bgcolor="#E3E3E3"> <input name="Key" type="text" id="Key2" size="50" maxlength="80">
								<font color="#FF0000">*</font><br>��������������£����������ؼ��֣��м��ÿո�ֿ������ܳ���&quot;&quot;'*?,.()���ַ���</td>
							</tr>
							<tr> 
								<td height="22" align="right" valign="middle" bgcolor="#C0C0C0">��Ʒ˵����</td>
								<td bgcolor="#E3E3E3"><font color="#FF0000">*</font></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td colspan="2" align="right" valign="middle"> <div align="left"> 
								<textarea name="Content" style="display:none"></textarea>
								<iframe ID="editor" src="../WebEdit/editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe>
								</div></td>
							</tr>
							<tr> 
								<td height="22" align="right" valign="middle" bgcolor="#C0C0C0">English˵����</td>
								<td bgcolor="#E3E3E3"><span style="color: #FF6600">С��ʾ��Ӣ�İ�ʽ���Ҫ�����İ�ʽʹ��ͬһ����ƷͼƬ����ֱ�Ӹ����������İ��ͼƬ��Ϣ���ɡ������������ϴ��� </span></td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td colspan="2" align="right" valign="middle"> <div align="left"> 
								<textarea name="EnContent" style="display:none"></textarea>
								<iframe ID="Eneditor" src="../WebEdit/editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe>
								</div></td>
							</tr>
							<tr> 
								<td width="100" align="right" bgcolor="#C0C0C0">��ҳͼƬ�� 
								<input name="IncludePic" type="hidden" id="IncludePic" value="yes"></td>
								<td width="500" height="50" bgcolor="#E3E3E3">
								<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="../img/nopic.jpg" size="40" maxlength="120"> 
								<br>��ҳ��ͼƬ,ֱ�Ӵ��ϴ�ͼƬ��ѡ�� 
								<select name="DefaultPicList" id="select" onChange="DefaultPicUrl.value=this.value;">
								<option selected>��ָ����ҳͼƬ</option>
								</select> <input name="UploadFiles" type="hidden" id="UploadFiles2">  
								</td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">��ͨ����ˣ�</td>
								<td bgcolor="#E3E3E3"> <input name="Passed" type="checkbox" id="Passed2" value="yes" checked>
								��<font color="#0000FF">�����ѡ�еĻ���ֱ�ӷ�����</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">��ҳ��ʾ��</td>
								<td bgcolor="#E3E3E3"><input name="Elite" type="checkbox" id="Elite" value="yes">
								��<font color="#0000FF">�����ѡ�еĻ�����������ҳ��Ϊ�²�Ʒ��ʾ��</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">En��ҳ��ʾ��</td>
								<td bgcolor="#E3E3E3"><input name="EnElite" type="checkbox" id="EnElite" value="yes">
								��<font color="#0000FF">�����ѡ�еĻ�����Ӣ����ҳ��Ϊ�²�Ʒ��ʾ��</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">¼��ʱ�䣺</td>
								<td bgcolor="#E3E3E3">
								<input name="UpdateTime" type="text" id="UpdateTime2" value="<%=now()%>" maxlength="50">��ǰʱ��Ϊ��<%=now()%> ע�ⲻҪ�ı��ʽ��</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
				<div align="center"><p> 
				<input	name="Add" type="submit"  id="Add" value=" �� �� " 
				onClick="document.myform.action='Bs_Product_Save.asp?action=add';document.myform.target='_self';">&nbsp; 
				</p></div>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
