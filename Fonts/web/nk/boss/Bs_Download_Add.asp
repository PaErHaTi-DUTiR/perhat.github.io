<!--#include file="bsconfig.asp"-->

<%
'=========================================================
'
'��Ʒ���ƣ��Ƽ� ��˾(��ҵ)��վ����ϵͳ����ƣ�www.web300.cn��
'��Ȩ���У�www.web300.cn
'����������QQ812256@hotmail.com
'��ϵ ��ʽ��QQ ��812256
'Copyright 2006-2008 www.web300.cn - All Rights Reserved. 
'
'========================================================
%>

<%
dim rs
dim sql
dim count
dim Bs_BigClassName,Bs_SmallClassName

set rs=server.createobject("adodb.recordset")
sql = "select * from Bs_DownSmallClass order by Bs_SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("Bs_SmallClassName"))%>","<%= trim(rs("Bs_BigClassName"))%>","<%= trim(rs("Bs_SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.Bs_SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.Bs_SmallClassName.options[document.myform.Bs_SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    

function CheckForm()
{
  if (document.myform.Bs_Title.value=="")
  {
    alert("���ر��ⲻ��Ϊ�գ�");
	document.myform.Bs_Title.focus();
	return false;
  }
	if (document.myform.Bs_FileSize.value == "")
	{
		alert("�����С����Ϊ�գ�");
	document.myform.Bs_FileSize.focus();
	return false;
	}
  if (document.myform.Bs_DownloadUrl.value=="")
  {
    alert("�������Ӳ���Ϊ�գ�");
	document.myform.Bs_DownloadUrl.focus();
	return false;
  }
  if (document.myform.Bs_Content.value=="")
  {
    alert("�����鲻��Ϊ�գ�");
	document.myform.Bs_Content.focus();
	return false;
  }
  return true;  
}

</script>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="610" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>�� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<form method="POST" name="myform" onSubmit="return CheckForm();" action="Bs_Download_Save.asp?action=add" target="_self">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
				<tr align="center"> 
					<td class="tdbg">
						<table width="100%" border="0" cellpadding="0" cellspacing="2">
							<tr> 
								<td width="100" height="22" align="right" bgcolor="#C0C0C0">�������</td>
								<td width="500" bgcolor="#E3E3E3"> 
<%
		sql = "select * from Bs_DownBigClass"
		rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "���������Ŀ��"
else
%>

<select name="Bs_BigClassName" onChange="changelocation(document.myform.Bs_BigClassName.options[document.myform.Bs_BigClassName.selectedIndex].value)" size="1">
<option selected value="<%=trim(rs("Bs_BigClassName"))%>"><%=trim(rs("Bs_BigClassName"))%></option>
<%
dim selclass
		selclass=rs("Bs_BigClassName")
		rs.movenext
	do while not rs.eof
%>
<option value="<%=trim(rs("Bs_BigClassName"))%>"><%=trim(rs("Bs_BigClassName"))%></option>
<%
			rs.movenext
		loop
end if
	rs.close
%>
</select>

<select name="Bs_SmallClassName">
<option value="" selected>��ָ��С��</option>
<%
sql="select * from Bs_DownSmallClass where Bs_BigClassName='" & selclass & "'"
if yzcv<>zcv then
response.end
End if
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
%>
<option value="<%=rs("Bs_SmallClassName")%>"><%=rs("Bs_SmallClassName")%></option>
<%
rs.movenext
do while not rs.eof
%>
<option value="<%=rs("Bs_SmallClassName")%>"><%=rs("Bs_SmallClassName")%></option>
<%
rs.movenext
loop
end if
rs.close
%>
</select>
								</td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">�������ƣ�</td>
								<td bgcolor="#E3E3E3">
								<input type="text" name="Bs_Title" id="Bs_Title" size="50" maxlength="80">
								<font color="#FF0000">��</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">�ļ���С��</td>
								<td bgcolor="#E3E3E3"> <INPUT maxLength=10 size=14 name=Bs_FileSize> <INPUT type=radio CHECKED value=KB name=Bs_FileSize> KB <INPUT type=radio value=MB name=Bs_FileSize> MB <FONT color=#ff0000>��</FONT> </td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">������ԣ�</td>
								<td bgcolor="#E3E3E3"> <INPUT type=radio CHECKED value=�������� name=Bs_Language>�������� <INPUT type=radio value=�������� name=Bs_Language>�������� <INPUT type=radio value=Ӣ�� name=Bs_Language>Ӣ�� <INPUT type=radio value=������� name=Bs_Language>������� <FONT color=#ff0000>��</FONT></td>
							</tr>
							<tr> 
								<td height="50" align="right" bgcolor="#C0C0C0">��Ȩ���ͣ�</td>
								<td bgcolor="#E3E3E3"> <INPUT type=radio CHECKED value=������� name=Bs_LicenceType>������� <INPUT type=radio value=������ name=Bs_LicenceType>������ <INPUT type=radio value=������� name=Bs_LicenceType>������� <INPUT type=radio value=������� name=Bs_LicenceType>������� <INPUT type=radio value=��ʾ��� name=Bs_LicenceType>��ʾ��� <BR><INPUT type=radio value=��ҵ��� name=Bs_LicenceType>��ҵ��� <INPUT type=radio value=�ƽ���� name=Bs_LicenceType>�ƽ���� <INPUT type=radio value=ע����� name=Bs_LicenceType>ע����� <FONT color=#ff0000>��</FONT> </td>
							</tr>
							<tr> 
								<td height="50" align="right" bgcolor="#C0C0C0">���л�����</td>
								<td bgcolor="#E3E3E3"> <INPUT type=checkbox value=Win2003 name=Bs_System>Win2003 <INPUT type=checkbox CHECKED value=WinXP name=Bs_System>WinXP <INPUT type=checkbox CHECKED value=Win2000 name=Bs_System>Win2000 <INPUT type=checkbox CHECKED value=NT name=Bs_System>NT <INPUT type=checkbox CHECKED value=WinME name=Bs_System>WinME <INPUT type=checkbox CHECKED value=Win9X name=Bs_System>Win9X <BR><INPUT type=checkbox value=Dos name=Bs_System>Dos <INPUT type=checkbox value=Linux name=Bs_System>Linux <INPUT type=checkbox value=Unix name=Bs_System>Unix <INPUT type=checkbox value=Mac name=Bs_System>Mac <INPUT type=checkbox value=ASP���� name=Bs_System>ASP���� <INPUT type=checkbox value=PHP���� name=Bs_System>PHP���� <INPUT type=checkbox value=CGI���� name=Bs_System>CGI���� <FONT color=#ff0000>��</FONT> </td>
							</tr>
							<tr> 
								<td width="100" height="25" align="right" bgcolor="#C0C0C0">ͼƬ��ַ��</td>
								<td width="500" bgcolor="#E3E3E3">
								<input type="text" name="Bs_PhotoUrl" id="Bs_PhotoUrl" value="../img/nopic.jpg" size="50" maxlength="120"> 
								</td>
							</tr>
							<tr>
								<td width="100" align="right" bgcolor="#C0C0C0">ͼƬ�ϴ���</td>
								<td width="500" height="30" bgcolor="#E3E3E3"><div align="center"
><iframe name="ad" frameborder=0 width=100% height=30 scrolling=no src="../WebEdit/Upload_SoftPic.asp"></iframe></div></td>
							</tr>
							<tr> 
								<td height="50" align="right" valign="middle" bgcolor="#C0C0C0">UBB���룺</td>
								<td bgcolor="#E3E3E3">
								<script src=Inc/eshopcode.js></script>
								<!--#include file=Inc/Ubb.inc -->
								</td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td height="25" align="right" bgcolor="#C0C0C0">�����飺</td>
								<td align="right" valign="middle"> <div align="left"> 
								<textarea name="Bs_Content" cols="66" rows="12" wrap="PHYSICAL"></textarea>
								</div></td>
							</tr>
							<tr> 
								<td width="100" align="right" bgcolor="#C0C0C0">���ص�ַһ</td>
								<td width="500" height="25" bgcolor="#E3E3E3">
								<input name="Bs_DownloadUrl" type="text" id="Bs_DownloadUrl" size="50" maxlength="120"> 
								<font color="#FF0000">��</font></td>
							</tr>
							<tr> 
								<td width="100" align="right" bgcolor="#C0C0C0">���ص�ַ��</td>
								<td width="500" height="25" bgcolor="#E3E3E3">
								<input type="text" name="Bs_DownloadUrl2" id="Bs_DownloadUrl2" size="50" maxlength="120"> 
								</td>
							</tr>
							<tr>
								<td width="100" align="right" bgcolor="#C0C0C0">�����ϴ���</td>
								<td width="500" height="30" bgcolor="#E3E3E3"><div align="center"
><iframe name="ad" frameborder=0 width=100% height=30 scrolling=no src="../WebEdit/Upload_Soft.asp"></iframe></div></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">��ͨ����ˣ�</td>
								<td bgcolor="#E3E3E3">
								<input type="checkbox" name="Bs_Passed" id="Bs_Passed" value="yes" checked>
								<font color="#0000FF">�����ѡ�еĻ���ֱ�ӷ�����</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">�����Ƽ���</td>
								<td bgcolor="#E3E3E3">
								<input type="checkbox" name="Bs_Elite" id="Bs_Passed" value="yes">
								<font color="#0000FF">���Ƿ���Ϊ�Ƽ���</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">����ʱ�䣺</td>
								<td bgcolor="#E3E3E3">
								<input type="text" name="Bs_UpDateTime" id="Bs_UpDateTime" value="<%=now()%>" maxlength="50">��ǰʱ��Ϊ��<%=now()%> ע�ⲻҪ�ı��ʽ��</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
				<div align="center"><p> 
				<input	name="Add" type="submit"  id="Add" value=" �� �� " 
				onClick="document.myform.action='Bs_Download_Save.asp?action=add';document.myform.target='_self';" style="cursor: hand;background-color: #cccccc;">&nbsp;
				<input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='Bs_Download_ADD.asp'" style="cursor: hand;background-color: #cccccc;">
				</p></div>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
