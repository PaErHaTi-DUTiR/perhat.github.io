<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="Inc/Function.asp"-->
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
dim Bs_DownID,rsDown,FoundErr,ErrMsg

Bs_DownID=trim(request("Bs_DownID"))
FoundErr=False
if Bs_DownID="" then 
	response.Redirect("Bs_Download.asp")
end if
sql="select * from Bs_Download where Bs_DownID=" & Bs_DownID & ""
Set rsDown= Server.CreateObject("ADODB.Recordset")
rsDown.open sql,conn,1,1
if FoundErr=True then
	call WriteErrMsg()
else
%>
<%
dim rs
dim sql
dim count
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
	if (document.myform.Title.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
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
    alert("�������ݲ���Ϊ�գ�");
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
			<form method="POST" name="myform" onSubmit="return CheckForm();" action="Bs_Download_Save.asp?action=Modify">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
				<tr align="center">
					<td class="tdbg">
						<table width="100%" border="0" cellpadding="0" cellspacing="2">
							<tr> 
								<td width="100" height="22" align="right" bgcolor="#C0C0C0">������Ŀ��</td>
								<td width="500" bgcolor="#E3E3E3"> 
<%
if session("purview")=3 or session("purview")=4 then
	response.write rsDown("Bs_BigClassName") & "<input name='Bs_BigClassName'type='hidden'value='" & rsDown("Bs_BigClassName") & "'>&gt;&gt;"
else		
			sql = "select * from Bs_DownBigClass"
			rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "���������Ŀ��"
	else
%>
<select name="Bs_BigClassName" 
onChange="changelocation(document.myform.Bs_BigClassName.options[document.myform.Bs_BigClassName.selectedIndex].value)" size="1">
<%
do while not rs.eof
%>
<option <% if rs("Bs_BigClassName")=rsDown("Bs_BigClassName") then response.Write("selected") end if%> 
value="<%=trim(rs("Bs_BigClassName"))%>"><%=trim(rs("Bs_BigClassName"))%></option>
<%
rs.movenext
loop
end if
rs.close
%>
</select> 
<%
end if
if session("purview")=4 then
	response.write rsDown("Bs_SmallClassName") & "<input name='Bs_SmallClassName'type='hidden'value='" & rsDown("Bs_SmallClassName") & "'>"
else
%>
<select name="Bs_SmallClassName">
<option value="" <%if rsDown("Bs_SmallClassName")="" then response.write "selected"%>>��ָ��С��</option>
<%
sql="select * from Bs_DownSmallClass where Bs_BigClassName='" & rsDown("Bs_BigClassName") & "'"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
do while not rs.eof
%>
<option <% if rs("Bs_SmallClassName")=rsDown("Bs_SmallClassName") then response.Write("selected") end if%> 
value="<%=rs("Bs_SmallClassName")%>"><%=rs("Bs_SmallClassName")%></option>
<%
rs.movenext
loop
end if
rs.close
%>
</select> 
<%
end if
%>
								</td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">�������ƣ�</td>
								<td bgcolor="#E3E3E3"> <input type="text" name="Bs_Title" id="Bs_Title" value="<%=trim(rsDown("Bs_Title"))%>" size="50" maxlength="80">
								<font color="#FF0000">��</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">�ļ���С��</td>
								<td bgcolor="#E3E3E3"> <input type="text" name="Bs_FileSize" id="Bs_FileSize" value="<%=trim(rsDown("Bs_FileSize"))%>" size="50" maxlength="80">
								<FONT color=#ff0000>��</FONT> </td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">������ԣ�</td>
								<td bgcolor="#E3E3E3"> <input type="text" name="Bs_Language" id="Bs_Language" value="<%=trim(rsDown("Bs_Language"))%>" size="50" maxlength="80">
								<font color="#FF0000">��</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">��Ȩ���ͣ�</td>
								<td bgcolor="#E3E3E3"> <input type="text" name="Bs_LicenceType" id="Bs_LicenceType" value="<%=trim(rsDown("Bs_LicenceType"))%>" size="50" maxlength="80">
								<font color="#FF0000">��</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">���л�����</td>
								<td bgcolor="#E3E3E3"> <input type="text" name="Bs_System" id="Bs_System" value="<%=trim(rsDown("Bs_System"))%>" size="50" maxlength="80">
								<font color="#FF0000">��</font></td>
							</tr>
							<tr> 
								<td width="100" height="25" align="right" bgcolor="#C0C0C0">ͼƬ��ַ��</td>
								<td width="500" bgcolor="#E3E3E3">
								<input type="text" name="Bs_PhotoUrl" id="Bs_PhotoUrl" value="<%=trim(rsDown("Bs_PhotoUrl"))%>" size="50" maxlength="120"> 
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
<textarea name="Bs_Content" cols="66" rows="12">
<%if rsDown("html")=false then
Bs_Content=replace(rsDown("Bs_Content"),"<br>",chr(13))
Bs_Contentt=replace(Bs_Content,"&nbsp;"," ")
else
Bs_Content=rsDown("Bs_Content")
end if
response.write Bs_Content%></textarea>
								</div></td>
							</tr>
							<tr> 
								<td width="100" align="right" bgcolor="#C0C0C0">���ص�ַһ</td>
								<td width="500" height="25" bgcolor="#E3E3E3">
								<input type="text" name="Bs_DownloadUrl" id="Bs_DownloadUrl" value="<%=trim(rsDown("Bs_DownloadUrl"))%>" size="50" maxlength="120"> 
								<font color="#FF0000">��</font></td>
							</tr>
							<tr> 
								<td width="100" align="right" bgcolor="#C0C0C0">���ص�ַ��</td>
								<td width="500" height="25" bgcolor="#E3E3E3">
								<input type="text" name="Bs_DownloadUrl2" id="Bs_DownloadUrl2" value="<%=trim(rsDown("Bs_DownloadUrl2"))%>" size="50" maxlength="120"> 
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
								<input type="checkbox" name="Bs_Passed" id="Bs_Passed" value="yes" <% if rsDown("Bs_Passed")=true then response.Write("checked") end if%>>
								<font color="#0000FF">�����ѡ�еĻ���ֱ�ӷ�����</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">�����Ƽ���</td>
								<td bgcolor="#E3E3E3">
								<input type="checkbox" name="Bs_Elite" id="Bs_Elite" value="yes" <% if rsDown("Bs_Elite")=true then response.Write("checked") end if%>>
								<font color="#0000FF">���Ƿ���Ϊ�Ƽ���</font></td>
							</tr>
							<tr> 
								<td height="25" align="right" bgcolor="#C0C0C0">����ʱ�䣺</td>
								<td bgcolor="#E3E3E3">
								<input type="text" name="Bs_UpDateTime" id="Bs_UpDateTime" value="<%=trim(rsDown("Bs_UpDateTime"))%>" maxlength="50">
								��ǰʱ��Ϊ��<%=now()%> ע�ⲻҪ�ı��ʽ��</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
				<div align="center"><p>
				<input name="Bs_DownID" type="hidden" id="Bs_DownID" value="<%=rsDown("Bs_DownID")%>">
				<input name="Save" type="submit"  id="Save" value="�����޸Ľ��" style="cursor: hand;background-color: #cccccc;">&nbsp; 
				<input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='Bs_Download.asp'"  style="cursor: hand;background-color: #cccccc;">
				</p></div>
			</form>
		</td>
	</tr>
</table>
<BR>
<%
end if
rsDown.close
set rsDown=nothing
call CloseConn()
%>
<%htmlend%>
