<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
'=========================================================
'
'产品名称：良精科技 公司(企业)网站管理系统（简称：liangjing.net）
'版权所有: liangjing.net 
'程序制作：liangjing.net开发团队
'Copyright 2004-2005 liangjing.net - All Rights Reserved. 
'
'========================================================
%>
<%
dim ArticleID,rsArticle,FoundErr,ErrMsg

ArticleID=trim(request("ArticleID"))
FoundErr=False
if ArticleID="" then 
	response.Redirect("Bs_Product.asp")
end if
sql="select * from Bs_Product where ArticleID=" & ArticleID & ""
Set rsArticle= Server.CreateObject("ADODB.Recordset")
rsArticle.open sql,conn,1,1
if FoundErr=True then
	call WriteErrMsg()
else
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
    alert("文章标题不能为空！");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Key.value=="")
  {
    alert("关键字不能为空！");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("文章内容不能为空！");
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
		<td class="a1" height="25" align="center"><strong>修 改 产 品</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
			<form method="POST" name="myform" onSubmit="return CheckForm();" action="Bs_Product_Save.asp?action=Modify">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
				<tr align="center">
					<td class="tdbg">
						<table width="100%" border="0" cellpadding="0" cellspacing="2">
							<tr> 
								<td width="100" height="22" align="right" bgcolor="#C0C0C0">所属栏目：</td>
								<td width="500" bgcolor="#E3E3E3"> 
<%
if session("purview")=3 or session("purview")=4 then
	response.write rsArticle("BigClassName") & "<input name='BigClassName'type='hidden'value='" & rsArticle("BigClassName") & "'>&gt;&gt;"
else		
			sql = "select * from Bs_PrBigClasss"
			rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "请先添加栏目。"
	else
%>
<select name="BigClassName" 
onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
<%
do while not rs.eof
%>
<option <% if rs("BigClassName")=rsArticle("BigClassName") then response.Write("selected") end if%> 
value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
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
	response.write rsArticle("SmallClassName") & "<input name='SmallClassName'type='hidden'value='" & rsArticle("SmallClassName") & "'>"
else
%>
<select name="SmallClassName">
<option value="" <%if rsArticle("SmallClassName")="" then response.write "selected"%>>不指定小类</option>
<%
sql="select * from Bs_PrSmallClass where BigClassName='" & rsArticle("BigClassName") & "'"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
do while not rs.eof
%>
<option <% if rs("SmallClassName")=rsArticle("SmallClassName") then response.Write("selected") end if%> 
value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
<%
rs.movenext
loop
end if
rs.close
%>
</select> 
<%
end if
%>								</td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">产品编号：</td>
								<td bgcolor="#E3E3E3">
								<input name="Product_Id" type="text" id="Product_Id" value="<%=rsArticle("Product_Id")%>" size="15" maxlength="15"> 
								<font color="#FF0000">*</font> </td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">产品名称：</td>
								<td bgcolor="#E3E3E3">
								<input name="Title" type="text" id="Title" value="<%=rsArticle("Title")%>" size="50" maxlength="80"> 
								<font color="#FF0000">*</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">English名称：</td>
								<td bgcolor="#E3E3E3">
								<input name="EnTitle" type="text" id="EnTitle" value="<%=rsArticle("EnTitle")%>" size="50" maxlength="80"> 
								<font color="#FF0000">*</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">产品规格：</td>
								<td bgcolor="#E3E3E3">
								<input name="Spec" type="text" id="Spec" value="<%=rsArticle("Spec")%>" size="50" maxlength="80"></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">English规格：</td>
								<td bgcolor="#E3E3E3">
								<input name="EnSpec" type="text" id="EnSpec" value="<%=rsArticle("EnSpec")%>" size="50" maxlength="80"></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">条形码及编号：</td>
								<td bgcolor="#E3E3E3">
								<input name="Size" type="text" id="Size" value="<%=rsArticle("Size")%>" size="50" maxlength="80">								</td>
							</tr>
<!-- 							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">产品备注：</td>
								<td bgcolor="#E3E3E3">
								<input name="Memo" type="text" id="Memo" value="<%=rsArticle("Memo")%>" size="50" maxlength="80"> 
							</td>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">English备注：</td>
								<td bgcolor="#E3E3E3">
								<input name="EnMemo" type="text" id="EnMemo" value="<%=rsArticle("EnMemo")%>" size="50" maxlength="80"> 
								</td>
							</tr> -->
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">关键字：</td>
								<td bgcolor="#E3E3E3">
								<input name="Key" type="text" id="Key" value="<%=trim(rsArticle("Key"))%>" size="50" maxlength="255"> 
								<font color="#FF0000">*</font><br>用来查找相关文章，可输入多个关键字，中间用空格分开。不能出现&quot;'*?()等字符。</td>
							</tr>
							<tr> 
								<td height="22" align="right" valign="middle" bgcolor="#C0C0C0">产品说明：</td>
								<td bgcolor="#E3E3E3"> </td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td colspan="2" align="right" valign="middle"><div align="left"> 
								<textarea name="Content" style="display:none"><%=rsArticle("Content")%></textarea>
								<iframe ID="editor" src="../WebEdit/editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe>
								</div></td>
							</tr>
							<tr> 
								<td height="22" align="right" valign="middle" bgcolor="#C0C0C0">English说明：</td>
								<td bgcolor="#E3E3E3"><span style="color: #FF6600">小提示：英文版式如果要与中文版式使用同一幅产品图片，可直接复制上面中文版的图片信息即可。无须再重新上传。</span> </td>
							</tr>
							<tr bgcolor="#E3E3E3"> 
								<td colspan="2" align="right" valign="middle"><div align="left"> 
								<textarea name="EnContent" style="display:none"><%=rsArticle("EnContent")%></textarea>
								<iframe ID="Eneditor" src="../WebEdit/editor.asp" frameborder=1 scrolling=no width="600" height="405"></iframe>
								</div></td>
							</tr>
							<tr>
								<td width="100" align="right" bgcolor="#C0C0C0">首页图片： 
								<input name="IncludePic" type="hidden" id="IncludePic" value="yes"></td>
								<td width="500" bgcolor="#E3E3E3">
								<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="<%=rsArticle("DefaultPicUrl")%>" size="50" maxlength="180">
								<br>首页的图片,直接从上传图片中选择： 
<select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
<option value=""<% if rsArticle("DefaultPicUrl")="" then response.write "selected" %>>不指定首页图片</option>
<%
if rsArticle("UploadFiles")<>"" then
	dim IsOtherUrl
	IsOtherUrl=True
	if instr(rsArticle("UploadFiles"),"|")>1 then
		dim arrUploadFiles,intTemp
		arrUploadFiles=split(rsArticle("UploadFiles"),"|")						
		for intTemp=0 to ubound(arrUploadFiles)
			if rsArticle("DefaultPicUrl")=arrUploadFiles(intTemp) then
				response.write "<option value='" & arrUploadFiles(intTemp) & "'selected>" & arrUploadFiles(intTemp) & "</option>"
				IsOtherUrl=False
			else
				response.write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
			end if
		next
	else
		if rsArticle("UploadFiles")=rsArticle("DefaultPicUrl") then
			response.write "<option value='" & rsArticle("UploadFiles") & "'selected>" & rsArticle("UploadFiles") & "</option>"
			IsOtherUrl=False
		else
			response.write "<option value='" & rsArticle("UploadFiles") & "'>" & rsArticle("UploadFiles") & "</option>"		
		end if
	end If
	if IsOtherUrl=True then
		response.write "<option value='" & rsArticle("DefaultPicUrl") & "'selected>" & rsArticle("DefaultPicUrl") & "</option>"
	end if
end if
%>
</select>
								<input name="UploadFiles" type="hidden" id="UploadFiles" value="<%=rsArticle("UploadFiles")%>">								</td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">已通过审核：</td>
								<td bgcolor="#E3E3E3">
								<input name="Passed" type="checkbox" id="Passed" value="yes" <% if rsArticle("Passed")=true then response.Write("checked") end if%>>
								是<font color="#0000FF">（如果选中的话将直接发布）</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">首页显示：</td>
								<td bgcolor="#E3E3E3">
								<input name="Elite" type="checkbox" id="Elite" value="yes" <% if rsArticle("Elite")=true then response.Write("checked") end if%>>
								是<font color="#0000FF">（如果选中的话将在中文首页做为新产品显示）</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">En首页显示：</td>
								<td bgcolor="#E3E3E3">
								<input name="EnElite" type="checkbox" id="EnElite" value="yes" <% if rsArticle("EnElite")=true then response.Write("checked") end if%>>
								是<font color="#0000FF">（如果选中的话将在英文首页做为新产品显示）</font></td>
							</tr>
							<tr> 
								<td height="22" align="right" bgcolor="#C0C0C0">录入时间：</td>
								<td bgcolor="#E3E3E3">
								<input name="UpdateTime" type="text" id="UpdateTime" value="<%=rsArticle("UpdateTime")%>" maxlength="50">
								当前时间为：<%=now()%> 注意不要改变格式。</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
				<div align="center"><p>
				<input name="ArticleID" type="hidden" id="ArticleID" value="<%=rsArticle("ArticleID")%>">
				<input name="Save" type="submit"  id="Save" value="保存修改结果">
				</p></div>
			</form>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
end if
rsArticle.close
set rsArticle=nothing
call CloseConn()
%>
