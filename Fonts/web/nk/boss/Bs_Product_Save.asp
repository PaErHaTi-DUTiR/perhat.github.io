<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

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
dim rs,sql,ErrMsg,FoundErr
dim ArticleID,Product_Id,BigClassName,EnBigClassName,SmallClassName,EnSmallClassName,Title,EnTitle,Spec,EnSpec,Size,Memo,EnMemo,Content,EnContent,key,UpdateTime,Hits
dim IncludePic,DefaultPicUrl,UploadFiles,Elite,EnElite,Passed,arrUploadFiles
dim ObjInstalled

ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=false
ArticleID=Trim(Request.Form("ArticleID"))
Product_Id=trim(request.form("Product_Id"))
BigClassName=trim(request.form("BigClassName"))
SmallClassName=trim(request.form("SmallClassName"))
Title=trim(request.form("Title"))
EnTitle=trim(request.form("EnTitle"))
Spec=trim(request.form("Spec"))
EnSpec=trim(request.form("EnSpec"))
Size=trim(request.form("Size"))
'Memo=trim(request.form("Memo"))
'EnMemo=trim(request.form("EnMemo"))
Key=trim(request.form("Key"))
Content=trim(request.form("Content"))
EnContent=trim(request.form("EnContent"))
UpdateTime=trim(request.form("UpdateTime"))
IncludePic=trim(request.form("IncludePic"))
DefaultPicUrl=trim(request.form("DefaultPicUrl"))
UploadFiles=trim(request.form("UploadFiles"))
Passed=trim(request.form("Passed"))
Elite=trim(request.form("Elite"))
EnElite=trim(request.form("EnElite"))
Hits=trim(request.form("Hits"))

sqlBig="select * from Bs_PrBigClasss where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
EnBigClassName=rsBig("EnBigClassName")
rsBig.close

sqlSmall="select * from Bs_PrSmallClass where SmallClassName='" & SmallClassName & "'and  BigClassName='" & BigClassName & "'"
Set rsSmall= Server.CreateObject("ADODB.Recordset")
rsSmall.open sqlSmall,conn,1,1
EnSmallClassName=rsSmall("EnSmallClassName")
rsSmall.close

if BigClassName="" then
	founderr=true
	errmsg=errmsg+"<li>未指定文章所属大类</li>"
end if
if Title="" then
	founderr=true
	errmsg="<li>文章标题不能为空</li>"
end if
if yzcv<>zcv then
	founderr=true
	errmsg="<li>英文文章标题不能为空</li>"
end if
if Key="" then
	founderr=true
	errmsg=errmsg+"<li>请输入文章关键字</li>"
end if
if Content="" then
	founderr=true
	errmsg=errmsg+"<li>文章内容不能为空</li>"
end if

if founderr=false then
	Title=dvhtmlencode(Title)
	Key=replace(replace(replace(replace(replace(replace(replace(Key,"'",""),"*",""),"?",""),"(",""),")",""),",",""),".","")
	Key=Key & " "
	Content=ubbcode(Content)
	EnContent=ubbcode(EnContent)
	if UpdateTime<>"" and IsDate(UpdateTime)=true then
		UpdateTime=CDate(UpdateTime)
	else
		UpdateTime=now()
	end if
	if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
	end if
	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		sql="select top 1 * from Bs_Product" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()
		'rs("Editor")=Editor
		rs.update
		ArticleID=rs("ArticleID")
		rs.close
		set rs=nothing
	elseif request("action")="Modify" then
  		if ArticleID<>"" then
			sql="select * from Bs_Product where articleid=" & ArticleID
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
				call SaveData()
				rs.update
				rs.close
				set rs=nothing
 			else
				founderr=true
				errmsg=errmsg+"<li>找不到此文章，可能已经被其他人删除。</li>"
				call WriteErrMsg()
			end if
		else
			founderr=true
			errmsg=errmsg+"<li>不能确定ArticleID的值</li>"
			call WriteErrMsg()
		end if
	else
		founderr=true
		errmsg=errmsg+"<li>没有选定参数</li>"
		call WriteErrMsg()
	end if

	call CloseConn()
%>
<!-- #include file="Inc/Head.asp" -->
<BR>
<table cellpadding="2" cellspacing="1" border="0" width="510" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>添加 修改 产品</strong></td>
	</tr>
	<tr class="a4">
		<td align="center">
      <table class="border" align=center width="500" border="0" cellpadding="0" cellspacing="2" bordercolor="#999999">
        <tr align=center bgcolor="#999999"> 
          <td  height="25" colspan="2" class="title"><b> 
            <%if request("action")="add" then%>
            添加 
            <%else%>
            修改 
            <%end if%>
            产品成功</b></td>
        </tr>
        <tr> 
          <td width="19%" height="22" bgcolor="#C0C0C0" class="tdbg"> <p align="right">产品序号：</p></td>
          <td width="81%" bgcolor="#E3E3E3" class="tdbg"><%=ArticleID%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">产品编号：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Product_Id%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">产品名称：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Title%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">English名称：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=EnTitle%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">产品规格：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Spec%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">English规格：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=EnSpec%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">尺　寸：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Size%></td>
        </tr>
<!--         <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">产品备注：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Memo%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">English备注：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=EnMemo%></td>
        </tr> -->
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">所属类别：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"> 
<%
response.write BigClassName
if SmallClassName<>"" then response.write " &gt;&gt; " & SmallClassName
if SpecialName<>"" then response.write "<br>所属专题：" & SpecialName
%>
          </td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">关 
              键 字：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Key%></td>
        </tr>
        <tr> 
          <td height="22" colspan="2" bgcolor="#E3E3E3" class="tdbg"> 
            <p align="center"> 【<a href="Bs_Product_edit.asp?ArticleID=<%=ArticleID%>">修改产品</a>】&nbsp;【<a href="Bs_Product_Add.asp">继续添加产品</a>】&nbsp;【<a href="Bs_Product.asp">产品管理</a>】&nbsp;【<a href="../Chinese/Bs_ProductShow.asp?ArticleID=<%=ArticleID%>" target="_blank">预览产品内容</a>】</p></td>
        </tr>
      </table>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>
<%
else
	WriteErrMsg
end if

sub SaveData()
	rs("Product_Id")=Product_Id
	rs("BigClassName")=BigClassName
	rs("EnBigClassName")=EnBigClassName
	rs("SmallClassName")=SmallClassName
	rs("EnSmallClassName")=EnSmallClassName
	'rs("SpecialName")=SpecialName
	rs("Title")=Title
	rs("EnTitle")=EnTitle
	rs("Spec")=Spec
	rs("EnSpec")=EnSpec
	rs("Size")=Size
'	rs("Memo")=Memo
'	rs("EnMemo")=EnMemo
	rs("Content")=Content
	rs("EnContent")=EnContent
	rs("Key")=Key
	rs("Hits")=Hits
	'rs("Author")=Author
	'rs("CopyFrom")=CopyFrom
	if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
	if Passed="yes" then
		rs("Passed")=True
	else
		if EnableArticleCheck="No" then
			rs("Passed")=True
		else
			rs("Passed")=False
		end if
	end if
	'if OnTop="yes" then
		'rs("OnTop")=True
	'else
		'rs("OnTop")=False
	'end if
	'if Hot="yes" then
		'rs("Hot")=True
	'else
		'rs("Hot")=False
	'end if
	if Elite="yes" then
		rs("Elite")=True
	else
		rs("Elite")=False
	end if
	if EnElite="yes" then
		rs("EnElite")=True
	else
		rs("EnElite")=False
	end if
	rs("UpdateTime")=UpdateTime

	'***************************************
	'删除无用的上传文件
	if ObjInstalled=True and DelUploadProductPic="Yes" then
		dim fso,strRubbishFile
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if instr(UploadFiles,"|")>1 then
			dim arrUploadFiles,intTemp
			arrUploadFiles=split(UploadFiles,"|")
			UploadFiles=""
			for intTemp=0 to ubound(arrUploadFiles)
				if instr(Content,arrUploadFiles(intTemp))<=0 and arrUploadFiles(intTemp)<>DefaultPicUrl then
					strRubbishFile=server.MapPath("" & arrUploadFiles(intTemp))
					if fso.FileExists(strRubbishFile) then
						fso.DeleteFile(strRubbishFile)
						response.write "<br><li>" & arrUploadFiles(intTemp) & "在文章中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
					end if
				else
					if intTemp=0 then
						UploadFiles=arrUploadFiles(intTemp)
					else
						UploadFiles=UploadFiles & "|" & arrUploadFiles(intTemp)
					end if
				end if
			next
		else
			if instr(Content,UploadFiles)<=0 and UploadFiles<>DefaultPicUrl then
				strRubbishFile=server.MapPath("" & UploadFiles)
				if fso.FileExists(strRubbishFile) then
					fso.DeleteFile(strRubbishFile)
					response.write "<br><li>" & UploadFiles & "在文章中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
				end if
				UploadFiles=""
			end if
		end if
		set fso=nothing
	end If
	'结束
	'***************************************
	rs("DefaultPicUrl")=DefaultPicUrl
	rs("UploadFiles")=UploadFiles
end sub
	
%>
