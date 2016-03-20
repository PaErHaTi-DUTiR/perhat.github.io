<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
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
dim Bs_DownID,Bs_BigClassName,Bs_SmallClassName
dim Bs_Title,Bs_UpDateTime,Bs_FileSize,Bs_Language,Bs_LicenceType,Bs_System,Bs_Content,Bs_Passed,Bs_Elite,Bs_PhotoUrl,Bs_DownloadUrl,Bs_DownloadUrl2
dim ObjInstalled

ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=false
Bs_DownID=Trim(Request.Form("Bs_DownID"))
Bs_BigClassName=trim(request.form("Bs_BigClassName"))
Bs_SmallClassName=trim(request.form("Bs_SmallClassName"))
Bs_Title=trim(request.form("Bs_Title"))
Bs_UpDateTime=trim(request.form("Bs_UpDateTime"))
Bs_FileSize=trim(request.form("Bs_FileSize"))
Bs_Language=trim(request.form("Bs_Language"))
Bs_LicenceType=trim(request.form("Bs_LicenceType"))
Bs_System=trim(request.form("Bs_System"))
Bs_Content=trim(request.form("Bs_Content"))
Bs_Passed=trim(request.form("Bs_Passed"))
Bs_Elite=trim(request.form("Bs_Elite"))
Bs_PhotoUrl=trim(request.form("Bs_PhotoUrl"))
Bs_DownloadUrl=trim(request.form("Bs_DownloadUrl"))
Bs_DownloadUrl2=trim(request.form("Bs_DownloadUrl2"))
'============
sqlBig="select * from Bs_DownBigClass where Bs_BigClassName='" & Bs_BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
rsBig.close

sqlSmall="select * from Bs_DownSmallClass where Bs_SmallClassName='" & Bs_SmallClassName & "'and  Bs_BigClassName='" & Bs_BigClassName & "'"
Set rsSmall= Server.CreateObject("ADODB.Recordset")
rsSmall.open sqlSmall,conn,1,1
rsSmall.close
'============
if Bs_BigClassName="" then
	founderr=true
	errmsg=errmsg+"<li>未指定下载所属大类！</li>"
end if
if Bs_Title="" then
	founderr=true
	errmsg="<li>下载标题不能为空！</li>"
end if
if yzcv<>zcv then
	founderr=true
	errmsg="<li>文件大小不能为空</li>"
end if
'if Bs_FileSize="" then
'	founderr=true
'	errmsg="<li>文件大小不能为空！</li>"
'end if
if Bs_DownloadUrl="" then
	founderr=true
	errmsg=errmsg+"<li>请输入下载连接！</li>"
end if
if Bs_Content="" then
	founderr=true
	errmsg=errmsg+"<li>软件简介容不能为空！</li>"
end if
'============
if founderr=false then
	Bs_Title=htmlencode2(Bs_Title)
	Bs_Content=htmlencode2(Bs_Content)
	Bs_FileSize=replace(replace(replace(replace(replace(replace(Bs_FileSize,"'",""),"*",""),"?",""),"(",""),")",""),",","")
	Bs_FileSize=Bs_FileSize & " "

	if Bs_UpDateTime<>"" and IsDate(Bs_UpDateTime)=true then
		Bs_UpDateTime=CDate(Bs_UpDateTime)
	else
		Bs_UpDateTime=now()
	end if
'============
	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		sql="select top 1 * from Bs_Download" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()
'		rs("Editor")=Editor
		rs.update
		Bs_DownID=rs("Bs_DownID")
		rs.close
		set rs=nothing
	elseif request("action")="Modify" then
  		if Bs_DownID<>"" then
			sql="select * from Bs_Download where Bs_DownID=" & Bs_DownID
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
				call SaveData()
				rs.update
				rs.close
				set rs=nothing
 			else
				founderr=true
				errmsg=errmsg+"<li>找不到此下载，可能已经被其他人删除。</li>"
				call WriteErrMsg()
			end if
		else
			founderr=true
			errmsg=errmsg+"<li>不能确定Bs_DownID的值</li>"
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
		<td class="a1" height="25" align="center"><strong>添加 修改 下载</strong></td>
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
            下载成功</b></td>
        </tr>
        <tr> 
          <td width="19%" height="22" bgcolor="#C0C0C0" class="tdbg"> <p align="right">下载序号：</p></td>
          <td width="81%" bgcolor="#E3E3E3" class="tdbg"><%=Bs_DownID%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">标题名称：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_Title%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">整理日期：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_UpDateTime%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">文件大小：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_FileSize%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">软件语言：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_Language%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">授权类型：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_LicenceType%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">运行环境：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_System%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">图片地址：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_PhotoUrl%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">下载地址一</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Bs_DownloadUrl%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">下载地址二</div></td>
          <td bgcolor="#E3E3E3" class="tdbg">
<%
if Bs_DownloadUrl2<>"" then
	response.Write Bs_DownloadUrl2
	else
	response.Write "-"
end if
%> </td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">所属类别：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"> 
<%
response.write Bs_BigClassName
if Bs_SmallClassName<>"" then response.write " &gt;&gt; " & Bs_SmallClassName
%>
          </td>
        </tr>
        <tr> 
          <td height="22" colspan="2" bgcolor="#E3E3E3" class="tdbg"> 
            <p align="center"> 【<a href="Bs_Download_edit.asp?Bs_DownID=<%=Bs_DownID%>">修改下载</a>】&nbsp;【<a href="Bs_Download_Add.asp">继续添加下载</a>】&nbsp;【<a href="Bs_Download.asp">下载管理</a>】&nbsp;【<a href="../Chinese/Bs_DownloadShow.asp?Bs_DownID=<%=Bs_DownID%>" target="_blank">预览下载内容</a>】</p></td>
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
	rs("Bs_BigClassName")=Bs_BigClassName
	rs("Bs_SmallClassName")=Bs_SmallClassName
	rs("Bs_Title")=Bs_Title
	rs("Bs_FileSize")=Bs_FileSize
	rs("Bs_Language")=Bs_Language
	rs("Bs_LicenceType")=Bs_LicenceType
	rs("Bs_System")=Bs_System
	rs("Bs_Content")=Bs_Content
	if Bs_Passed="yes" then
		rs("Bs_Passed")=True
	else
		if EnableArticleCheck="No" then
			rs("Bs_Passed")=True
		else
			rs("Bs_Passed")=False
		end if
	end if
	if Bs_Elite="yes" then
		rs("Bs_Elite")=True
	else
		rs("Bs_Elite")=False
	end if
	rs("Bs_UpDateTime")=Bs_UpDateTime
	rs("Bs_PhotoUrl")=Bs_PhotoUrl
	rs("Bs_DownloadUrl")=Bs_DownloadUrl
	rs("Bs_DownloadUrl2")=Bs_DownloadUrl2
end sub
	
%>
