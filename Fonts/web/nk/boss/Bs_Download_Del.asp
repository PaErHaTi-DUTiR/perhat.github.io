<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="../Inc/Config.asp"-->
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
dim Bs_DownID,Action,sqlDel,rsDel,FoundErr,ErrMsg,ObjInstalled

Bs_DownID=trim(request("Bs_DownID"))
Action=Trim(Request("Action"))
FoundErr=False
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

if Bs_DownID="" or Action<>"Del" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if instr(Bs_DownID,",")>0 then
		dim idarr,i
		idArr=split(Bs_DownID)
		for i = 0 to ubound(idArr)
			call DelArticle(clng(idarr(i)))
		next
	else
		call DelArticle(clng(Bs_DownID))
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "Bs_Download.asp"
else
	call CloseConn()
	call WriteErrMsg()
end if

sub DelArticle(ID)
	PurviewChecked=False
	sqlDel="select * from Bs_Download where Bs_DownID=" & CLng(ID)
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	if FoundErr=False then
		if DelUpFiles="Yes" and ObjInstalled=True then
			dim fso,strUploadFiles,arrUploadFiles
			strUploadFiles=rsDel("Bs_UploadFiles") & ""
			if strUploadFiles<>"" then
				Set fso = CreateObject("Scripting.FileSystemObject")
				if instr(strUploadFiles,"|")>1 then
					arrUploadFiles=split(strUploadFiles,"|")
					for i=0 to ubound(arrUploadFiles)
						if fso.FileExists(server.MapPath("../" & arrUploadfiles(i))) then
							fso.DeleteFile(server.MapPath("../" & arrUploadfiles(i)))
						end if
					next
				else
					if fso.FileExists(server.MapPath("../" & strUploadfiles)) then
						fso.DeleteFile(server.MapPath("../" & strUploadfiles))
					end if
				end if
				Set fso = nothing
			end if
		end if
		rsDel.delete
		rsDel.update
		set rsDel=nothing
		'conn.execute "delete from Comment where Bs_DownID=" & CLng(ID)
	end if
end sub
%>
