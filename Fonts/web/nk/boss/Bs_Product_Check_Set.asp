<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Config.asp"-->
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
dim ArticleID,Action,sqlDel,rsDel,FoundErr,ErrMsg,PurviewChecked,ObjInstalled

ArticleID=trim(request("ArticleID"))
Action=Trim(Request("Action"))
FoundErr=False
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

if ArticleId="" or Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if instr(ArticleID,",")>0 then
		dim idarr,i
		idArr=split(ArticleID)
		for i = 0 to ubound(idArr)
			call CheckArticle(clng(idarr(i)),Action)
		next
	else
		call CheckArticle(clng(ArticleID),Action)
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "Bs_Product_Check.asp"
else
	call CloseConn()
	call WriteErrMsg()
end if

sub CheckArticle(ID,CheckAction)
	PurviewChecked=False
	sqlDel="select * from Bs_Product where ArticleID=" & CLng(ID)
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	if rsDel.bof and rsDel.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到文章：" & rsDel("ArticleID") & " </li>"
	end if
		if CheckAction="Check" then
			rsDel("Passed")=True
		elseif CheckAction="CancelCheck" then
			rsDel("Passed")=False
		end if
		rsDel.update
		set rsDel=nothing
end sub
%>
