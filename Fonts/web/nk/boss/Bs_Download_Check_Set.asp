<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Config.asp"-->
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
dim Bs_DownID,Action,sqlDel,rsDel,FoundErr,ErrMsg,PurviewChecked,ObjInstalled

Bs_DownID=trim(request("Bs_DownID"))
Action=Trim(Request("Action"))
FoundErr=False
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

if Bs_DownID="" or Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
end if
if FoundErr=False then
	if instr(Bs_DownID,",")>0 then
		dim idarr,i
		idArr=split(Bs_DownID)
		for i = 0 to ubound(idArr)
			call CheckArticle(clng(idarr(i)),Action)
		next
	else
		call CheckArticle(clng(Bs_DownID),Action)
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "Bs_Download_Check.asp"
else
	call CloseConn()
	call WriteErrMsg()
end if

sub CheckArticle(ID,CheckAction)
	PurviewChecked=False
	sqlDel="select * from Bs_Download where Bs_DownID=" & CLng(ID)
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	if rsDel.bof and rsDel.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ������£�" & rsDel("Bs_DownID") & " </li>"
	end if
		if CheckAction="Check" then
			rsDel("Bs_Passed")=True
		elseif CheckAction="CancelCheck" then
			rsDel("Bs_Passed")=False
		end if
		rsDel.update
		set rsDel=nothing
end sub
%>
