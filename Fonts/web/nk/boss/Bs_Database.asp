<%@ LANGUAGE=VBScript CodePage=936%>
<%
option explicit
response.buffer=true	
Const strDBPath = "Databackup\"
%>
<!--#include file="bsconfig.asp"-->
<%
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
dim dbpath
dim ObjInstalled
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
%>
<!-- #include file="Inc/Head.asp" -->

<BR>
<table cellpadding="2" cellspacing="1" border="0" width="600" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center" colspan="2"><strong>�� �� �� �� ��</strong></td>
	</tr>
	<tr class="a4">
		<td width="75" height="30">&nbsp;<strong>����������</strong></td>
    <td height="30">&nbsp;<a href="Bs_Database.asp?Action=Backup">�������ݿ�</a> | <a href="Bs_Database.asp?Action=Restore">�ָ����ݿ�</a> | <a href="Bs_Database.asp?Action=Compact">ѹ�����ݿ�</a> | <a href="Bs_Database.asp?Action=SpaceSize">ϵͳ�ռ�ռ�����</a>
		</td>
	</tr>
</table>
<BR>

<%
if Action="Backup" or Action="BackupData" then
	call ShowBackup()
	set conn=nothing 
elseif Action="Compact" or Action="CompactData" then
	call ShowCompact()
	set conn=nothing 
elseif Action="Restore" or Action="RestoreData" then
	call ShowRestore()
	set conn=nothing 
elseif Action="SpaceSize" then
	call SpaceSize()
	set conn=nothing 
else
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���������</li>"
	set conn=nothing 
end if
if FoundErr=True then
	call WriteErrMsg()
end if

'--------�������ݿ�--------
sub ShowBackup()
%>
<table cellpadding="2" cellspacing="1" border="0" width="600" align="center" class="a2">
<form method="post" action="Bs_Database.asp?action=BackupData">
  <tr>
	<td colspan="3" align="center" height="25" class="a1"><FONT COLOR="#CC0000"><b>�� �� �� �� ��</b></FONT></td>
  </tr>
<%    
if request("action")="BackupData" then
	call backupdata()
else
%>
          <tr class="a4"> 
            <td width="80" height="40" align="right">��ǰ���ݿ⣺</td>
            <td><input name="db" type="text" size="40" value="<%=db%>"></td>
            <td>���·��Ŀ¼</td>
          </tr>
          <tr class="a4"> 
            <td width="80" height="40" align="right">����Ŀ¼��</td>
            <td><input type=text size=40 name=bkfolder value="Databackup"></td>
            <td>���·��Ŀ¼����Ŀ¼�����ڣ����Զ�����</td>
          </tr>
          <tr class="a4"> 
            <td width="80" height="40" align="right">�������ƣ�</td>
            <td height="40"><input type=text size=40 name=bkDBname value="Bs_DataBack"></td>
            <td height="40">���������ļ�����׺��Ĭ��Ϊ��.asa����������ͬ���ļ���������</td>
          </tr>
          <tr class="a4"> 
            <td height="40" colspan="3" align="center"><input name="submit" type=submit value=" &nbsp;��ʼ����&nbsp; " <%If ObjInstalled=false Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;"></td>
          </tr>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
</form>
</table>
<%
end sub

'--------�ָ����ݿ�--------
sub ShowRestore()
%>
<table cellpadding="2" cellspacing="1" border="0" width="600" align="center" class="a2">
<form method="post" action="Bs_Database.asp?action=RestoreData">
  <tr>
	<td colspan="2" align="center" height="25" class="a1"><FONT COLOR="#CC0000"><b>�� �� �� �� ��</b></FONT></td>
  </tr>
<%
if request("action")="RestoreData" then
	call RestoreData()
else
%>
	<tr class="a4">
		<td width="150" height="40" align="right">�������ݿ�·������ԣ���</td>
		<td height="40"><input name=backpath type=text id="backpath" value="Databackup/Bs_DataBack.asa" size=50 maxlength="200"></td>
	</tr>
	<tr class="a4">
		<td width="150" height="40" align="right">��ǰ���ݿ�·������ԣ���</td>
		<td><input name="db" type="text" size="50" maxlength="200" value="<%=db%>"></td>
	</tr>
	<tr align="center" class="a4"> 
		<td colspan="2"><input name="submit" type=submit value=" &nbsp;�ָ�����&nbsp; " <%If ObjInstalled=false Then Response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;"></td>
	</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
            
</form>
</table>
<%
end sub

'--------ѹ�����ݿ�--------
sub ShowCompact()
%>
<table cellpadding="2" cellspacing="1" border="0" width="600" align="center" class="a2">
<form method="post" action="Bs_Database.asp?action=CompactData">
  <tr align="center">
	<td class="a1" height="25"><FONT COLOR="#CC0000"><b>�� �� �� �� �� ѹ ��</b></FONT></td>
	</tr>
<%    
if request("action")="CompactData" then
	call CompactData()
else
%>
	<tr align="left" class="a4">
		<td valign="middle" style="line-height: 150%"><br><font color="#FF6600"><b>ע1��</b></font>ѹ��ǰ�������ȱ������ݿ⣬���ⷢ��������� <br>
		<font color="#FF6600"><b>ע2��</b></font>����ʹ�������ݿⲻ��ѹ��,��ѡ�񱸷����ݿ����ѹ������(��ǰѹ�����ݿ���ΪĬ�ϱ����ļ���)</td>
	</tr>
	<tr  align="left" class="a4">
		<td height="40">���ݿ�λ�ã� <input name="db" type="text" id="db" size="50" value="Databackup/Bs_DataBack.asa"></td>
	</tr>
	<tr align="center" class="a4">
		<td height="40"><input name="submit" type=submit value=" &nbsp;ѹ�����ݿ�&nbsp; " <%If ObjInstalled=false Then Response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;"></td>
	</tr>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
</form>
</table>
<%
end sub

'--------ͳ�ռ�ռ�����--------
sub SpaceSize()
	on error resume next
%>
<table cellpadding="2" cellspacing="1" border="0" width="600" align="center" class="a2">
  <tr>
	<td colspan="2" align="center" height="25" class="a1"><FONT COLOR="#CC0000"><b>ϵ ͳ �� �� ռ �� �� ��</b></FONT></td>
	</tr>
  <tr class="a4"> 
    <td width="100%" height="150" valign="middle">
	<blockquote><br>
      ϵͳ����ռ�ÿռ䣺&nbsp;<img src="images/bar.gif" width=<%=drawbar("../database")%> height=10>&nbsp;
      <%showSpaceinfo("../database")%>
      <br>
      <br>
      ��������ռ�ÿռ䣺&nbsp;<img src="images/bar.gif" width=<%=drawbar("databackup")%> height=10>&nbsp;
      <%showSpaceinfo("databackup")%>
      <br>
      <br>
      �����ļ�ռ�ÿռ䣺&nbsp;<img src="images/bar.gif" width=<%=drawspecialbar%> height=10>&nbsp;
      <%showSpecialSpaceinfo("Program")%>
      <br>
      <br>
      ��ɫģ��ռ�ÿռ䣺&nbsp;<img src="images/bar.gif" width=<%=drawbar("../Skin")%> height=10>&nbsp;
      <%showSpaceinfo("../Images")%>
      <br>
      <br>
      ϵͳͼƬռ�ÿռ䣺&nbsp;<img src="images/bar.gif" width=<%=drawbar("../Img")%> height=10>&nbsp;
      <%showSpaceinfo("../Img")%>
      <br>
      <br>
      �ϴ��ļ�ռ�ÿռ䣺&nbsp;<img src="images/bar.gif" width=<%=drawbar("../UploadFiles")%> height=10>&nbsp;
      <%showSpaceinfo("../UploadFiles")%>
      <br>
      <br>
      ϵͳռ�ÿռ��ܼƣ�<BR><BR><img src="images/bar.gif" width=400 height=10> 
      <%showspecialspaceinfo("All")%>
	</blockquote> 	
    </td>
  </tr>
</table>
<%
end sub
%>
<BR>
<%htmlend%>

<%
'--------�������ݿ�--------
sub BackupData()
	dim bkfolder,bkdbname,fso
	db=trim(request.form("db"))
	bkfolder=trim(request("bkfolder"))
	bkdbname=trim(request("bkdbname"))
	if db="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�������ݿ�λ�ã�</li>"
	end if
	if bkfolder="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ������Ŀ¼��</li>"
	end if
	if bkdbname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�������ļ���</li>"
	end if
	if FoundErr=True then exit sub
	dbpath = server.mappath(db)
	bkfolder=server.MapPath(bkfolder)
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.FileExists(dbpath) then
		If fso.FolderExists(bkfolder)=false Then
			fso.CreateFolder(bkfolder)
		end if
		fso.copyfile dbpath,bkfolder & "\" & bkdbname & ".asa"
		response.write "<center>�������ݿ�ɹ������ݵ����ݿ�Ϊ " & bkfolder & "\" & bkdbname & ".asa</center>"
	Else
		response.write "<center>�Ҳ���Դ���ݿ��ļ�������inc/conn.asp�е����á�</center>"
	End if
end sub
'--------�ָ����ݿ�--------
sub RestoreData() 
	dim backpath,fso
	backpath=request.form("backpath")
	db=trim(request.form("db"))
	if backpath="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��ԭ���ݵ����ݿ��ļ�����<li>"
		exit sub	
	end if
	if db="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ����ǰ���ݿ��ļ�����<li>"
		exit sub	
	end if
	backpath=server.mappath(backpath)
	dbpath = server.mappath(db)
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.fileexists(backpath) then  					
		fso.copyfile Backpath,dbpath
		response.write "�ɹ��ָ����ݣ�"
	else
		response.write "�Ҳ���ָ���ı����ļ���"	
	end if
end sub
'--------ѹ�����ݿ�--------
sub CompactData() 
	Dim fso, Engine,strDBPath
	DB=trim(request.form("db")) 
	DBPath = server.mappath(DB)
'	strDBPath = server.mappath("data_backup\")
	if instr(DBPath,"/") then 
		strDBPath = left(DBPath,instrrev(DBPath,"\"))
	else
		strDBPath = left(DBPath,instrrev(DBPath,"\"))
	end if
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(DBPath) Then
		Set Engine = CreateObject("JRO.JetEngine")
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath," Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
		fso.CopyFile strDBPath & "temp.mdb",DBPath
		fso.DeleteFile(strDBPath & "temp.mdb")
		Set fso = nothing
		Set Engine = nothing
		response.write "���ݿ�ѹ���ɹ�!"
	Else
		response.write "���ݿ�û���ҵ�!"
	End If
end sub
%>
<%
Sub ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
End Sub	

Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 		
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("../Inc")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath) 		
	
	if method="All" then 		
		size=d.size
	elseif method="Program" then
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next	
	end if	
	
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
end sub 	 	 	

Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("../Inc")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	drvpath=server.mappath(drvpath) 		
	set d=fso.getfolder(drvpath)
	size=d.size
	
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 	

Function Drawspecialbar()
	dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("../Inc")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	set fc=d.files
	for each f1 in fc
		size=size+f1.size
	next	
	
	barsize=cint((size/totalsize)*400)
	Drawspecialbar=barsize
End Function
%>
<%
'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'   False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	dim fso
	folderpath=Server.MapPath(".")&"\"&folderpath
	Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(FolderPath) then
	'����
		CheckDir = True
	Else
	'������
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------����ָ����������Ŀ¼---------
Function MakeNewsDir(foldername)
	dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(foldername)
    MakeNewsDir = True
	Set fso = nothing
End Function

'**************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'**************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type'content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css'rel='stylesheet'type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100'valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub
%>