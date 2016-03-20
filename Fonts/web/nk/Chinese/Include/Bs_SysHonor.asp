<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
<%
dim strFileName,MaxPerPage
dim totalPut,CurrentPage,TotalPages
dim sqlInfo,rsInfo
dim Action,i
Action=trim(request("Action"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
%> 

<%
'------------ShowHonorTotal------------
sub ShowHonorTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_Honor"
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 条信息"
	else
		totalPut=rsTotal(0)
		Response.Write "共有 " & totalPut & " 条信息"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub
%> 

<%
'------------ShowImgTotal------------
sub ShowImgTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_Img"
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 条信息"
	else
		totalPut=rsTotal(0)
		Response.Write "共有 " & totalPut & " 条信息"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub
%> 

<%
'------------ShowHonorInfo------------
sub ShowHonorInfo()
	if currentpage<1 then
		currentpage=1
	end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
  end if
	if currentPage=1 then
        sqlInfo="select top " & MaxPerPage	
	else
		sqlInfo="select "
	end if

	sqlInfo=sqlInfo & " * from Bs_Honor Order BY ID desc"
	Set rsInfo= Server.CreateObject("ADODB.Recordset")
	rsInfo.open sqlInfo,conn,1,1

	if rsInfo.bof and  rsInfo.eof then
		Response.Write("<br><li>还没任何信息</li>")
	else
		if currentPage=1 then
			call ShowContent()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rsInfo.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rsInfo.bookmark
				call ShowContent()
      else
				currentPage=1
				call ShowContent()
	    end if
		end if
	end if

	rsInfo.close
	set rsInfo=nothing
end sub
%> 

<%
'------------ShowImgInfo------------
sub ShowImgInfo()
	if currentpage<1 then
		currentpage=1
	end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
  end if
	if currentPage=1 then
        sqlInfo="select top " & MaxPerPage	
	else
		sqlInfo="select "
	end if

	sqlInfo=sqlInfo & " * from Bs_Img Order BY ID desc"
	Set rsInfo= Server.CreateObject("ADODB.Recordset")
	rsInfo.open sqlInfo,conn,1,1

	if rsInfo.bof and  rsInfo.eof then
		Response.Write("<br><li>还没任何信息</li>")
	else
		if currentPage=1 then
			call ShowContent()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rsInfo.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rsInfo.bookmark
				call ShowContent()
      else
				currentPage=1
				call ShowContent()
	    end if
		end if
	end if

	rsInfo.close
	set rsInfo=nothing
end sub
%> 

<%
'------------ShowContent------------
Sub ShowContent()
i=1
Response.write "<Table width=100% border=0 cellspacing=0 cellpadding=0>"
i=1
Response.write "<TR><TD height=10><IMG src=img/1x1_pix.gif width=5 height=1></TD>"
Response.write "</TR><TR>"
Do While Not rsInfo.EOF
Response.write "<TD><Table width=100% border=0 cellspacing=0 cellpadding=0><TR>"
Response.write "<td><div align=center><a href='"& rsInfo("img") &"'target='_blank'><img src='"& rsInfo("img") &"' width=250 height=200 border='0'></a></div></td>"
Response.write "</TR><TR>"
Response.write "<td height=10></td>"
Response.write "</TR><TR>"
Response.write "<td><div align=center>"& rsInfo("title") &"</div></td>"
Response.write "</TR><TR>"
Response.write "<td height=10> </td>"
Response.write "</TR></Table></TD>"
if i mod 2 <>0 then
end if
if i mod 2 =0 then
Response.write "</TR><TR>"
end if
i=i+1 
if i >= MaxPerPage then exit do 
rsInfo.MoveNext 
loop
Response.write "</TR></Table>"
End Sub
%>
