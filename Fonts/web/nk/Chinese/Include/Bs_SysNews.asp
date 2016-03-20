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
dim strUrlFileName,strFileName,MaxPerPage
dim totalPut,CurrentPage,TotalPages
dim sqlNews,rsNews
dim Action
dim id,rsM,sqlM,nCounter,content
Action=trim(request("Action"))
strUrlFileName="Bs_NewsInfo.asp?Action="& Action &"&"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
%> 

<%
'------------ShowNewsCoTotal------------
sub ShowNewsCoTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_NewsCo"
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 条新闻"
	else
		totalPut=rsTotal(0)
		Response.Write "共有 " & totalPut & " 条新闻"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub
%> 

<%
'------------ShowNewsYeTotal------------
sub ShowNewsYeTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_NewsYe"
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 条新闻"
	else
		totalPut=rsTotal(0)
		Response.Write "共有 " & totalPut & " 条新闻"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub
%> 

<%
'------------ShowNewsPrTotal------------
sub ShowNewsPrTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_NewsPr"
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 条新闻"
	else
		totalPut=rsTotal(0)
		Response.Write "共有 " & totalPut & " 条新闻"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub
%> 

<%
'------------ShowNewsCo------------
sub ShowNewsCo(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if
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
        sqlNews="select top " & MaxPerPage	
	else
		sqlNews="select "
	end if

	sqlNews=sqlNews & " * from Bs_NewsCo Order BY ID desc"
	Set rsNews= Server.CreateObject("ADODB.Recordset")
	rsNews.open sqlNews,conn,1,1

	if rsNews.bof and  rsNews.eof then
		Response.Write("<br><li>还没任何新闻</li>")
	else
		if currentPage=1 then
			call NewsContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rsNews.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rsNews.bookmark
				call NewsContent(TitleLen)
      else
				currentPage=1
				call NewsContent(TitleLen)
	    end if
		end if
	end if

	rsNews.close
	set rsNews=nothing
end sub
%> 

<%
'------------ShowNewsYe------------
sub ShowNewsYe(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if
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
        sqlNews="select top " & MaxPerPage	
	else
		sqlNews="select "
	end if

	sqlNews=sqlNews & " * from Bs_NewsYe Order BY ID desc"
	Set rsNews= Server.CreateObject("ADODB.Recordset")
	rsNews.open sqlNews,conn,1,1

	if rsNews.bof and  rsNews.eof then
		Response.Write("<br><li>还没任何新闻</li>")
	else
		if currentPage=1 then
			call NewsContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rsNews.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rsNews.bookmark
				call NewsContent(TitleLen)
      else
				currentPage=1
				call NewsContent(TitleLen)
	    end if
		end if
	end if

	rsNews.close
	set rsNews=nothing
end sub
%> 

<%
'------------ShowNewsPr------------
sub ShowNewsPr(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if
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
        sqlNews="select top " & MaxPerPage	
	else
		sqlNews="select "
	end if

	sqlNews=sqlNews & " * from Bs_NewsPr Order BY ID desc"
	Set rsNews= Server.CreateObject("ADODB.Recordset")
	rsNews.open sqlNews,conn,1,1

	if rsNews.bof and  rsNews.eof then
		Response.Write("<br><li>还没任何新闻</li>")
	else
		if currentPage=1 then
			call NewsContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rsNews.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rsNews.bookmark
				call NewsContent(TitleLen)
      else
				currentPage=1
				call NewsContent(TitleLen)
	    end if
		end if
	end if

	rsNews.close
	set rsNews=nothing
end sub
%> 

<%
'------------NewsContent------------
sub NewsContent(intTitleLen)
   	dim i,lb
    i=0
	do while not rsNews.eof
			lb=""
			lb=lb & "<table width='100%'border='0'cellpadding='0'cellspacing='0'>"
			lb=lb & "<TR valign='middle'>" 
			lb=lb & "<td style='line-height: 150%;'height=28>&nbsp;&nbsp;<IMG src='../Skin/Img/article_common.gif' width=9 height=15 border=0>&nbsp;"
			lb=lb & "<a href='"& strUrlFileName &"id="& rsNews("id") &"'>"& rsNews("Title") &"</a>&nbsp;&nbsp;"
			lb=lb & "<font color='#0066CC'face='Arial'>"& rsNews("time") &"</font>&nbsp; "
 if rsNews("time")=date() then 
			lb=lb & "<strong><font color='#FF0000'face='Arial'>New</font></strong>"
 else
			lb=lb & ""
 end if
			lb=lb & "</td>"
			lb=lb & "</TR>"
			lb=lb & "<TR>"
			lb=lb & "<TD width='100%'height=1 background='../img/kropki_poziom_naBialym.gif'><IMG height=1 src='kropki_poziom_naBialym.gif' width=10></TD>"
			lb=lb & "</TR></table>"

		Response.Write lb
		rsNews.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 
%> 

<%
'------------M_NewsCo------------
Sub M_NewsCo()
if not isEmpty(request.QueryString("id")) then
	id=request.QueryString("id")
else
	id=1
end if

	Set rsM = Server.CreateObject("ADODB.Recordset")
	sqlM="Select * From Bs_NewsCo where id="&id
	rsM.Open sqlM,conn,3,3

'纪录访问次数
rsM("counter")=rsM("counter")+1
rsM.update
nCounter=rsM("counter")
'定义内容
content=ubbcode(rsM("content"))

	Call ShowNewsContent()

rsM.close
set rsM=nothing
End Sub
%>

<%
'------------M_NewsYe------------
Sub M_NewsYe()
if not isEmpty(request.QueryString("id")) then
	id=request.QueryString("id")
else
	id=1
end if

	Set rsM = Server.CreateObject("ADODB.Recordset")
	sqlM="Select * From Bs_NewsYe where id="&id
	rsM.Open sqlM,conn,3,3

'纪录访问次数
rsM("counter")=rsM("counter")+1
rsM.update
nCounter=rsM("counter")
'定义内容
content=ubbcode(rsM("content"))

Call ShowNewsContent()

rsM.close
set rsM=nothing
End Sub
%> 

<%
'------------M_NewsPr------------
Sub M_NewsPr()
if not isEmpty(request.QueryString("id")) then
	id=request.QueryString("id")
else
	id=1
end if

	Set rsM = Server.CreateObject("ADODB.Recordset")
	sqlM="Select * From Bs_NewsPr where id="&id
	rsM.Open sqlM,conn,3,3

'纪录访问次数
rsM("counter")=rsM("counter")+1
rsM.update
nCounter=rsM("counter")
'定义内容
content=ubbcode(rsM("content"))

Call ShowNewsContent()

rsM.close
set rsM=nothing
End Sub
%> 

<%
'------------ShowNewsContent------------
Sub ShowNewsContent()
Response.write "<table width='100%'border='0'cellpadding='0'cellspacing='0'>"
Response.write "<tr>" 
Response.write "<td align='center'style='line-height: 150%'><strong>"& rsM("title") &"</strong><BR><HR></td>"
Response.write "</tr>"
Response.write "<tr>" 
Response.write "<td style='line-height: 150%'><div align='center'>发布日期：["& rsM("time") &"]&nbsp;&nbsp;&nbsp; 共阅["& ncounter &"]次</div></td>"
Response.write "</tr>"
Response.write "<tr><td style='line-height: 150%'>"
Response.write "<SPAN class=content id=BodyLabel style='PADDING-RIGHT: 10px; DISPLAY: block; PADDING-LEFT: 10px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px'>"
Response.write "&nbsp;&nbsp;&nbsp;&nbsp;"& content &"</span>"
Response.write "</td></tr></table>"
End Sub
%> 
