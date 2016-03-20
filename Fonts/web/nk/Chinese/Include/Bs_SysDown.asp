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
dim strFileName,MaxPerPage,ShowSmallClassType
dim totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime
dim founderr, errmsg
dim BigClassName,SmallClassName,keyword,strField
dim Bs_BigClassName,Bs_SmallClassName,sql_DownBigClass,rs_DownBigClass,sql_Down,rs_Down
dim rs,sql,sqlArticle,rsArticle,sqlSearch,rsSearch,sqlBigClass,rsBigClass,sqlSpecial,rsSpecial
dim SpecialTotal
BeginTime=Timer
Bs_BigClassName=Trim(request("Bs_BigClassName"))
Bs_SmallClassName=Trim(request("Bs_SmallClassName"))
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=replace(replace(replace(replace(keyword,"'","‘"),"<","&lt;"),">","&gt;")," ","&nbsp;")
end if
strField=trim(request("Field"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sql_DownBigClass="select * from Bs_DownBigClass order by Bs_BigClassID"
Set rs_DownBigClass= Server.CreateObject("ADODB.Recordset")
rs_DownBigClass.open sql_DownBigClass,conn,1,1

'=================================================
'过程名：Show_DownSmallClass_Tree
'作  用：树形目录方式显示栏目
'参  数：无
'=================================================
sub Show_DownSmallClass_Tree()
%>
<SCRIPT language=javascript>
function opencat(cat,img){
	if(cat.style.display=="none"){
	cat.style.display="";
	img.src="../img/class2.gif";
	}	else {
	cat.style.display="none"; 
	img.src="../img/class1.gif";
	}
}
</Script>
<TABLE cellSpacing=0 cellPadding=0 width="99%" border=0>
<%
dim i
if rs_DownBigClass.eof and rs_DownBigClass.bof then
	Response.Write "栏目正在建设中……"
else
	i=1
	do while not rs_DownBigClass.eof
%>
	<TR>
		<TD language=javascript onmouseup="opencat(cat10<%=i%>000,&#13;&#10; img10<%=i%>000);" id=item$pval[catID]) style="CURSOR: hand" width=36 height=24 align=center><IMG id=img10<%=i%>000 src="../img/class1.gif" width=20 height=20></TD>
		<TD class='ML_Tdbg'><a href='Bs_Download.asp?Bs_BigClassName=<%=rs_DownBigClass("Bs_BigClassName")%>'><%=rs_DownBigClass("Bs_BigClassName")%></a></TD>
	</TR>
	<TR>
		<TD id=cat10<%=i%>000 style="DISPLAY: none" bgColor=#fefdf5 colspan="2">
<%
dim rss,sqls,j
set rss = server.CreateObject ("adodb.recordset")
sqls="select * from Bs_DownSmallClass where Bs_BigClassName='" & rs_DownBigClass("Bs_BigClassName") & "'order by Bs_SmallClassID"
rss.open sqls,conn,1,1
if rss.eof and rss.bof then
	Response.Write "没有小分类"
else
	j=1
	do while not rss.eof
%>
&nbsp;<IMG height=20 src="../img/class3.gif" width=26 align=absMiddle border=0 >&nbsp;<a href="Bs_Download.asp?Bs_BigClassName=<%=rss("Bs_BigClassName")%>&Bs_Smallclassname=<%=rss("Bs_SmallClassName")%>"><%=rss("Bs_SmallClassName")%></a><BR>
<%
	rss.movenext
	j=j+1
	loop
end if
rss.close
set rss=nothing
%>
		</TD>
	</TR>
<%
	rs_DownBigClass.movenext
	i=i+1
	loop
end if
%>
</TABLE>
<%
end sub

'=================================================
'过程名：ShowClass_DownGuide
'作  用：显示栏目导航位置
'参  数：无
'=================================================
sub ShowClass_DownGuide()
	response.write  "<a href='Bs_Download.asp'>下载中心</a>&nbsp;&gt;&gt;"
	if Bs_BigClassName="" and Bs_SmallClassName="" then
		response.write "&nbsp;所有下载"
	else
		if Bs_BigClassName<>"" then
			response.write "&nbsp;<a href='Bs_Download.asp?Bs_BigClassName=" & Bs_BigClassName & "'>" & Bs_BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if Bs_SmallClassName<>"" then
				response.write "<a href='Bs_Download.asp?Bs_BigClassName=" & Bs_BigClassName & "&Bs_SmallClassName=" & Bs_SmallClassName & "'>" & Bs_SmallClassName & "</a>"
			else
				response.write "所有小类"
			end if
		end if
	end if
end sub

'=================================================
'过程名：ShowDownTotal
'作  用：显示下载总数
'参  数：无
'=================================================
sub ShowDownTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_Download where Bs_Passed=True "
	if Bs_BigClassName<>"" then
		sqlTotal=sqlTotal & " and Bs_BigClassName='" & Bs_BigClassName & "'"
		if Bs_SmallClassName<>"" then
			sqlTotal=sqlTotal & " and Bs_SmallClassName='" & Bs_SmallClassName & "'"
		end if
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 个下载"
	else
		totalPut=rsTotal(0)
		response.Write "共有 " & totalPut & " 个下载"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub

'=================================================
'过程名：ShowDown
'=================================================
sub ShowDown(TitleLen)
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
        sql_Down="select top " & MaxPerPage	
	else
		sql_Down="select "
	end if

	sql_Down=sql_Down & " Bs_DownID,Bs_BigClassName,Bs_SmallClassName,Bs_Title,Bs_UpDateTime,Bs_FileSize,Bs_Language,Bs_LicenceType,Bs_System,Bs_Hits,Bs_DownloadUrl,Bs_Content from Bs_Download where Bs_Passed=True"
	
	if Bs_BigClassName<>"" then
		sql_Down=sql_Down & " and Bs_BigClassName='" & Bs_BigClassName & "'"
		if Bs_SmallClassName<>"" then
			sql_Down=sql_Down & " and Bs_SmallClassName='" & Bs_SmallClassName & "'"
		end if
	end if
	sql_Down=sql_Down & " order by Bs_DownID desc"
	Set rs_Down= Server.CreateObject("ADODB.Recordset")
	rs_Down.open sql_Down,conn,1,1
	if rs_Down.bof and  rs_Down.eof then
		response.Write("<br><li>没有任何下载</li>")
	else
		if currentPage=1 then
			call DownContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs_Down.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs_Down.bookmark
            	call DownContent(TitleLen)
        	else
	        	currentPage=1
           		call DownContent(TitleLen)
	    	end if
		end if
	end if
	rs_Down.close
	set rs_Down=nothing
end sub

sub DownContent(intTitleLen)
   	dim i,lb,strTemp
    i=0
	do while not rs_Down.eof
		lb=""
		lb=lb & "<table width=98% border=0 cellSpacing=1 cellpadding=3 bgColor=#b8b8b8>"
			lb=lb & "<tr><td width=60% height=22 bgColor=#ddffd0>"
			lb=lb & "<a href=Bs_DownloadShow.asp?Bs_DownID=" & rs_Down("Bs_DownID") & ">&nbsp;<b>" & rs_Down("Bs_Title") & "</b></a>"
			lb=lb & "</td>"
			lb=lb & "<td bgColor=#ddffd0 align=right>"
			lb=lb & "<a href=Bs_Download.asp?Bs_BigClassName=" & rs_Down("Bs_BigClassName") & ">" & rs_Down("Bs_BigClassName") & "</a> | "
			lb=lb & "<a href=Bs_Download.asp?Bs_BigClassName=" & rs_Down("Bs_BigClassName") & "&Bs_SmallClassName=" & rs_Down("Bs_SmallClassName") & ">"
			lb=lb & rs_Down("Bs_SmallClassName") & ""
			lb=lb & "</a>&nbsp;</td></tr>"

			lb=lb & "<tr><td height=22 bgColor=#ffffff>&nbsp"
			lb=lb & "<FONT COLOR=#0000FF>整理日期：</FONT>" & FormatDateTime(rs_Down("Bs_UpDateTime"),2) & "&nbsp;&nbsp;"
			lb=lb & "<FONT COLOR=#0000FF>文件大小：</FONT>" & rs_Down("Bs_FileSize") & "</td>"
			lb=lb & "<td height=22 bgColor=#ffffff>"
			lb=lb & "<table width=100% border=0 cellpadding=0 cellspacing=0 bgColor=#ffffff><tr>"
			lb=lb & "<td width=40% height=22>&nbsp"
			lb=lb & "<a href=" & rs_Down("Bs_DownloadUrl") & " target=_blank><FONT COLOR=#0000FF>下载</FONT></a> | "
			lb=lb & "<a href=Bs_DownloadShow.asp?Bs_DownID=" & rs_Down("Bs_DownID") & "><FONT COLOR=#0000FF>查看</FONT></a>&nbsp;"
			lb=lb & "</td><td align=center>"
			lb=lb & "人气：" & rs_Down("Bs_Hits") & ""
			lb=lb & "</td></tr></table>"
			lb=lb & "</td></tr>"

			lb=lb & "<tr><td height=22 bgColor=#f5f5f5 colspan=2><FONT COLOR=#FF0000>·</FONT>"
			if rs_Down("Bs_Content")<>"" then
				strTemp=replace(rs_Down("Bs_Content"),"<br>","")
				strTemp=replace(strTemp,"<BR>","")
				strTemp=replace(strTemp,"<b>","")
				strTemp=replace(strTemp,"<","")
				strTemp=replace(strTemp,">","")
				strTemp=replace(strTemp,"&nbsp;","")
				strTemp=replace(strTemp,"　","")
				strTemp=replace(strTemp," ","")
				lb=lb&""& gotTopic(strTemp,160)&""
				else
				lb=lb&"-"
			end if
			lb=lb & "</td></tr>"

			lb=lb & "<tr><td width=60% height=22 bgColor=#ffffff colspan=2>&nbsp;"
			lb=lb & "<FONT COLOR=#0000FF>软件语言：</FONT>" & rs_Down("Bs_Language") & "&nbsp;&nbsp;"
			lb=lb & "<FONT COLOR=#0000FF>授权类型：</FONT>" & rs_Down("Bs_LicenceType") & "&nbsp;&nbsp;"
			lb=lb & "<FONT COLOR=#0000FF>运行环境：</FONT>" & rs_Down("Bs_System") & ""
			lb=lb & "</td></tr>"
			lb=lb & ""

			lb=lb & "<tr><td height=1 colspan=2 bgcolor=#CCCCCC></td>"
			lb=lb & "</tr></table><BR>"
		response.write lb
		rs_Down.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 
%>
