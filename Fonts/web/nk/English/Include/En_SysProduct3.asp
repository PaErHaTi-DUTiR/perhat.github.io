<%
'������������������������������������������������������������
'�������������� COPYRIGHT ������������ ��
'���������ƣ���ҵ��վ����ϵͳMac3.0  (ASBDcorpweb Mac3.0)  �� 
'����Ȩ���У�WWW.ASBD.CN  BBS.ASBD.CN                      ��
'������������amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  �� 
'��Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ��
'������������������������������������������������������������
%>
<%
dim strFileName,MaxPerPage,ShowSmallClassType
dim totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime
dim founderr, errmsg
dim EnBigClassName,EnSmallClassName,SpecialName,keyword,strField
dim rs,sql,sqlArticle,rsArticle,sqlSearch,rsSearch,sqlBigClass,rsBigClass,sqlSpecial,rsSpecial
dim SpecialTotal
BeginTime=Timer
EnBigClassName=Trim(request("EnBigClassName"))
EnSmallClassName=Trim(request("EnSmallClassName"))
SpecialName=trim(request("SpecialName"))
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=replace(replace(replace(replace(keyword,"'","��"),"<","&lt;"),">","&gt;")," ","&nbsp;")
end if
strField=trim(request("Field"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sqlBigClass="select * from Bs_PrBigClasss order by BigClassID"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1


'=========================================================================
'��������ShowSmallClass_Tree
'��  �ã�����Ŀ¼��ʽ��ʾ��Ŀ
'��  ������
'=========================================================================
Sub ShowSmallClass_Tree()
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
if rsBigClass.eof and rsBigClass.bof then
	Response.Write "The column is under construction����"
else
	i=1
	do while not rsBigClass.eof
%>
	<TR>
		<TD language=javascript onmouseup="opencat(cat10<%=i%>000,&#13;&#10; img10<%=i%>000);" id=item$pval[catID]) style="CURSOR: hand" width=26 height=24 align=center><IMG id=img10<%=i%>000 src="../img/class1.gif" width=20 height=20></TD>
		<TD class='ML_Tdbg'><a href='En_Product.asp?EnBigClassName=<%=rsBigClass("EnBigClassName")%>'><%=rsBigClass("EnBigClassName")%></a></TD>
	</TR>
	<TR>
		<TD id=cat10<%=i%>000 style="DISPLAY: none" bgColor=#fefdf5 colspan="2">
<%
dim rss,sqls,j
set rss = server.CreateObject ("adodb.recordset")
sqls="select * from Bs_PrSmallClass where EnBigClassName='" & rsBigClass("EnBigClassName") & "'order by SmallClassID"
rss.open sqls,conn,1,1
if rss.eof and rss.bof then
	Response.Write "No small class"
else
	j=1
	do while not rss.eof
%>
<IMG height=20 src="../img/class3.gif" width=26 align=absMiddle border=0><a href="En_Product.asp?EnBigClassName=<%=rss("EnBigClassName")%>&EnSmallclassname=<%=rss("EnSmallClassName")%>"><%=rss("EnSmallClassName")%></a><BR>
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
	rsBigClass.movenext
	i=i+1
	loop
end if
%>
</TABLE>
<%
end Sub 

'=========================================================================
'��������ShowVote
'��  �ã���ʾ��վ����
'��  ������
'=========================================================================
sub ShowVote()
	dim sqlVote,rsVote,i
	sqlVote="select top 1 * from Vote where IsSelected=True"
	Set rsVote= Server.CreateObject("ADODB.Recordset")
	rsVote.open sqlVote,conn,1,1
	if rsVote.bof and rsVote.eof then 
		response.Write "&nbsp;There is not any investigation" 'There is not any investigationû���κε���
	else
		response.write "<form name='VoteForm'method='post'action='En_vote.asp'target='_blank'><td>"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & rsVote("EnTitle") & "<br>"
		if rsVote("VoteType")="Single" then
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='radio'name='VoteOption'value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		else
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='checkbox'name='VoteOption'value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		end if
		response.write "<br><input name='VoteType'type='hidden'value='" & rsVote("VoteType") & "'>"
		response.write "<input name='Action'type='hidden'value='Vote'>"
		response.write "<input name='ID'type='hidden'value='" & rsVote("ID") & "'>"
		response.write "<div align='center'>"
		response.write "<a href='javascript:VoteForm.submit();'><img src='images/voteSubmit.gif' width='52'height='18'border='0'></a>&nbsp;&nbsp;"
        response.write "<a href='En_Vote.asp?Action=Show'target='_blank'><img src='images/voteView.gif' width='52'height='18'border='0'></a>"
		response.write "</div></td></form>"
	end if
	rsVote.close
	set rsVote=nothing
end sub

'=========================================================================
'��������ShowClassGuide
'��  �ã���ʾ��Ŀ����λ��
'��  ������
'=========================================================================
sub ShowClassGuide()
	response.write  "<a href='En_Product.asp'>ProductShow</a>&nbsp;&gt;&gt;" 'ProductShow ��Ʒչʾ
	if EnBigClassName="" and EnSmallClassName="" then
		response.write "&nbsp;All products"
	else
		if EnBigClassName<>"" then
			response.write "&nbsp;<a href='En_Product.asp?EnBigClassName=" & EnBigClassName & "'>" & EnBigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if EnSmallClassName<>"" then
				response.write "<a href='En_Product.asp?EnBigClassName=" & EnBigClassName & "&EnSmallClassName=" & EnSmallClassName & "'>" & EnSmallClassName & "</a>"
			else
				response.write "All small class"
			end if
		end if
	end if
end sub

'=========================================================================
'��������ShowArticleTotal
'��  �ã���ʾ��������
'��  ������
'=========================================================================
sub ShowArticleTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_Product where Passed=True "
	if EnBigClassName<>"" then
		sqlTotal=sqlTotal & " and EnBigClassName='" & EnBigClassName & "'"
		if EnSmallClassName<>"" then
			sqlTotal=sqlTotal & " and EnSmallClassName='" & EnSmallClassName & "'"
		end if
	else
		if SpecialName<>"" then
			sqlTotal=sqlTotal & " and SpecialName='" & SpecialName & "'"
		end if
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "Altogether 0 a product" '����  ����Ʒ
	else
		totalPut=rsTotal(0)
		response.Write "Altogether " & totalPut & " a product" '����  ����Ʒ
	end if
	rsTotal.close
	set rsTotal=nothing
end sub

'=========================================================================
'��������ShowArticle
'=========================================================================
sub ShowArticle(TitleLen)
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
        sqlArticle="select top " & MaxPerPage	
	else
		sqlArticle="select "
	end if

	sqlArticle=sqlArticle & " ArticleID,Product_Id,EnBigClassName,EnSmallClassName,IncludePic,EnTitle,EnSpec,Size,EnMemo,DefaultPicUrl,UpdateTime,Hits from Bs_Product where Passed=True "
	
	if EnBigClassName<>"" then
		sqlArticle=sqlArticle & " and EnBigClassName='" & EnBigClassName & "'"
		if EnSmallClassName<>"" then
			sqlArticle=sqlArticle & " and EnSmallClassName='" & EnSmallClassName & "'"
		end if
	else
		if SpecialName<>"" then
			sqlArticle=sqlArticle & " and SpecialName='" & SpecialName & "'"
		end if
	end if
	sqlArticle=sqlArticle & " order by articleid desc"
	Set rsArticle= Server.CreateObject("ADODB.Recordset")
	rsArticle.open sqlArticle,conn,1,1
	if rsArticle.bof and  rsArticle.eof then
		response.Write("<br><li>There are not any products</li>")
	else
		if currentPage=1 then
			call ArticleContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsArticle.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsArticle.bookmark
            	call ArticleContent(TitleLen)
        	else
	        	currentPage=1
           		call ArticleContent(TitleLen)
	    	end if
		end if
	end if
	rsArticle.close
	set rsArticle=nothing
end sub

sub ArticleContent(intTitleLen)
   	dim i,strTemp
    i=0
	do while not rsArticle.eof
		strTemp=""
		'strTemp = strTemp & ""
		strTemp= strTemp & "<table width=100% border=0 cellspacing=3 cellpadding=0>"
			strTemp= strTemp & "<tr>"
			strTemp= strTemp & "<td width=30% rowspan=6>"
			strTemp= strTemp & "<div align=center><a href=En_ProductShow.asp?ArticleID=" & rsArticle("articleid") & ">" 
			strTemp= strTemp & "<img border=0 src='" & rsArticle("DefaultPicUrl") & "' width=120 height=120>" 
			strTemp= strTemp & "</a></div></td>"
			strTemp= strTemp & "<td width=12% height=12>Name��</td>"
			strTemp= strTemp & "<td><a href=En_ProductShow.asp?ArticleID=" & rsArticle("articleid") & ">" & rsArticle("EnTitle") & "</a></td>"
			strTemp= strTemp & "</tr>"

			strTemp= strTemp & "<tr>"
			strTemp= strTemp & "<td width=12% height=12>Spec��</td>"
			strTemp= strTemp & "<td>" & rsArticle("EnSpec") & "</td>"
			strTemp= strTemp & "</tr>"

			strTemp= strTemp & "<tr>"
			strTemp= strTemp & "<td width=12% height=12>Size��</td>"
			strTemp= strTemp & "<td>" & rsArticle("Size") & "</td>"
			strTemp= strTemp & "</tr>"

			strTemp= strTemp & "<tr>"
			strTemp= strTemp & "<td height=12>class��</td>"
			strTemp= strTemp & "<td><a href='En_Product.asp?EnBigClassName=" & rsArticle("EnBigClassName") & "'>" & rsArticle("EnBigClassName") & "</a> �� "
			strTemp= strTemp & "<a href='En_Product.asp?EnBigClassName=" & rsArticle("EnBigClassName") & "&EnSmallClassName=" & rsArticle("EnSmallClassName") & "'>" & rsArticle("EnSmallClassName") & ""
			strTemp= strTemp & "</a></td>"
			strTemp= strTemp & "</tr>" 

			strTemp= strTemp & "<tr>"
			strTemp= strTemp & "<td height=12>Info��</td>"
			strTemp= strTemp & "<td><a href=En_ProductShow.asp?ArticleID=" & rsArticle("articleid") & "><img src=../Img/En_arrow_7.gif border=0></a></td>"
			strTemp= strTemp & "</tr>"

			strTemp= strTemp & "<tr>"
			strTemp= strTemp & "<td colspan=2>"
			strTemp= strTemp & "<table width=100% border=0 cellpadding=0 cellspacing=0>"
			strTemp= strTemp & "<tr>" 
			strTemp= strTemp & "<td width=50% height=12>"
			strTemp= strTemp & "<div align=center></div></td>"
			strTemp= strTemp & "<td width=50% height=12>"
			strTemp= strTemp & "<div align=center><a href='javascript:eshop(" & rsArticle("Product_Id") & ")'><img border=0 src=../Img/En_addtocart.gif width=95 height=17>"
			strTemp= strTemp & "</a></div></td>"
			strTemp= strTemp & "</tr>"
			strTemp= strTemp & "</table>"
			strTemp= strTemp & "</td>"
			strTemp= strTemp & "</tr>" 

			strTemp= strTemp & "<tr>" 
			strTemp= strTemp & "<td height=1 colspan=3 bgcolor=#CCCCCC></td>"
			strTemp= strTemp & "</tr>"
			strTemp= strTemp & "</table>"
		response.write strTemp
		rsArticle.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 

'=========================================================================
'��������ShowSearchTerm
'��  �ã���ʾ����������Ϣ
'��  ������
'=========================================================================
sub ShowSearchTerm()
	response.write "|&nbsp;ProductSearch&nbsp;&gt;&gt; " 'Product search��Ʒ����
	if EnBigClassName<>"" then
		response.write "<a href='En_Default.asp?EnBigClassName=" & EnBigClassName & "'>" & EnBigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		if EnSmallClassName<>"" then
			response.write "<a href='En_Default.asp?EnBigClassName=" & EnBigClassName & "&EnSmallClassName=" & EnSmallClassName & "'>" & EnSmallClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		else
			response.write "All Small Class &nbsp;&gt;&gt;&nbsp;" 'All Small Class����С��
		end if
	end if
	if keyword<>"" then
	  select case strField
		case "EnTitle"
			response.Write "Contain in the name <font color=red>"&keyword&"</font> Products" '�����к���  �Ĳ�Ʒ
		case "EnContent"
			response.Write "Contain description <font color=red>"&keyword&"</font> Products" '˵������  �Ĳ�Ʒ 
		case else
			response.Write "Contain name <font color=red>"&keyword&"</font> Products" '�����к���  �Ĳ�Ʒ 
	  end select
	else
	  response.write "&nbsp;All products" '���в�Ʒ
	end if
end sub

'=========================================================================
'��������SearchResultTotal
'��  �ã���ʾ�����������
'��  ������
'=========================================================================
sub SearchResultTotal()
	dim rsTotal,sqlTotal
	sqlTotal="select count(*) from Bs_Product where Passed=True "
	if EnBigClassName<>"" then
		sqlTotal=sqlTotal & " and EnBigClassName='" & EnBigClassName & "'"
		if EnSmallClassName<>"" then
			sqlTotal=sqlTotal & " and EnSmallClassName='" & EnSmallClassName & "'"
		end if
	else
		if SpecialName<>"" then
			sqlTotal=sqlTotal & " and SpecialName='" & SpecialName & "'"
		end if
	end if
	if keyword<>"" then
		select case strField
			case "EnTitle"
				sqlTotal=sqlTotal & " and EnTitle like '%" & keyword & "%'"
			case "EnContent"
				sqlTotal=sqlTotal & " and EnContent like '%" & keyword & "%'"
			case else
				sqlTotal=sqlTotal & " and EnTitle like '%" & keyword & "%'"
		end select
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "Altogether 0 a product" '���� 0 ����Ʒ
	else
		totalPut=rsTotal(0)
		response.Write "Find altogether " & totalPut & " a product" '���ҵ�  ����Ʒ
	end if
end sub

'=========================================================================
'��������ShowSearchResult
'��  �ã���ҳ��ʾ�������
'��  ������
'=========================================================================
sub ShowSearchResult()
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
        sqlSearch="select top " & MaxPerPage	
	else
		sqlSearch="select "
	end if

	sqlSearch=sqlSearch & " * from Bs_Product where Passed=True "
	if EnBigClassName<>"" then
		sqlSearch=sqlSearch & " and EnBigClassName='" & EnBigClassName & "'"
		if EnSmallClassName<>"" then
			sqlSearch=sqlSearch & " and EnSmallClassName='" & EnSmallClassName & "'"
		end if
	else
		if SpecialName<>"" then
			sqlSearch=sqlSearch & " and SpecialName='" & SpecialName & "'"
		end if
	end if
	if keyword<>"" then
		select case strField
			case "EnTitle"
				sqlSearch=sqlSearch & " and EnTitle like '%" & keyword & "%'"
			case "EnContent"
				sqlSearch=sqlSearch & " and EnContent like '%" & keyword & "%'"
			case else
				sqlSearch=sqlSearch & " and EnTitle like '%" & keyword & "%'"
		end select
	end if
	sqlSearch=sqlSearch & " order by articleid desc"
	Set rsSearch= Server.CreateObject("ADODB.Recordset")
	rsSearch.open sqlSearch,conn,1,1
 	if rsSearch.eof and rsSearch.bof then 
       		response.write "<p align='center'><br><br>Have not had or found any products</p>" 'û�л�û���ҵ��κβ�Ʒ
   	else 
   		if currentPage=1 then 
       		call SearchResultContent()
   		else 
       		if (currentPage-1)*MaxPerPage<totalPut then 
       			rsSearch.move  (currentPage-1)*MaxPerPage 
       			dim bookmark 
       			bookmark=rsSearch.bookmark 
       			call SearchResultContent()
      		else 
        		currentPage=1 
       			call SearchResultContent()
      		end if 
	   	end if 
   	end if 
   	rsSearch.close 
   	set rsSearch=nothing   
end sub

sub SearchResultContent()
   	dim i,strTemp,Encontent
	i=1
	do while not rsSearch.eof
		strTemp=""
		strTemp=strTemp & cstr(i) & ".<a href='En_ProductShow.asp?ArticleID=" & rsSearch("articleid") & "'>"
		if strField="EnTitle" then
			strTemp=strTemp & "<b>" & replace(rsSearch("EnTitle"),""&keyword&"","<font color=red>"&keyword&"</font>") & "</b></font></a>"
		else
			strTemp=strTemp & "<b>" & rsSearch("EnTitle") & "</b></a>"
		end if
		strTemp=strTemp & " [" & FormatDateTime(rsSearch("UpdateTime"),1) & "]"
		EnContent=left(nohtml(rsSearch("EnContent")),200)
		if strField="EnContent" then
			strTemp=strTemp & "<div style='padding:10px 20px'>" & replace(EnContent,""&keyword&"","<font color=red>"&keyword&"</font>") & "����</div>"
		else
			strTemp=strTemp & "<div style='padding:10px 20px'>" & EnContent & "����</div>"
		end if
		'strTemp=strTemp & "</a>"
		response.write strTemp
		i=i+1
		if i>MaxPerPage then exit do
		rsSearch.movenext
	loop
end sub 


'=========================================================================
'��������ShowSearch
'��  �ã���ʾ����������
'��  ����ShowType ----��ʾ��ʽ��1Ϊ����2Ϊ����
'=========================================================================
sub ShowSearch(ShowType)
	dim count
	if ShowType<>1 and ShowType<>2 then
		ShowType=1
	end if
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
subcat[<%=count%>] = new Array("<%= trim(rs("EnSmallClassName"))%>","<%= trim(rs("EnBigClassName"))%>","<%= trim(rs("EnSmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.EnSmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.EnSmallClassName.options[document.myform.EnSmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    
</script>
<table border="0" cellpadding="2" cellspacing="0" align="center">
<form method="Get" name="myform" action="En_search.asp">
<tr><td height="28">
<select name="Field" size="1">
    <option value="EnTitle" selected>Produc name</option><!-- ��Ʒ���� -->
    <option value="EnContent">Product description</option><!-- ��Ʒ˵�� -->
</select>
<%if ShowType=1 then%>
</td></tr>
<tr><td height="28">
<%end if%>
<select name="EnBigClassName" onChange="changelocation(document.myform.EnBigClassName.options[document.myform.EnBigClassName.selectedIndex].value)" size="1">
	<option selected value="">All big class</option>
<%
if not (rsBigClass.bof and rsBigClass.eof) then
	rsBigClass.movefirst
	do while not rsBigClass.eof
        response.Write "<option value='" & trim(rsBigClass("EnBigClassName")) & "'>" & trim(rsBigClass("EnBigClassName")) & "</option>"
	   	rsBigClass.movenext
	loop
end if
%>
</select>
<%if ShowType=1 then%>
</td></tr>
<tr><td height="28">
<%end if%>
<select name="EnSmallClassName" size="1">                  
    <option selected value="">All small class</option>
</select>
<%if ShowType=1 then%>
</td></tr>
<tr><td height="28">
<%end if%>
<input type="text" name="keyword"  size=16 value="Key word" maxlength="50" onFocus="this.select();">
<input type="submit" name="Submit"  value="Search">
</td></tr>
</form>
</table>
<%
end sub

'=========================================================================
'��������ShowAllClass
'��  �ã���ʾ������Ŀ����Ŀ������
'��  ������
'=========================================================================
sub ShowAllClass()
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "&nbsp;There is not any column" 'There is not any columnû���κ���Ŀ
	else
		dim sqlClass,rsClass,strClassName
		rsBigClass.movefirst
		do while not rsBigClass.eof
			strClassName= "��<a href='En_Product.asp?EnBigClassName=" & rsBigClass("EnBigClassName") & "'><b>" & rsBigClass("EnBigClassName") & "</b></a>��<br><br>"
			sqlClass="select * from Bs_PrSmallClass where EnBigClassName='" & rsBigClass("EnBigClassName") & "'Order by SmallClassID"
			Set rsClass= Server.CreateObject("ADODB.Recordset")
			rsClass.open sqlClass,conn,1,1
			do while not rsClass.eof
				strClassName=strClassName & "&nbsp;<a href='En_Product.asp?EnBigClassName=" & rsClass("EnBigClassName") & "&EnSmallClassName=" & rsClass("EnSmallClassName") & "'>" & rsClass("EnSmallClassName") & "</a>&nbsp;"
				rsClass.movenext
			loop
			response.write strClassName & "<br><br>"
			rsBigClass.movenext
		loop
		rsClass.close
		set rsClass=nothing
	end if
end sub


'=========================================================================
'��������ShowArticleContent
'��  �ã���ʾ���¾�������ݣ����Է�ҳ��ʾ
'��  ������
'"<img border=0 src=" & rsArticle("DefaultPicUrl") & " width=100 height=120>"
'=========================================================================
sub ShowArticleContent()
	dim ArticleID,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ArticleID=rs("ArticleID")
	strContent=rs("EnContent")
	ContentLen=len(strContent)
	CurrentPage=trim(request("ArticlePage"))
	if ShowContentByPage="No" or ContentLen<=200000 then
		response.write strContent
		if ShowContentByPage="Yes" then
			response.write "</p><p align='center'></p>"
		end if
	else
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		pages=ContentLen\MaxPerPage_Content
		if MaxPerPage_Content*pages<ContentLen then
			pages=pages+1
		end if
		lngBound=MaxPerPage_Content          '�����Χ
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		dim lngTemp
		dim lngTemp1,lngTemp1_1,lngTemp1_2,lngTemp1_1_1,lngTemp1_1_2,lngTemp1_1_3,lngTemp1_2_1,lngTemp1_2_2,lngTemp1_2_3
		dim lngTemp2,lngTemp2_1,lngTemp2_2,lngTemp2_1_1,lngTemp2_1_2,lngTemp2_2_1,lngTemp2_2_2
		dim lngTemp3,lngTemp3_1,lngTemp3_2,lngTemp3_1_1,lngTemp3_1_2,lngTemp3_2_1,lngTemp3_2_2
		dim lngTemp4,lngTemp4_1,lngTemp4_2,lngTemp4_1_1,lngTemp4_1_2,lngTemp4_2_1,lngTemp4_2_2
		dim lngTemp5,lngTemp5_1,lngTemp5_2
		dim lngTemp6,lngTemp6_1,lngTemp6_2
		
		if CurrentPage=1 then
			BeginPoint=1
		else
			BeginPoint=MaxPerPage_Content*(CurrentPage-1)+1
			
			lngTemp1_1_1=instr(BeginPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(BeginPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(BeginPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(BeginPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(BeginPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(BeginPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=BeginPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2
				else
					lngTemp1=lngTemp1_1+8
				end if
			end if

			lngTemp2_1_1=instr(BeginPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(BeginPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(BeginPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(BeginPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lntTemp2=BeginPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngtemp2=lngTemp2_2
				else
					lngTemp2=lngTemp2_1+4
				end if
			end if

			lngTemp3_1_1=instr(BeginPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(BeginPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(BeginPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(BeginPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=BeginPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2
				else
					lngTemp3=lngTemp3_1+5
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>BeginPoint and lngTemp<=BeginPoint+lngBound then
				BeginPoint=lngTemp
			else
				lngTemp4_1_1=instr(BeginPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(BeginPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(BeginPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(BeginPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=BeginPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2
					else
						lngTemp4=lngTemp4_1+5
					end if
				end if
				
				if lngTemp4>BeginPoint and lngTemp4<=BeginPoint+lngBound then
					BeginPoint=lngTemp4
				else					
					lngTemp5_1=instr(BeginPoint,strContent,"<img",1)
					lngTemp5_2=instr(BeginPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2
					else
						lngTemp5=BeginPoint
					end if
					
					if lngTemp5>BeginPoint and lngTemp5<BeginPoint+lngBound then
						BeginPoint=lngTemp5
					else
						lngTemp6_1=instr(BeginPoint,strContent,"<br>",1)
						lngTemp6_2=instr(BeginPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2
						else
							lngTemp6=0
						end if
					
						if lngTemp6>BeginPoint and lngTemp6<BeginPoint+lngBound then
							BeginPoint=lngTemp6+4
						end if
					end if
				end if
			end if
		end if

		if CurrentPage=pages then
			EndPoint=ContentLen
		else
		  EndPoint=MaxPerPage_Content*CurrentPage
		  if EndPoint>=ContentLen then
			EndPoint=ContentLen
		  else
			lngTemp1_1_1=instr(EndPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(EndPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(EndPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(EndPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(EndPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(EndPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=EndPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2-1
				else
					lngTemp1=lngTemp1_1+7
				end if
			end if

			lngTemp2_1_1=instr(EndPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(EndPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(EndPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(EndPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=EndPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngTemp2=lngTemp2_2-1
				else
					lngTemp2=lngTemp2_1+3
				end if
			end if

			lngTemp3_1_1=instr(EndPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(EndPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(EndPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(EndPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=EndPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2-1
				else
					lngTemp3=lngTemp3_1+4
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>EndPoint and lngTemp<=EndPoint+lngBound then
				EndPoint=lngTemp
			else
				lngTemp4_1_1=instr(EndPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(EndPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(EndPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(EndPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=EndPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2-1
					else
						lngTemp4=lngTemp4_1+4
					end if
				end if
				
				if lngTemp4>EndPoint and lngTemp4<=EndPoint+lngBound then
					EndPoint=lngTemp4
				else					
					lngTemp5_1=instr(EndPoint,strContent,"<img",1)
					lngTemp5_2=instr(EndPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1-1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2-1
					else
						lngTemp5=EndPoint
					end if
					
					if lngTemp5>EndPoint and lngTemp5<EndPoint+lngBound then
						EndPoint=lngTemp5
					else
						lngTemp6_1=instr(EndPoint,strContent,"<br>",1)
						lngTemp6_2=instr(EndPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1+3
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2+3
						else
							lngTemp6=EndPoint
						end if
					
						if lngTemp6>EndPoint and lngTemp6<EndPoint+lngBound then
							EndPoint=lngTemp6
						end if
					end if
				end if
			end if
		  end if
		end if
		response.write mid(strContent,BeginPoint,EndPoint-BeginPoint)
		
		response.write "</p><p align='center'><b>"
		if CurrentPage>1 then
			response.write "<a href='En_ProductShow.asp?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage-1 & "'>Previous page</a>&nbsp;&nbsp;"
		end if
		for i=1 to pages
			if i=CurrentPage then
				response.write "<font color='red'>[" & cstr(i) & "]</font>&nbsp;"
			else
				response.write "<a href='En_ProductShow.asp?ArticleID=" & ArticleID & "&ArticlePage=" & i & "'>[" & i & "]</a>&nbsp;"
			end if
		next
		if CurrentPage<pages then
			response.write "&nbsp;<a href='En_ProductShow.asp?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage+1 & "'>Next page</a>"
		end if
		response.write "</b></p>"
	end if

end sub
%>
