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
dim strFileName,MaxPerPage
dim totalPut,CurrentPage,TotalPages
dim sqlInfo,rsInfo
dim i
'Action=trim(request("Action"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
%> 

<%
'------------ShowJobTotal------------
sub ShowJobTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Bs_Job"
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		Response.Write "���� 0 ����Ƹ��Ϣ"
	else
		totalPut=rsTotal(0)
		Response.Write "���� " & totalPut & " ����Ƹ��Ϣ"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub
%> 

<%
'------------ShowJobInfo------------
sub ShowJobInfo()
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

	sqlInfo=sqlInfo & " * from Bs_Job Order BY ID desc"
	Set rsInfo= Server.CreateObject("ADODB.Recordset")
	rsInfo.open sqlInfo,conn,1,1

	if rsInfo.bof and  rsInfo.eof then
		Response.Write("<br><li>��û�κ���Ƹ��Ϣ</li>")
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
dim lb
Response.write "<Table width=90% border=0 cellspacing=0 cellpadding=0>"
i=0
do while not rsInfo.eof
	lb=""
	lb=lb & "<TR>" 
	lb=lb & "<TD height=1><br>"
	lb=lb & "<table width='90%'border='0' align='center'cellpadding='3'cellspacing='1'bgcolor='#000000'>"
	lb=lb & "<tr>"
	lb=lb & "<td width='24%'bgcolor='#F3F3DE'> <div align='center'>��Ƹ����</div></td>"
	lb=lb & "<td width='76%'bgcolor='#F8FCF8'>"& rsInfo("Duix") &"</td>"
	lb=lb & "</tr>"
	lb=lb & "<tr>"
	lb=lb & "<td bgcolor='#F3F3DE'> <div align='center'>��Ƹ����</div></td>"
	lb=lb & "<td bgcolor='#F8FCF8'>"& rsInfo("Rens") &"</td>"
	lb=lb & "</tr>"
	lb=lb & "<tr>"
	lb=lb & "<td bgcolor='#F3F3DE'> <div align='center'>�����ص�</div></td>"
	lb=lb & "<td bgcolor='#F8FCF8'>"& rsInfo("did") &"</td>"
	lb=lb & "</tr>"
	lb=lb & "<tr>"
	lb=lb & "<td bgcolor='#F3F3DE'> <div align='center'>���ʴ���</div></td>"
	lb=lb & "<td bgcolor='#F8FCF8'>"& rsInfo("daiy") &"</td>"
	lb=lb & "</tr>"
	lb=lb & "<tr>"
	lb=lb & "<td bgcolor='#F3F3DE'> <div align='center'>����ʱ��</div></td>"
	lb=lb & "<td bgcolor='#F8FCF8'>"& rsInfo("time") &"</td>"
	lb=lb & "</tr>"
	lb=lb & "<tr>"
	lb=lb & "<td bgcolor='#F3F3DE'> <div align='center'>��Ч����</div></td>"
	lb=lb & "<td bgcolor='#F8FCF8'>"& rsInfo("Qix") &"</td>"
	lb=lb & "</tr>"
	lb=lb & "<tr>"
	lb=lb & "<td height='22'bgcolor='#F3F3DE'> <div align='center'>��ƸҪ��</div></td>"
	lb=lb & "<td bgcolor='#F8FCF8'>"& rsInfo("yaoq") &"</td>"
	lb=lb & "</tr>"
	lb=lb & "<tr bgcolor='#FFFFFF'>"
	lb=lb & "<td height='22'colspan='2'><div align='right'><a href='Bs_JobAccept.asp?job="& rsInfo("Duix") &"'>ӦƸ�˸�λ</a></div></td>"
	lb=lb & "</tr>"
	lb=lb & "</table>"
	lb=lb & "</TD>"
	lb=lb & "</TR>"
	Response.write lb
i=i+1 
if i >= MaxPerPage then exit do 
rsInfo.movenext 
loop 
Response.write "</Table>"
End Sub
%>
