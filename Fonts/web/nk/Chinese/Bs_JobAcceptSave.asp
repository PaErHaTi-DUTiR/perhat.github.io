<!--#include file="../Inc/Conn.asp"-->

<!--#include file="../Inc/articleCHAR.INC"-->
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
dim rs,sql
If request.form("jobname")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û����ӦƸ��λ���뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("mane")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����������뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("sex")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����Ա��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("birthday")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û����������ڣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("marry")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�������״�����뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("school")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����ҵԺУ���뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("studydegree")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û����ѧ �����뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("specialty")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û����ר ҵ���뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("gradyear")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����ҵʱ�䣬�뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("telephone")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û����� �����뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("email")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û����E-mail���뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("address")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������ϵ��ַ���뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("ability")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û����ˮƽ���������뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("resumes")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������˼������뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_Jobbook"
rs.open sql,conn,1,3

rs.addnew
rs("jobname")=request.form("jobname")
rs("mane")=request.form("mane")
rs("sex")=request.form("sex")
rs("birthday")=request.form("birthday")
rs("marry")=request.form("marry")
rs("school")=request.form("school")
rs("studydegree")=request.form("studydegree")
rs("specialty")=request.form("specialty")
rs("gradyear")=request.form("gradyear")
rs("telephone")=request.form("telephone")
rs("email")=request.form("email")
rs("address")=request.form("address")
rs("ability")=htmlencode2(request.form("ability"))
rs("resumes")=htmlencode2(request.form("resumes"))
rs("time")=now()
rs.update
rs.close
  response.write"<SCRIPT language=JavaScript>alert('ӦƸ�������ύ�ɹ���');"
  response.write"javascript:history.go(-2)</SCRIPT>"
%>