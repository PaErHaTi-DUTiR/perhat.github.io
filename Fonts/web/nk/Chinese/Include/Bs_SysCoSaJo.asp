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
'��վ����������
Dim rs_Main,sql_Main,CSJ_Profile,CSJ_Ceo,CSJ_Culture,CSJ_Organize,CSJ_Principle
Dim CSJ_Contact,CSJ_Bulletin,CSJ_Jobs,CSJ_Sale,CSJ_Salea
Set rs_Main= Server.CreateObject("ADODB.Recordset")
sql_Main="select * from Bs_Company"
rs_main.open sql_Main,conn,1,1
CSJ_Profile=ubbcode(rs_Main("Profile")) '��˾���
CSJ_Ceo=ubbcode(rs_Main("Ceo")) '�ܲ��´�
CSJ_Culture=ubbcode(rs_Main("Culture")) '��˾�Ļ�
CSJ_Organize=ubbcode(rs_Main("Organize")) '��֯����
CSJ_Principle=ubbcode(rs_Main("Principle")) '��������
CSJ_Contact=ubbcode(rs_Main("Contact")) '��ϵ����
CSJ_Bulletin=ubbcode(rs_Main("Bulletin")) '��˾����
CSJ_Sale=ubbcode(rs_Main("Sale"))
CSJ_Salea=ubbcode(rs_Main("Salea"))
CSJ_Jobs=ubbcode(rs_Main("Job"))
rs_Main.close
set rs_Main=nothing
%>
