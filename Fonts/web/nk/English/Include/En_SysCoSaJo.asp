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
'打开站点设置连接
Dim rs_Main,sql_Main,CSJ_Profile,CSJ_Ceo,CSJ_Culture,CSJ_Organize,CSJ_Principle
Dim CSJ_Contact,CSJ_Bulletin,CSJ_Sale,CSJ_Salea,CSJ_Jobs

Set rs_Main= Server.CreateObject("ADODB.Recordset")
sql_Main="select * from Bs_Company"
rs_main.open sql_Main,conn,1,1

CSJ_Profile=ubbcode(rs_Main("EnProfile")) '公司简介
CSJ_Ceo=ubbcode(rs_Main("EnCeo")) '总裁致辞
CSJ_Culture=ubbcode(rs_Main("EnCulture")) '公司文化
CSJ_Organize=ubbcode(rs_Main("EnOrganize")) '组织机构
CSJ_Principle=ubbcode(rs_Main("EnPrinciple")) '精神理念
CSJ_Contact=ubbcode(rs_Main("EnContact")) '联系我们
CSJ_Bulletin=ubbcode(rs_Main("EnBulletin")) '公司公告
CSJ_Sale=ubbcode(rs_Main("EnSale"))
CSJ_Salea=ubbcode(rs_Main("EnSalea"))
CSJ_Jobs=ubbcode(rs_Main("EnJob"))
rs_Main.close
set rs_Main=nothing
%>
