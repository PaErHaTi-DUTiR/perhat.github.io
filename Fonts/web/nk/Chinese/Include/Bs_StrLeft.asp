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
Sub Style_Bulletin()
%>
<% if strSkins <= 200 and strSkins >= 1 then%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� վ �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
 	<TR> 
		<TD vAlign=top><DIV align="center">
			<table width="168" height="120" border="0" cellPadding="0" cellSpacing="0" class="unnamed1">
				<tr>
				  <td align="center">
<marquee scrollamount='1'scrolldelay='170'direction= 'up' width='170'height='118'id="xiaoqing" onmouseover="xiaoqing.stop()" onmouseout="xiaoqing.start()">
            <%=CSJ_Bulletin%>
          </marquee></td>
				</tr>
			</table>
		</DIV></TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE></DIV>
<%
elseif strSkins <= 250 and strSkins >= 201 then
%>
<DIV align="center">
<table border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img border="0" src="../Skin/201/window/zs-glay-top.gif" width="206" height="48" /></td>
    </tr>
    <tr>
      <td background="../Skin/201/window/glay-mid.gif">
        <DIV align="center">
			<table width="168" height="120" border="0" cellPadding="0" cellSpacing="0" class="unnamed1">
				<tr>
				  <td align="center">
<marquee scrollamount='1'scrolldelay='170'direction= 'up' width='170'height='118'id="xiaoqing" onmouseover="xiaoqing.stop()" onmouseout="xiaoqing.start()"><%=CSJ_Bulletin%></marquee>
                  </td>
				</tr>
			</table>
		</DIV>
	  </td>
    </tr>
    <tr>
      <td><img border="0" src="../Skin/201/window/glay-buttom.gif" width="206" height="15" /></td>
    </tr>
</table>
<table cellpadding="0" cellspacing="0" border="0">
    <tr>
      <td height="8"><img src='../img/1x1_pix.gif' width="3" height="1" /></td>
    </tr>
</table>
</DIV>
<%end if%>
<%
End Sub
Sub Style_Login()
if strSkins <= 100 and strSkins >= 1 then
%>
<%
elseif strSkins <= 200 and strSkins >= 101 then
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� Ա �� ½</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD vAlign=top>
			<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
					<td><% call ShowUserLogin() %></td>
				</tr>
			</table>
		</TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
elseif strSkins <= 250 and strSkins >= 201 then
%>
<DIV align="center">
  <% call ShowUserLogin4() %>
  <TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>

</DIV>
<%
end if
End Sub

Sub Style_CoProfile()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� ˾ �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Profile">�� ˾ �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Ceo">�� �� �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Culture">�� ˾ �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Organize">�� ֯ �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Principle">�� �� �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_CoProfile.asp?Action=Contact">�� ϵ �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_News()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� �� �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_News.asp?Action=Co">�� ˾ �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_News.asp?Action=Ye">ҵ �� �� Ѷ</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_News.asp?Action=Pr">�� Ʒ �� ̬</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Product()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� Ʒ �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD vAlign=top><% call ShowSmallClass_Tree() %></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Search()
%>
<DIV align="center">
<%if strSkins <= 200 and strSkins >= 1 then%>
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� Ʒ �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD vAlign=top><% call ShowSearch(1) %></TD>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
<%
elseif strSkins <= 250 and strSkins >= 201 then '�·��==============================================================
%>
<table border="0" cellspacing="0" cellpadding="0">
  
  <tr>
    <td><img border="0" src="../Skin/201/window/search-glay-top.gif" width="206" height="48" /></td>
  </tr>
  <tr>
    <td background="../Skin/201/window/glay-mid.gif"><div align="center">
        <table width="92%" border="0" align="center" cellpadding="2" cellspacing="0">
          <tr>
            <td align="left"><form action="../chinese/Bs_Search.asp" method="post">
                <font color="#666666">�����������ؼ��֣�<br />
                <input name="keyword" type="text" class="form" id="keyword" size="19" />
                <input type="submit" value="�� �� " name="sub" class="form" />
                </font>
                </p>
            </form></td>
          </tr>
      </table></td>
  </tr>
  <tr>
    <td><img border="0" src="../Skin/201/window/glay-buttom.gif" width="206" height="15" /></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" border="0">
  <tr></tr>
  <tr>
    <td height="8"><img src='../img/1x1_pix.gif' width="3" height="1" /></td>
  </tr>
</table>
<% end if%>
</DIV>
<%
End Sub

Sub Style_Honor()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� ˾ �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Honor.asp?Action=Honor">�� ˾ �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Honor.asp?Action=Img">�� ˾ �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Sale()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>Ӫ �� �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Sale.asp?Action=Sale">�� �� �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Sale.asp?Action=Salea">�� �� �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Job()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� �� �� Ƹ</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Job.asp">�� �� �� Ƹ</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Jobs.asp">�� �� �� ��</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Download()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� �� �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR>
		<TD vAlign=top><% call Show_DownSmallClass_Tree() %></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Member()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� Ա �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Faq.asp">����������</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Went.asp">�û���Ϣ����</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_Manage.asp">�޸Ļ�Ա����</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_ManagePwd.asp">�޸Ļ�Ա����</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_E_shop.asp">���ﶩ����ѯ</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_NetBook.asp">վ����������</a></TD>
	</TR>
	<TR><TD colspan=2 class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<TD class='ML_TdImgLI'></TD>
		<TD class='ML_Tdbg'><a href="Bs_UserLogout.asp">�˳���Ա����</a></TD>
	</TR>
	<TR><TD colspan=2 height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Mail()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� �� �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 align="center" class="ML_Titlebg">
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
	<TR> 
		<form method="POST" action="Mail_chk.asp">
			<TD vAlign=top>
				<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
					<tr> 
						<td height="25"> <div align="center"><input name='email'type='text'value="@" size='15' class=form></div></td>
					</tr>
					<tr> 
						<td height="25"><div align="center"> 
							<input type="submit" value="����" name="action">
							&nbsp; 
							<input type="submit" value="�˶�" name="action">
						</div></td>
					</tr>
				</table>
			</TD>
		</form>
	</TR>
	<TR><TD height=5><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub

Sub Style_Fellow()
%>
<DIV align="center">
<TABLE cellPadding=0 cellSpacing=0 class='ML_TitlePt'>
	<TR> 
		<TD class='ML_TdPtbg'><SPAN class=GlowLeft>�� �� �� ��</SPAN></TD>
	</TR>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 class="ML_Titlebg">
	<TR><TD height=2><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<%
dim rs_links,sqltext4,i
set rs_links=server.createobject("adodb.recordset")
sqltext4="select top 10 * from Bs_Links order by id desc"
rs_links.open sqltext4,conn,1,1

i=0
do while not rs_links.eof
%>
				<TR vAlign=center>
					<TD align='center' class='ML_Tdbg'>
					<a href="<%=rs_links("link")%>" title="<%=rs_links("note")%>"target="_blank"><%=rs_links("name")%></a></TD>
				</TR>
				<TR><TD class='TdDotted'><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR>
<%
rs_links.movenext
i=i+1
if i=10 then exit do
loop
rs_links.close 
%>
</TABLE>
<TABLE cellPadding=0 cellSpacing=0 border=0><TR><TR><TD height=8><IMG src='../img/1x1_pix.gif' width=3 height=1></TD></TR></TABLE>
</DIV>
<%
End Sub
Sub Style_Adcolumn()
%>
<SCRIPT LANGUAGE="JavaScript" src="../Skin/<%=strSkins%>/BsImgLogo.js"></SCRIPT>
<%
End Sub
'��ҳ
Sub Left_Index() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
Call Style_Bulletin() '��վ����
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_CoProfile() '��˾���
Call Style_Product() '��Ʒ�б�
Call Style_Mail() '�ʼ��б�
Call Style_Fellow() '��������
End Sub

'��ҳ��
Sub Right_Index() 
Call Style_Login() '��Ա��½
Call Style_Bulletin() '��վ����
Call Style_Search() '��Ʒ����
End Sub

'��������
Sub Right_Download() 
Call Style_Login() '��Ա��½
Call Style_Search() '��Ʒ����
End Sub

'��˾���
Sub Left_CoProfile() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_CoProfile() '��˾���
Call Style_Mail() '�ʼ��б�
Call Style_Fellow() '��������
End Sub

'������Ѷ
Sub Left_News() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_News() '��������
Call Style_Mail() '�ʼ��б�
Call Style_Fellow() '��������
End Sub

'��Ʒչʾ
Sub Left_Product() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Product() '��Ʒ�б�
Call Style_Fellow() '��������
End Sub

'��˾����
Sub Left_Honor() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Honor() '��˾����
Call Style_Mail() '�ʼ��б�
Call Style_Fellow() '��������
End Sub

'Ӫ������
Sub Left_Sale() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Sale() 'Ӫ������
Call Style_Mail() '�ʼ��б�
Call Style_Fellow() '��������
End Sub

'�˲���Ƹ
Sub Left_Job() 
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Job() '�˲���Ƹ
Call Style_Mail() '�ʼ��б�
Call Style_Fellow() '��������
End Sub

'��������
Sub Left_Download() '�����б�
if strSkins <= 200 and strSkins >= 1 then
Call Style_Login() '��Ա��½
elseif strSkins <= 250 and strSkins >= 201 then 
end if
Call Style_Download() '��Ʒ�б�
Call Style_Fellow() '��������
End Sub

'��Ա����
Sub Left_Member() 
Call Style_Member() '��Ա����
Call Style_Mail() '�ʼ��б�
Call Style_Fellow() '��������
End Sub

'call Style_Bulletin()
%>

