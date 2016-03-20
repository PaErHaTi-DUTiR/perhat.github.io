<%
'--------版权说明------------------
'web300.cn防注入程序 V2006
'BlackOut站点：http://www.web300.cn 
'Mail:noimpulse@hotmail.com

'------------禁止外部提交------------
dim server_v1,server_v2
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
'If server_v1<>"" Then
'	if  mid(server_v1,8,len(server_v2))<>server_v2  then
'		Response.Redirect server_v2
'	end if
'End if 

'--------定义部份------------------
Dim N_Post,N_Get,N_In,N_Inf,N_Xh,N_db,N_dbstr,alert_info,alert_url,N_type,Sec_Forms,Sec_Form_open,Sec_Form,no_Check
Dim ApplicationValue(7)
ApplicationValue(0)="'|;|and|(|)|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare" '过滤字符
ApplicationValue(1)="http://www.asbd.cn"														'出错后跳转到的地址
ApplicationValue(2)="警告：请不要在参数中包含非法字符尝试注入！\n\n" '警告提示信息
ApplicationValue(3)=4																											'处理方式:1 直接关闭网页 2 警告后关闭 3 跳转到指定页面 4 警告后跳转
ApplicationValue(4)="form1|form2"																					'安全表单
ApplicationValue(5)=0																											'是否启用安全表单：0 不启用 1 启用
ApplicationValue(6)="boss/"																								'不检查的路径,一般填写后台路径

N_In = ApplicationValue(0)
alert_url = ApplicationValue(1)
alert_info = ApplicationValue(2)
N_type = ApplicationValue(3)
Sec_Forms = ApplicationValue(4)
Sec_Form_open = ApplicationValue(5)
no_Check = ApplicationValue(6)
Sec_Form = split(Sec_Forms,"|")
N_Inf = split(N_In,"|")

if instr(lcase(server_v1),lcase(no_Check))=0 then

	'If Request.Form<>"" Then StopInjection(Request.Form)

	'If Request.QueryString<>"" Then StopInjection(Request.QueryString)

	'If Request.Cookies<>"" Then StopInjection(Request.Cookies)
end if

Function N_Alert(alert_info)
	Dim str
	str = "<"&"Script Language=JavaScript"&">"
	Select Case N_type
		Case 1
			str = str & "window.opener=null; window.close();"
		Case 2
			str = str & "alert('"&alert_info&"Http://www.web300.cn\n\nBy:BlackOut');window.opener=null; window.close();"
		Case 3
			str = str & "location.href='"&alert_url&"';"
		Case 4
			str = str & "alert('"&alert_info&"');location.href='"&alert_url&"';"
	end Select
	str = str & "<"&"/Script"&">"
	response.write  str
End Function 

Function alt(str)
		response.write "<"&"Script Language=JavaScript"&">alert('" & str & "');<"&"/Script"&">"
End Function 

Function StopInjection(values)
	For Each N_Get In values
		If values = Request.Form Then
			If Sec_Form_open = 1 Then 
				Security_From(values)
			Else
				Select_BadChar(values)
			End If 
		Else
			Select_BadChar(values)
		End If
	Next
End Function 

Function Select_BadChar(values)
	For N_Xh=0 To Ubound(N_Inf)
		If Instr(LCase(values(N_Get)),N_Inf(N_Xh))<>0 Then
			N_Alert(alert_info)
			Response.End
		End If
	Next
End Function

Function Security_From(values)
	For N_i=0 To UBound(Sec_Form)
		response.write N_Get
		If Instr(LCase(N_Get),Sec_Form(N_i))= 0 Then Select_BadChar(values)
	Next
End Function 
%>
