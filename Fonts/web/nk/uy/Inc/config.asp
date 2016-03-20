<%
Const NewsMaxPerPage_Default=20        '新闻每页标题数
Const MaxPerPage_Default=13        '产品每页显示数
Const MaxPerPage_Search=10        '产品搜索页每页显示数
Const MaxBigClassNumPerLine=10        '每行显示产品大类数
Const MaxSmallClassNumPerLine=10        '每行显示产品小类数
Const MaxPerPage_Downt=10        '下载中心每页显示数
Const ShowSmallClassType_Default="Menu"        '首页的子栏目显示样式
Const ShowSmallClassType_Search="Tree"        '文章搜索页的子栏目显示样式
Const ShowSmallClassType_Article="Top"        '文章内容页的子栏目显示样式
Const ShowSmallClassType_Other="Menu"        '其他页面的子栏目显示样式
Const ShowContentByPage="Yes"        '产品内容是否分页显示
Const MaxPerPage_Content=200000        '每页显示大约字符数
Const ShowRunTime="Yes"        '是否显示页面执行时间
Const EnableArticleCheck="Yes"        '是否启用产品审核功能
Const EnableLinkReg="Yes"        '是否开放友情链接申请
Const PopAnnounce="Yes"        '是否弹出公告窗口
Const HitsOfHot=500        '热门文章点击数
Const MailObject="Jmail"        '邮件发送组件

Const EnableUploadProductPic="Yes"        '是否开放产品图片上传
Const DelUploadProductPic="No"        '要自动删除无用的产品图片上传文件请设为"Yes"
Const SaveUpProductPicPath="UploadProductPic"        '存放上传产品图片的目录
Const UpProductPicType="png|gif|jpg|bmp"        '允许的上传产品图片类型
Const MaxProductPicSize=512        '上传产品图片大小限制(500K)更大的文件请使用FTP上传

Const EnableUploadSoft="Yes"        '是否开放软件上传
Const SaveUpSoftPath="UploadSoft"        '存放上传软件的目录
Const UpSoftType="rar|zip|exe|mpg|rm|wav|mid"        '允许的上传软件类型
Const MaxSoftSize=10240        '上传软件大小限制(10M)更大的文件请使用FTP上传

Const EnableUploadSoftPic="Yes"        '是否开放软件图片上传
Const SaveUpSoftPicPath="UploadSoftPic"        '存放上传软件图片的目录
Const UpSoftPicType="jpg|gif|png|bmp"        '允许的上传软件图片类型
Const MaxSoftPicSize=512        '上传软件图片大小限制(500K)更大的文件请使用FTP上传

Const EnableUploadFile="Yes"        '是否开放文件上传
Const SaveUpFilesPath="UploadFiles"        '存放上传文件的目录
Const UpFileType="gif|jpg|bmp|png|swf|doc"        '允许的上传文件类型
Const MaxFileSize=512        '上传文件大小限制(500K)更大的文件请使用FTP上传

Const DelUpFiles="Yes"        '删除文章时是否同时删除文章中的上传文件

Const SessionTimeout=60     'Session会话的保持时间

%>
<%
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃★★★★★★★★★★★ COPYRIGHT ★★★★★★★★★★★ ┃
'┃程序名称：企业网站管理系统Mac3.0  (ASBDcorpweb Mac3.0)  ┃ 
'┃版权所有：WWW.ASBD.CN  BBS.ASBD.CN                      ┃
'┃程序制作：amcen  QQ:125842009  E-mail:ASBD-VIP@163.COM  ┃ 
'┃Copyright 2006-2008 WWW.ASBD.CN - All Rights Reserved.  ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
%>
