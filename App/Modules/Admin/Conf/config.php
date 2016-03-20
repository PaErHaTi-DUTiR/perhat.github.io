<?php
return array(
	'TMPL_PARSE_STRING'  	=>array(
    '__PUBLIC__' 			=> __ROOT__.'/'.APP_NAME.'/Modules/' . GROUP_NAME .'/Tpl/Public', // 更改默认的__PUBLIC__ 替换规则
	'__UPLOAD__' => '/App/uploads',
	),
	'TMPL_L_DELIM'			=>	'{',			//设置左定界符
	'TMPL_R_DELIM'			=>	'}',			//设置右定界符	
	
);
?>