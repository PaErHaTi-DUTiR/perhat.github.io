<?php
return array(
	//加载扩展配置
	'LOAD_EXT_CONFIG' => 'system',
	'TMPL_PARSE_STRING'  	=>array(
    '__PUBLIC__' 			=> __ROOT__.'/'.APP_NAME.'/Modules/' . GROUP_NAME .'/Tpl/Public', // 更改默认的__PUBLIC__ 替换规则
	'__UPLOAD__' => '/App/uploads',
	 '__PLAY__' 			=> __ROOT__.'/'.APP_NAME.'/Modules/' . GROUP_NAME .'/Tpl/Public/Play', // 更改默认的__PUBLIC__ 替换规则
	),

);



?>