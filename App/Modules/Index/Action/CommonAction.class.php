<?php
class CommonAction extends Action{
	Public function _initialize(){
		$status=M('system')->getField('WEB_ON');
		$this->system=M('system')->select();
		$telme=M('news')->where(array('type'=>'联系我们'))->order('addtime DESC')->select();
		$str='';
		foreach($telme as $v){
			
			$str=$v;
			$str['content']=htmlspecialchars_decode($str['content']);
		}
		$this->assign('telme',$str);
		
	}


}
?>