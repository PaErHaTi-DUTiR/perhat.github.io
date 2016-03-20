<?php
// 本类由系统自动生成，仅供测试用途
class IndexAction extends CommonAction {
    public function index(){
		$this->display();
    }
		/**
	 * 后台信息页
	 */
	Public function copy () {
		$db = M('indexuser');
		$this->user = $db->count();
		
		$xcount=M('book')->where(array('bookstatus'=>'0'))->count();
		$this->assign('xcount',$xcount);
		
		$count=M('book')->count();
		$this->assign('count',$count);
	
		$this->display();
	}
	/* 退出登录 */
	Public function loginOut(){
		//卸载SESSION
		session_unset();
		session_destroy();
		
		//跳转到登录页
		header('Content-Type:text/html;Charset=UTF-8');
		redirect(U('Login/index'),1,'退出成功，正在跳转...');
	}
	
}/***************CLASS****END********************/