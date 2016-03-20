<?php
/*
 *后台用户管理控制器 
 */
class IndexUserAction extends CommonAction{
	/* 前台列表视图 */
	Public function index(){
		$db=M('indexuser');
		$list=$db->select();
		$this->assign('list',$list);
		$this->display('admin');
	}
	/**
	 * 锁定前台
	 */
	Public function lockIndex() {
		$data = array(
			'id' => $this->_get('id', 'intval'),
			'lock' => $this->_get('lock', 'intval')
			);

		$msg = $data['lock'] ? '锁定' : '解锁';
		if (M('indexuser')->save($data)) {
			$this->success($msg . '成功', U('index'));
		} else {
			$this->error($msg . '失败，请重试...');
		}
	}

	/**
	 * 删除前台用户
	 */
	Public function delIndex () {
		$id = $this->_get('id', 'intval');

		if (M('indexuser')->delete($id)) {
			$this->success('删除成功', U('index'));
		} else {
			$this->error('删除失败，请重试...');
		}
	}
	/* 用户检索 */
	Public function sechIndexUser(){
		if(isset($_GET['sech']) && isset($_GET['type'])){
			$where=$_GET['type']? array('email'=>array('LIKE','%'.$this->_get('sech').'%')): array('truename'=>array('LIKE','%'.$this->_get('sech').'%'));
			//p($where);die;
			$indexuser=M('indexuser')->where($where)->select();
			$this->circuit=$indexuser?$indexuser:false;
		}
		$this->display();
	}
	/*
	*修改密码
    */
	Public function editPwd(){
		$this->display();
	}
	Public function runEditPwd () {
		if (!$this->isPost()) {
			halt('页面不存在');
		}

		$db = M('indexuser');
		//验证旧密码
		$where = array('id' => session('uid'));
		$old = $db->where($where)->getField('password');
		
		if ($this->_post('old', 'md5') != $old) {
				$this->error('旧密码错误');
		}

		if ($this->_post('new') != $this->_post('newed')) {
			$this->error('两次密码不一致');
		}

		$newPwd = $this->_post('new', 'md5');
		$data = array(
			'id' => session('uid'),
			'password' => $newPwd
			);

		if ($db->save($data)) {
			$this->success('修改成功', U('index'));
		} else {
			$this->error('修改失败，请重试...');
		}
	}
}//======CLASS===END


?>