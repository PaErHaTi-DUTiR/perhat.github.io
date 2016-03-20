<?php
/*
 *后台用户管理控制器 
 */
class UserAction extends CommonAction{
	/* 后台管理员列表视图 */
	Public function index(){
		$db=M('adminuser');
		$list=$db->select();
		$this->assign('list',$list);
		$this->display('admin');
	}
	/* 添加后台管理员 */
	Public function addAdmin(){
		$this->display();
	}
	/* 添加管理员处理 */
	Public function runAddAdmin(){
		if(!$this->isPost()){
			halt('页面不存在');
		}
		if($_POST['pwded']==$_POST['pwd']){
			$data=array(
				'username'=>$this->_post('username'),
				'password'=>$this->_post('pwd','md5'),
				'admin'=>$this->_post('admin','intval'),
				'logintime'=>time(),
				'loginip'=>get_client_ip(),
				'lock'=>0
			);
			$id=M('adminuser')->data($data)->add();
			if($id){
				$this->success('添加成功','index');
			}else{
				$this->error('添加失败');
			}
		}else{
			$this->error('俩次密码不一致');
		}
	
	}
	/**
	 * 锁定后台管理员
	 */
	Public function lockAdmin () {
		$data = array(
			'id' => $this->_get('id', 'intval'),
			'lock' => $this->_get('lock', 'intval')
			);

		$msg = $data['lock'] ? '锁定' : '解锁';
		if (M('adminuser')->save($data)) {
			$this->success($msg . '成功', U('index'));
		} else {
			$this->error($msg . '失败，请重试...');
		}
	}

	/**
	 * 删除后台管理员
	 */
	Public function delAdmin () {
		$id = $this->_get('id', 'intval');

		if (M('adminuser')->delete($id)) {
			$this->success('删除成功', U('index'));
		} else {
			$this->error('删除失败，请重试...');
		}
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

		$db = M('adminuser');
		//验证旧密码
		$where = array('id' => session('id'));
		$old = $db->where($where)->getField('password');
		
		if ($this->_post('old', 'md5') != $old) {
				$this->error('旧密码错误');
		}

		if ($this->_post('pwd') != $this->_post('pwded')) {
			$this->error('两次密码不一致');
		}

		$newPwd = $this->_post('pwd', 'md5');
		$data = array(
			'id' => session('id'),
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