<?php
/* 前台用户中心控制器 */
class IndexUserAction extends CommonAction{
	/*会员中心首页*/
	Public function index(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		$db	=	M('book');
			/* 未完成的订单 */
			$w=array('buystatus'=>'0','contact'=>$_COOKIE['user']);
			/* 完成的订单 */
			$where=array('buystatus'=>1,'contact'=>$_COOKIE['user']);
			import('ORG.Util.Page');// 导入分页类
			$count      = $db->where($w)->count('*');// 查询满足要求的总记录数
			$Page       = new Page($count,4);// 实例化分页类 传入总记录数和每页显示的记录数
			$overcou      = $db->where($where)->count('*');// 查询满足要求的总记录数
			$Pages       = new Page($overcou,4);// 实例化分页类 传入总记录数和每页显示的记录数
			$show       = $Page->show();// 分页显示输出
			$shows      = $Pages->show();// 分页显示输出
			// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
			$noover = $db->where($w)->order('addtime DESC')->limit($Page->firstRow.','.$Page->listRows)->select();
			$over = $db->order('addtime DESC')->where($where)->limit($Pages->firstRow.','.$Pages->listRows)->select();
			$this->assign('page',$show);// 赋值分页输出      //输出页码 
			$this->assign('pages',$shows);// 赋值分页输出      //输出页码 
			$this->assign('noover',$noover);// 赋值数据集
			$this->assign('over',$over);// 赋值数据集
		$this->display('huiyuan');
	}
	/* 分类订单 */
	Public function libBook(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		if(!$this->isGet()){
			$this->error('页面不存在');
		}
			$str='';
			switch($_GET['fenlei']){
				case 1:$str='酒店订单';break;
				case 2:$str='线路订单';break;
				case 3:$str='租车订单';break;
				case 4:$str='佛事订单';break;
				case 5:$str='会议订单';break;
				case 6:$str='商城订单';break;
			}
			$db	=	M('book');
			/* 未完成的订单 */
			$w=array('type'=>$str,'contact'=>$_COOKIE['user']);
			import('ORG.Util.Page');// 导入分页类
			$count      = $db->where($w)->count('*');// 查询满足要求的总记录数
			$Page       = new Page($count,9);// 实例化分页类 传入总记录数和每页显示的记录数
			$show       = $Page->show();// 分页显示输出
			// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
			$list = $db->where($w)->order('addtime DESC')->limit($Page->firstRow.','.$Page->listRows)->select();
			$this->assign('page',$show);// 赋值分页输出      //输出页码 
			$this->assign('list',$list);// 赋值数据集
			$this->display('huiyuan_dingdan');
	
	}
	/* 订单修改 */
	Public function saveBook(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		if(!$this->isGet()){
			$this->error('页面不存在');
		}
		$id=$_GET['id'];
		$list=M('book')->where(array('id'=>$id))->select();
		$this->assign('list',$list);
		$this->display();
	}
	Public function runSaveBook(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		if(!$this->isPost()){
			$this->error('页面不存在');
		}
		$data=array(
			'number'=>$this->_post('number'),
			'contact'=>$this->_post('contact'),
			'tel'=>$this->_post('tel'),
			'action'=>$this->_post('action'),
			'status'=>1,
		);
		$data['allprice']=$_POST['price']*$_POST['number'];
		if(M('book')->where(array('booknum'=>$this->_post('booknum')))->save($data)){
			$this->success('修改成功');
		}else{
			$this->error('修改失败');
		}
	}
	/* 删除订单 */
	Public function delBook(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		if(!$this->isGet()){
			$this->error('页面不存在');
		}
		if(M('book')->where(array('id'=>$_GET['id']))->limit('1')->delete()){
			$this->success('删除成功');
		}else{
			$this->error('删除失败');
		}
	}
	/* 基本资料 */
	Public function ziLiao(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		$db=M('indexuser');
		$w=array('truename'=>$_COOKIE['user']);
		$list=$db->where($w)->select();
		$this->assign('list',$list);
		$this->display('huiyuan_ziliao');
	}
	/* 资料修改 */
	Public function runZiLiao(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		if(!$this->isPost()){
			$this->error('页面不存在');
		}
		$data=array(
			'phone'=>$this->_post('phone'),
			'truename'=>$this->_post('truename'),
			'sex'=>$this->_post('sex'),
			'QQ'=>$this->_post('QQ'),
			'location'=>$this->_post('dizhi'),
		);
		if(M('indexuser')->where(array('truename'=>$_COOKIE['user']))->save($data)){
			$this->success('修改成功');
		}else{
			$this->error('修改失败');
		}
	}
	/* 修改密码 */
	Public function editPwd(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		$this->display('huiyuan_pwd');
	}
	Public function runEditPwd(){
		if(!isset($_COOKIE['user'])){
			redirect(U('Login/index'),2,'<p style="font-size:20px; color:blue; padding-top:200px; padding-left:510px;">aOh! 您还未登录，请先登录！</p>');
		}
		if(!$this->isPost()){
			$this->error('页面不存在');
		}
		/*p($old);die; */
		$old=$this->_post('oldpassword',md5);
		$new=$this->_post('newpassword',md5);
		$repwd=$this->_post('repassword',md5);
		$oldpwd=M('indexuser')->where(array('truename'=>$_COOKIE['user']))->getField('password');
		if($old!=$oldpwd){
			$this->error('旧密码不正确');
		}elseif($new!=$repwd){
			$this->error('重复密码不正确');
		}else{
			if(M('indexuser')->where(array('truename'=>$_COOKIE['user']))->save(array('password'=>$new))){
				cookie('user',null);
				cookie('phone',null);
				cookie('dizhi',null);
				cookie('dzyx',null);
				$this->success('修改成功','__APP__');
			}else{
				$this->error('修改失败');
			}
		}
	}
}
?>