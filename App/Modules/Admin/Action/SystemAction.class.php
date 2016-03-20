<?php
/**
 * 系统设置控制器
 */
Class SystemAction extends CommonAction {

	/**
	 * 网站设置
	 */
	Public function index () {
		$con=M('system')->select();
		$this->assign('con',$con);
		$this->display();
	}

	/**
	 * 修改网站设置
	 */
	Public function runEdit () {
		$db=M('system');
		//删除原配置
		$db->where(array('id' => 0))->delete();
		
		$config['WEBNAME'] = $_POST['webname'];
		$config['COPY'] = $_POST['copy'];
		$config['REGIS_ON'] = $_POST['regis_on'];
		$config['WEBDESCRIPTION'] = $_POST['description'];
		$config['WEBKEYWORDS'] = $_POST['keywords'];
		$config['ICP'] = $_POST['icp'];
		$config['WEB_ON'] = $_POST['web'];
		$config['TEL'] = $_POST['tel'];
		
		

		if(M('system')->data($config)->add()){
			$this->success('设置成功！');
		
		}else{
			$this->error('设置失败！');
		}
	}

	/**
	 * 关键设置视图
	 */
	Public function filter () {
		$config = include 'qc/Modules/Index/Conf/system.php';
		$this->filter = implode('|', $config['FILTER']);
		
		$this->display();
	}

	/**
	 * 执行修改关键词
	 */
	Public function runEditFilter () {
		$path = './Index/Conf/system.php';
		$config = include $path;
		$config['FILTER'] = explode('|', $_POST['filter']);

		$data = "<?php\r\nreturn " . var_export($config, true) . ";\r\n?>";

		if (file_put_contents($path, $data)) {
			$this->success('修改成功', U('filter'));
		} else {
			$this->error('修改失败， 请修改' . $path . '的写入权限');
		}
	}
}
?>