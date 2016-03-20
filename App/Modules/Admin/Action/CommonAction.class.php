<?php
/**
*公用控制器
*/
class CommonAction extends Action{
	//自动执行方法
	Public function _initialize(){
		//判断是否登录
		if(!isset($_SESSION['id'])||!isset($_SESSION['username'])){
			redirect(U('Login/index'));
		}
	}
	/* 
	*图片上传
	*/
	Public function uploadPic () {
		if (!$this->isPost()) {
			halt('页面不存在');
		}
		$upload = $this->_upload('Pic', '900,600', '600,400');
		echo json_encode($upload);
	}
    
    Public function uploadNews () {
		if (!$this->isPost()) {
			halt('页面不存在');
		}
		$upload = $this->_upload('News', '330,219', '230,119');
		echo json_encode($upload);
	}
	/**
	 * 图片上传处理
	 * @param  [String] $path   [保存文件夹名称]
	 * @param  [String] $width  [缩略图宽度多个用，号分隔]
	 * @param  [String] $height [缩略图高度多个用，号分隔(要与宽度一一对应)]
	 * @return [Array]         [图片上传信息]
	 */
	Private function _upload ($path, $width, $height) {
		import('ORG.Net.UploadFile');	//引入ThinkPHP文件上传类
		$obj = new UploadFile();	//实例化上传类
		$obj->maxSize = 20000000;	//图片最大上传大小
		$obj->savePath = 'uploads/' . $path . '/';	//图片保存路径
		$obj->saveRule = 'uniqid';	//保存文件名
		$obj->uploadReplace = true;	//覆盖同名文件
		$obj->allowExts = C('UPLOAD_EXTS');	//允许上传文件的后缀名
		$obj->thumb = true;	//生成缩略图
		$obj->thumbMaxWidth = $width;	//缩略图宽度
		$obj->thumbMaxHeight = $height;	//缩略图高度
		$obj->thumbPrefix = 'max_,medium_';	//缩略图后缀名
		$obj->thumbPath = $obj->savePath . date('Y_m') . '/';	//缩略图保存图径
		$obj->thumbRemoveOrigin = true;	//删除原图
		$obj->autoSub = true;	//使用子目录保存文件
		$obj->subType = 'date';	//使用日期为子目录名称
		$obj->dateFormat = 'Y_m';	//使用 年_月 形式

		if (!$obj->upload()) {
			import('ORG.Util.Image');
			Image::water('./uploads/'.$obj[0]['savename'],'./Data/logo.png');
			return array('status' => 0, 'msg' => $obj->getErrorMsg());
		} else {
			$info = $obj->getUploadFileInfo();
			$pic = explode('/', $info[0]['savename']);
			return array(
				'status' => 1,
				'path' => array(
					'max' => $pic[0] . '/max_' . $pic[1],
					'medium' => $pic[0] . '/medium_' . $pic[1]
					)
				);
		}
	}

}/*****CLASS****END*****************/


?>