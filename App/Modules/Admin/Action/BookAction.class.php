<?php
/*信息管理控制器 */
class BookAction extends CommonAction{
	/* 信息列表视图 */
	Public function index(){
		$db=M('book');
		$list=$db->order('addtime','DESC')->select();
		import('ORG.Util.Page');// 导入分页类
		$count      = $db->count('*');// 查询满足要求的总记录数
		$Page       = new Page($count,10);// 实例化分页类 传入总记录数和每页显示的记录数
		$show       = $Page->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list = $db->order('addtime')->limit($Page->firstRow.','.$Page->listRows)->select();
		$this->assign('list',$list);// 赋值数据集
		$this->assign('page',$show);// 赋值分页输出      //输出页码
		
		$this->display();
	}
	/* 新增信息 */
	Public function addBook(){
		$this->display();
	}
	/* 添加信息表单接受处理 */
	Public function runBook(){
		if(!$this->isPost()){
			halt('页面不存在');
		} 
			$data=array(
				'bookstatus'=>$this->_post('bookstatus')
			);
			if($hid=M('book')->data($data)->add()){ 
				
				$this->success('发布成功','index');	
			}else{
				$this->error('发布失败');
			}
		
	}
	/* 信息编辑 */
	Public function saveBook(){
		if(!$this->isGet()){
			halt('页面不存在');
		}
		$id=$_GET['id'];
		$w=array('id'=>$id);
		$db=M('book');
		$book=$db->where($w)->select();
		/* p($book);die; */
		$this->assign('book',$book);
		$this->display();
	}
	/* 信息编辑接受表单 */
	Public function sendBook(){
		if(!$this->isPost()){
			halt('页面不存在');
		}
		if(empty($_POST)){
			$this->error('请插入数据再提交');
		}else{
		/* p($_POST);die; */
		$data=array('bookstatus'=>$this->_post('bookstatus'));
		$w=array('id'=>$this->_post('id'));
		$db=M('book');
		if($db->where($w)->save($data)){
			$this->success('修改成功',U('index'));
		}else{
			$this->error('修改失败');
		}
		}
	}
	/* 信息删除 */
	Public function delBook(){
		$w=array('id'=>$_GET['id']);
		if(M('book')->where($w)->limit('1')->delete()){
			$this->success('删除成功',U('index'));
		}else{
			$this->error('删除失败');
		}
	}
	/* 信息检索 */
	Public function sechBook(){
		if(isset($_GET['sech']) && isset($_GET['booknum'])){
			$where=$_GET['booknum']? array('title'=>array('LIKE','%'.$this->_get('sech').'%')): array('title'=>array('LIKE','%'.$this->_get('sech').'%'));
			/* p($where);die; */
			$book=M('book')->where($where)->select();
			$this->book=$book?$book:false;
		}
		$this->display();
	} 
	
	//编辑器上传图片处理
	public function upload(){
		import('ORG.Net.UploadFile');
		$upload= new UploadFile();
        $upload-> autoSub    =  true;// 启用子目录保存文件
        $upload-> subType    =  'date';// 子目录创建方式 可以使用hash date custom
        $upload-> dateFormat =  'Ymd';
		//$upload->savePath = './uploads/';	//图片保存路径
		$upload->thumb = true;	//生成缩略图
		$upload->thumbMaxWidth = '330,220';	//缩略图宽度
		$upload->thumbMaxHeight = '219,119';	//缩略图高度
		$upload->thumbPrefix = 'max_';	//缩略图后缀名
		$upload->thumbPath = './uploads/' . date('Y_m') . '/';	//缩略图保存图径
        if($upload->upload('./uploads/')){
			$info=$upload->getUploadFileInfo();
			import('ORG.Util.Image');
			Image::water('./uploads/'.$info[0]['savename'],'./Data/logo.png');
			echo json_encode(array(
				'url'=>$info[0]['savename'],
				'title'=>htmlspecialchars($_POST['Booktitle'], ENT_QUOTES),
				'bookstatus'=>htmlspecialchars($_POST['bookstatus'], ENT_QUOTES),
				'original'=>$info[0]['name'],
				'state'=>'SUCCESS',
				'path' => array(
					'max' => $Book[0] . '/max_' . $Book[1]
					)
			));
		}else{
			echo json_encode(array(
				'state'=>$upload->getErrorMsg()
			));
		
		}

	}

}/*********CLASS****END***********************/


?>