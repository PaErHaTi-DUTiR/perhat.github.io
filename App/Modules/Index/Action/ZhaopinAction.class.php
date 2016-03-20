<?php
/* 前台酒店控制器 */
class ZhaopinAction extends CommonAction{
	/* 模版视图 */
	public function index(){
	   
	   $this->display();
	}
    
    public function booking(){
    	if($this->isPost()){
			if(!empty($_POST)){
				$data=array(
					'type'=>$this->_post('type'),
					'note'=>'备注：'.$this->_post('note'),
					'name'=>$this->_post('title'),
					'contact'=>$this->_post('name'),
					'tel'=>$this->_post('phone'),
					'email'=>$this->_post('email'),
					'p_num'=>$this->_post('num'),
					'addtime'=>time(),
					'booknum'=>date('Ymd',time()).rand('0','99'),
					'bookstatus'=>0,
				
				);
				if(M('book')->add($data)){
					$this->success('预定成功');
				}else{
					$this->error('预定失败');
				}
			}
		
		}
    }

	public function welnut(){
        $db	=	M('production');

		    $where['type'] = '核桃';    
		    $welnut = $db->where($where)->select();
		    $this->assign('welnut',$welnut);

        $this->display();		
	}

	public function atlas(){
		$db	=	M('production');

		    $where['type'] = '艾特莱斯';    
		    $atlas = $db->where($where)->select();
		    $this->assign('atlas',$atlas);

        $this->display();
	}

	public function badam(){
		$db	=	M('production');

		    $where['type'] = '巴达木';    
		    $badam = $db->where($where)->select();
		    $this->assign('badam',$badam);

        $this->display();
	}

	public function carpet(){
		$db	=	M('production');

		    $where['type'] = '地毯';    
		    $carpet = $db->where($where)->select();
		    $this->assign('carpet',$carpet);

        $this->display();
	}

	public function matang(){
		$db	=	M('production');

		    $where['type'] = '玛仁糖';    
		    $matang = $db->where($where)->select();
		    $this->assign('matang',$matang);

        $this->display();
	}

	public function raisins(){
		$db	=	M('production');

		    $where['type'] = '葡萄';    
		    $raisins = $db->where($where)->select();
		    $this->assign('raisins',$raisins);

        $this->display();
	}

	public function reddates(){

		$db	=	M('production');

		    $where['type'] = '红枣';    
		    $reddates = $db->where($where)->select();
		    $this->assign('reddates',$reddates);

        $this->display();
		/*
		$db	=	M('production');
			import('ORG.Util.Page');// 导入分页类
			$count      = $db->count('*');// 查询满足要求的总记录数
			$Page       = new Page($count,6);// 实例化分页类 传入总记录数和每页显示的记录数
			$show       = $Page->show();// 分页显示输出
			// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
			$list = $db->order('addtime DESC')->limit($Page->firstRow.','.$Page->listRows)->select();
			$this->assign('list',$list);// 赋值数据集
			$this->assign('page',$show);// 赋值分页输出      //输出页码
		$this->display();
		*/
	}
}
?>