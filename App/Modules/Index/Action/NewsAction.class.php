<?php
/* 信息展示 */
class NewsAction extends CommonAction{
	/* 内容显示 */
	Public function index(){
		$db=M('news');

		import('ORG.Util.Page');// 导入分页类
		$where_n['type'] = '公司新闻';
		$count_n      = $db->where($where_n)->count();// 查询满足要求的总记录数
		$Page_n       = new Page($count_n,5);// 实例化分页类 传入总记录数和每页显示的记录数
		$show_n       = $Page_n->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list_n = $db->order('addtime DESC')->limit($Page_n->firstRow.','.$Page_n->listRows)->where($where_n)->select();
		$this->assign('list_n',$list_n);// 赋值数据集
		$this->assign('page_n',$show_n);// 赋值分页输出      //输出页码

        $where_h['type'] = '行业动态';
		$count_h      = $db->where($where_h)->count();// 查询满足要求的总记录数
		$Page_h       = new Page($count_h,5);// 实例化分页类 传入总记录数和每页显示的记录数
		$show_h       = $Page_h->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list_h = $db->order('addtime DESC')->limit($Page_h->firstRow.','.$Page_h->listRows)->where($where_h)->select();
		$this->assign('list_h',$list_h);// 赋值数据集
		$this->assign('page_h',$show_h);// 赋值分页输出      //输出页码

		$where_z['type'] = '政策法规';
		$count_z      = $db->where($where_z)->count();// 查询满足要求的总记录数
		$Page_z       = new Page($count_z,5);// 实例化分页类 传入总记录数和每页显示的记录数
		$show_z       = $Page_z->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list_z = $db->order('addtime DESC')->limit($Page_z->firstRow.','.$Page_z->listRows)->where($where_z)->select();
		$this->assign('list_z',$list_z);// 赋值数据集
		$this->assign('page_z',$show_z);// 赋值分页输出      //输出页码

		$this->display();
	}

	Public function index_show(){
		if($this->isGet()){
			if($_GET['id']){
			$db	=	M('news');
			$data = $db->find($_GET['id']);
			$data['content']=htmlspecialchars_decode($data['content']);
			/* 上一篇文章 */
			$id=$_GET['id']-1;
			$this->prev=$db->find($id);
			/* 下一篇文章 */
			$id=$_GET['id']+1;
			$this->next=$db->find($id);
			/* p($data);die; */
			$this->tuijie= M('news')->where(array('remed'=>'1'))->select();
			$this->assign('list',$data);
			$this->display('index_show');
			}else{
			$this->error('您访问的页面不存在');
			}
		}else{
		$this->error('页面不存在');
		}
	}
}

?>